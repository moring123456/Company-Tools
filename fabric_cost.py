import re
from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Optional
import pandas as pd
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class FabricCostCalculator:
    """布料费用计算器"""

    def __init__(self, excel_file):
        # 接收 Streamlit 传过来的文件对象
        self.excel_file = excel_file
        self.data = {}

    def load_data(self):
        logger.info("正在读取文件并加载工作表...")
        required_sheets = ['布料跟踪表', '订单表', 'SKU-款号映射', '款号信息表']
        try:
            for sheet in required_sheets:
                self.data[sheet] = pd.read_excel(self.excel_file, sheet_name=sheet)
            logger.info("数据加载成功")
        except Exception as e:
            raise Exception(f"加载 Excel 文件失败，请确认工作表是否完整: {str(e)}")

    def normalize_color_for_matching(self, color_str: str) -> str:
        if pd.isna(color_str):
            return ''
        color_str = str(color_str).strip()
        prefixes_to_remove = ['40STR汗布', '40STR', '32STR汗布', '32STR']
        for prefix in prefixes_to_remove:
            if color_str.startswith(prefix):
                color_str = color_str[len(prefix):].strip()
                break
        if not color_str:
            return color_str
        if '印' not in color_str and 'YH' not in color_str:
            if not color_str.startswith('#'):
                color_str = '#' + color_str
        return color_str

    def extract_color_name(self, color_str: str) -> Optional[str]:
        if pd.isna(color_str):
            return None
        color_str = str(color_str).strip()
        for prefix in ['40STR汗布', '40STR', '32STR汗布', '32STR']:
            if color_str.startswith(prefix):
                color_str = color_str[len(prefix):].strip()
                break
        match = re.search(r'#?\d+\s+(.+)', color_str)
        if match:
            return match.group(1).strip()
        match2 = re.search(r'^[#\d]+\s*(.+)', color_str)
        if match2:
            return match2.group(1).strip()
        return color_str if color_str else None

    def extract_color_number(self, color_str: str) -> Optional[str]:
        if pd.isna(color_str):
            return None
        color_str = str(color_str).strip()
        for prefix in ['40STR汗布', '40STR', '32STR汗布', '32STR']:
            if color_str.startswith(prefix):
                color_str = color_str[len(prefix):].strip()
                break
        match = re.search(r'#(\w+)', color_str)
        if match:
            return '#' + match.group(1)
        return None

    def extract_print_code(self, color_str: str) -> Optional[str]:
        if pd.isna(color_str):
            return None
        color_str = str(color_str).strip()
        match = re.search(r'(YH\w+)', color_str)
        if match:
            return match.group(1)
        return None

    def extract_style_from_sku(self, sku: str) -> Optional[str]:
        if pd.isna(sku):
            return None
        sku = str(sku).strip().upper()
        match = re.match(r'^(ADM\d+)', sku)
        if match:
            return match.group(1)
        return None

    def extract_style_from_product_name(self, product_name: str) -> Optional[str]:
        if pd.isna(product_name):
            return None
        product_name = str(product_name).strip().upper()
        match = re.match(r'^(ADM\d+)', product_name)
        if match:
            return match.group(1)
        return None

    def is_solid(self, product_name: str) -> bool:
        if pd.isna(product_name):
            return False
        product_name = str(product_name).upper()
        if '印' in product_name and '印字' not in product_name:
            return False
        return True

    def is_print(self, product_name: str) -> bool:
        if pd.isna(product_name):
            return False
        product_name = str(product_name).upper()
        if '印' in product_name and '印字' not in product_name:
            return True
        return False

    def build_df0_prepare(self, cutoff_date: datetime = None) -> pd.DataFrame:
        if cutoff_date is None:
            cutoff_date = datetime.now()
        six_months_ago = cutoff_date - timedelta(days=180)

        order_df = self.data['订单表']
        valid_order_nos = set(order_df['订单编号'].dropna().unique())

        fabric_df = self.data['布料跟踪表'].copy()
        fabric_df['金额（元）'] = pd.to_numeric(fabric_df['金额（元）'], errors='coerce')
        fabric_df['下单日期'] = pd.to_datetime(fabric_df['下单日期'], errors='coerce')

        df0 = fabric_df[
            (~fabric_df['订单编号'].isin(valid_order_nos)) &
            (fabric_df['金额（元）'] > 0) &
            (fabric_df['下单日期'] >= six_months_ago)
            ].copy()

        df0['备布标记'] = '备布'
        return df0

    def build_df1(self) -> pd.DataFrame:
        df1 = self.data['布料跟踪表'].copy()
        df1['下单数量'] = 0.0
        df1['已收货数量'] = 0.0
        df1['款号'] = None
        df1['颜色匹配状态'] = '未匹配'

        order_df = self.data['订单表'].copy()
        sku_mapping = self.data['SKU-款号映射']

        order_df['颜色名称'] = order_df['颜色'].apply(self.extract_color_name)
        order_df['颜色编号'] = order_df['颜色'].apply(self.extract_color_number)
        order_df['印花色号'] = order_df['颜色'].apply(self.extract_print_code)

        for idx, row in df1.iterrows():
            order_no = row['订单编号']
            product_name = row['产品名称']
            color_code = row['颜色&色号']
            normalized_color = self.normalize_color_for_matching(color_code)

            if self.is_solid(product_name):
                color_name = self.extract_color_name(color_code)
                matched_orders = order_df[(order_df['订单编号'] == order_no) & (order_df['颜色名称'] == color_name)]
                if len(matched_orders) == 0:
                    color_number = self.extract_color_number(color_code)
                    if color_number:
                        matched_orders = order_df[
                            (order_df['订单编号'] == order_no) & (order_df['颜色编号'] == color_number)]
            elif self.is_print(product_name):
                print_code = self.extract_print_code(color_code)
                matched_orders = order_df[(order_df['订单编号'] == order_no) & (order_df['印花色号'] == print_code)]
            else:
                matched_orders = order_df[(order_df['订单编号'] == order_no) & (order_df['颜色'] == normalized_color)]

            if len(matched_orders) > 0:
                df1.at[idx, '下单数量'] = matched_orders['下单数量'].sum()
                df1.at[idx, '已收货数量'] = matched_orders['已收货数量'].sum()
                df1.at[idx, '颜色匹配状态'] = '匹配成功'

                sku = matched_orders['SKU编码'].iloc[0]
                style_no = self.extract_style_from_sku(sku)
                if style_no:
                    df1.at[idx, '款号'] = style_no
                else:
                    style_match = sku_mapping[sku_mapping['SKU'] == sku]['款号']
                    if len(style_match) > 0 and pd.notna(style_match.iloc[0]):
                        df1.at[idx, '款号'] = style_match.iloc[0]
            else:
                df1.at[idx, '颜色匹配状态'] = '匹配失败'

        df1_matched = df1[df1['颜色匹配状态'] == '匹配成功'].copy()
        return df1_matched

    def build_df1_1(self, df1: pd.DataFrame) -> pd.DataFrame:
        df1['金额（元）'] = pd.to_numeric(df1['金额（元）'], errors='coerce')
        df1_1 = df1[df1['金额（元）'].notna()].copy()
        df1_1 = df1_1[df1_1['金额（元）'] > 0]
        return df1_1

    def build_df1_2(self, df1_1: pd.DataFrame) -> pd.DataFrame:
        df1_1['已收货数量'] = pd.to_numeric(df1_1['已收货数量'], errors='coerce')
        df1_1['下单数量'] = pd.to_numeric(df1_1['下单数量'], errors='coerce')
        df1_2 = df1_1[df1_1['已收货数量'].notna()].copy()
        df1_2 = df1_2[df1_2['已收货数量'] > 0]
        df1_2 = df1_2[df1_2['已收货数量'] >= df1_2['下单数量'] * 0.70]
        df1_2 = df1_2[df1_2['款号'].notna()]
        return df1_2

    def get_valid_date_range(self, style_df: pd.DataFrame) -> Tuple[datetime, datetime, List[str]]:
        if len(style_df) == 0:
            return None, None, []
        style_df = style_df.sort_values('下单日期', ascending=False)
        all_orders = style_df[['下单日期', '订单编号']].drop_duplicates().sort_values('下单日期', ascending=False)
        if len(all_orders) == 0:
            return None, None, []

        end_date = all_orders['下单日期'].iloc[0]
        start_date = all_orders['下单日期'].iloc[-1]
        months_diff = (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month)

        if months_diff > 6:
            cutoff_date = end_date - timedelta(days=180)
            recent_orders = all_orders[all_orders['下单日期'] >= cutoff_date]
            if len(recent_orders) > 0:
                start_date = recent_orders['下单日期'].iloc[-1]
                end_date = recent_orders['下单日期'].iloc[0]
            else:
                start_date = all_orders['下单日期'].iloc[-1]
                end_date = all_orders['下单日期'].iloc[0]

        valid_orders = all_orders[(all_orders['下单日期'] >= start_date) & (all_orders['下单日期'] <= end_date)][
            '订单编号'].tolist()
        return start_date, end_date, valid_orders

    def build_df2(self, df1_2: pd.DataFrame, style_no: str) -> pd.DataFrame:
        style_df = df1_2[df1_2['款号'] == style_no].copy()
        if len(style_df) == 0:
            return pd.DataFrame()

        start_date, end_date, valid_orders = self.get_valid_date_range(style_df)
        if not valid_orders:
            return pd.DataFrame()

        style_df = style_df[style_df['订单编号'].isin(valid_orders)]
        grouped = style_df.groupby(by=['产品名称', '颜色&色号', '订单编号'], dropna=False).agg({
            '布料送货数量（匹）': 'sum',
            '金额（元）': 'sum',
            '下单数量': 'sum',
            '已收货数量': 'sum'
        }).reset_index()

        for idx, row in grouped.iterrows():
            matched_rows = style_df[
                (style_df['产品名称'] == row['产品名称']) &
                (style_df['颜色&色号'] == row['颜色&色号']) &
                (style_df['订单编号'] == row['订单编号'])
                ]
            if len(matched_rows) > 0:
                grouped.at[idx, '下单数量'] = matched_rows['下单数量'].iloc[0]
                grouped.at[idx, '已收货数量'] = matched_rows['已收货数量'].iloc[0]

        return grouped

    def calculate_cost(self, df2: pd.DataFrame, q_override: float = None) -> Dict:
        if len(df2) == 0:
            return {'K1': 0.0, 'K1-s': 0.0, 'K1-p': 0.0, 'Q1': 0.0}

        C1 = df2['金额（元）'].sum()
        N1 = df2['布料送货数量（匹）'].sum()
        M1 = df2['已收货数量'].sum()

        Q1 = q_override if q_override is not None else (M1 / N1 if N1 > 0 else 0.0)
        K1 = C1 / N1 / Q1 if N1 > 0 and Q1 > 0 else 0.0

        solid_mask = df2['产品名称'].apply(self.is_solid)
        df2_s = df2[solid_mask]
        if len(df2_s) > 0:
            C1_s = df2_s['金额（元）'].sum()
            N1_s = df2_s['布料送货数量（匹）'].sum()
            M1_s = df2_s['已收货数量'].sum()
            Q1_s = M1_s / N1_s if N1_s > 0 else 0.0
            K1_s = C1_s / N1_s / Q1_s if N1_s > 0 and Q1_s > 0 else 0.0
        else:
            K1_s = 0.0

        print_mask = df2['产品名称'].apply(self.is_print)
        df2_p = df2[print_mask]
        if len(df2_p) > 0:
            C1_p = df2_p['金额（元）'].sum()
            N1_p = df2_p['布料送货数量（匹）'].sum()
            M1_p = df2_p['已收货数量'].sum()
            Q1_p = M1_p / N1_p if N1_p > 0 else 0.0
            K1_p = C1_p / N1_p / Q1_p if N1_p > 0 and Q1_p > 0 else 0.0
        else:
            K1_p = 0.0

        return {'K1': round(K1, 2), 'K1-s': round(K1_s, 2), 'K1-p': round(K1_p, 2), 'Q1': round(Q1, 0)}

    def process_backup_fabric(self, df0: pd.DataFrame, q_values: Dict[str, float]) -> pd.DataFrame:
        if len(df0) == 0:
            return pd.DataFrame()

        df0['成衣数Q'] = 0.0
        df0['已收货数量'] = 0.0

        for idx, row in df0.iterrows():
            style_no = self.extract_style_from_product_name(row['产品名称'])
            q = q_values.get(style_no, 0.0) if style_no else 0.0
            fabric_qty = row['布料送货数量（匹）'] if pd.notna(row['布料送货数量（匹）']) else 0.0
            df0.at[idx, '成衣数Q'] = q
            df0.at[idx, '已收货数量'] = q * fabric_qty

        return df0

    def get_valid_style_numbers(self) -> List[str]:
        style_info = self.data['款号信息表']
        return style_info[style_info['产品定位'] != '淘汰款']['款号'].tolist()

    def calculate(self) -> Tuple[List[Dict], pd.DataFrame]:
        self.load_data()
        df0_raw = self.build_df0_prepare()
        df1 = self.build_df1()
        df1_1 = self.build_df1_1(df1)
        df1_2 = self.build_df1_2(df1_1)

        target_styles = self.get_valid_style_numbers()
        q_values = {}

        # 第一轮计算 Q1
        for style_no in target_styles:
            df2 = self.build_df2(df1_2, style_no)
            if len(df2) == 0:
                continue
            cost_results = self.calculate_cost(df2)
            q_values[style_no] = cost_results['Q1']

        # 处理备布
        df0 = self.process_backup_fabric(df0_raw.copy(), q_values)

        # 第二轮计算
        results = []
        for style_no in target_styles:
            df2 = self.build_df2(df1_2, style_no)

            if len(df2) == 0:
                style_backup = df0[
                    df0.apply(lambda x: self.extract_style_from_sku(str(x['订单编号'])) == style_no, axis=1)]
                if len(style_backup) > 0:
                    q = q_values.get(style_no, 0)
                    backup_amount = style_backup['金额（元）'].sum()
                    backup_fabric_qty = style_backup['布料送货数量（匹）'].sum()
                    if q > 0 and backup_fabric_qty > 0:
                        K1 = backup_amount / backup_fabric_qty / q
                        results.append({
                            '款号': style_no, '成衣数Q1': q, '布料费用K1': round(K1, 2),
                            '纯色布料费用K1-s': 0.0, '印花布料费用K1-p': 0.0
                        })
                continue

            style_backup = df0[
                df0.apply(lambda x: self.extract_style_from_product_name(x['产品名称']) == style_no, axis=1)]
            if len(style_backup) > 0:
                backup_records = []
                for _, backup_row in style_backup.iterrows():
                    backup_records.append({
                        '产品名称': backup_row['产品名称'], '颜色&色号': backup_row['颜色&色号'],
                        '订单编号': backup_row['订单编号'] + '_备布',
                        '布料送货数量（匹）': backup_row['布料送货数量（匹）'],
                        '金额（元）': backup_row['金额（元）'], '下单数量': 0, '已收货数量': backup_row['已收货数量']
                    })
                df2_with_backup = pd.concat([df2, pd.DataFrame(backup_records)],
                                            ignore_index=True) if backup_records else df2
            else:
                df2_with_backup = df2

            cost_results = self.calculate_cost(df2_with_backup, q_override=q_values.get(style_no, 0))
            results.append({
                '款号': style_no, '成衣数Q1': q_values.get(style_no, 0),
                '布料费用K1': cost_results['K1'], '纯色布料费用K1-s': cost_results['K1-s'],
                '印花布料费用K1-p': cost_results['K1-p']
            })

        return results, df0


# 暴露给 Streamlit 调用的统一接口
def run_fabric_calculation(uploaded_file):
    calculator = FabricCostCalculator(uploaded_file)
    results, df0 = calculator.calculate()

    # 格式化布料费用结果
    df_results = pd.DataFrame(results)
    if not df_results.empty:
        df_results = df_results[['款号', '成衣数Q1', '布料费用K1', '纯色布料费用K1-s', '印花布料费用K1-p']]

    # 格式化备布数据提取指定列
    output_columns = ['订单编号', '产品名称', '颜色&色号', '布料送货数量（匹）', '金额（元）', '成衣数Q', '已收货数量',
                      '备布标记']
    if not df0.empty:
        df0_output = df0[[col for col in output_columns if col in df0.columns]]
    else:
        df0_output = pd.DataFrame(columns=output_columns)

    return {
        "布料费用结果": df_results,
        "备布数据": df0_output
    }