import pandas as pd
import re
from datetime import datetime, timedelta
import logging

# 设置基础的日志记录，替代 Coze 的 logger
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class FabricCostCalculator:
    """布料费用计算器"""

    # 注意：这里的 excel_path 现在可以接收 Streamlit 上传的文件对象
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.data = {}
        self.logger = logger

    def load_data(self):
        self.logger.info("正在加载 Excel 文件...")
        try:
            self.data['布料跟踪表'] = pd.read_excel(self.excel_path, sheet_name='布料跟踪表', engine='openpyxl')
            self.data['订单表'] = pd.read_excel(self.excel_path, sheet_name='订单表', engine='openpyxl')
            self.data['SKU-款号映射'] = pd.read_excel(self.excel_path, sheet_name='SKU-款号映射', engine='openpyxl')
            self.data['款号信息表'] = pd.read_excel(self.excel_path, sheet_name='款号信息表', engine='openpyxl')
            self.logger.info("数据加载成功")
        except Exception as e:
            raise Exception(f"加载 Excel 文件失败: 请确保表格包含所有的Sheet页。错误详情: {str(e)}")

    def is_solid(self, product_name: str) -> bool:
        if pd.isna(product_name): return False
        return str(product_name).strip().endswith('纯色')

    def is_print(self, product_name: str) -> bool:
        if pd.isna(product_name): return False
        return str(product_name).strip().endswith('印花')

    def extract_color_name(self, color_code: str) -> str:
        if pd.isna(color_code): return ""
        parts = str(color_code).strip().split()
        return parts[-1] if len(parts) >= 2 else str(color_code).strip()

    def extract_print_code(self, color_code: str) -> str:
        if pd.isna(color_code): return ""
        color_str = str(color_code).strip()
        match = re.match(r'(YH\d+)', color_str)
        return match.group(1) if match else color_str

    def build_df1(self) -> pd.DataFrame:
        df1 = self.data['布料跟踪表'].copy()
        df1['下单数量'] = 0.0
        df1['已收货数量'] = 0.0
        df1['款号'] = None

        order_df = self.data['订单表'].copy()
        sku_mapping = self.data['SKU-款号映射']

        order_df['颜色名称'] = order_df['颜色'].apply(self.extract_color_name)
        order_df['印花色号'] = order_df['颜色'].apply(self.extract_print_code)

        for idx, row in df1.iterrows():
            order_no = row['订单编号']
            product_name = row['产品名称']
            color_code = row['颜色&色号']

            if self.is_solid(product_name):
                color_name = self.extract_color_name(color_code)
                matched_orders = order_df[(order_df['订单编号'] == order_no) & (order_df['颜色名称'] == color_name)]
            elif self.is_print(product_name):
                print_code = self.extract_print_code(color_code)
                matched_orders = order_df[(order_df['订单编号'] == order_no) & (order_df['印花色号'] == print_code)]
            else:
                matched_orders = order_df[(order_df['订单编号'] == order_no) & (order_df['颜色'] == color_code)]

            if len(matched_orders) > 0:
                df1.at[idx, '下单数量'] = matched_orders['下单数量'].sum()
                df1.at[idx, '已收货数量'] = matched_orders['已收货数量'].sum()
                sku = matched_orders['SKU编码'].iloc[0]
                style_match = sku_mapping[sku_mapping['SKU'] == sku]['款号']
                if len(style_match) > 0:
                    df1.at[idx, '款号'] = style_match.iloc[0]
        return df1

    def build_df1_1(self, df1: pd.DataFrame) -> pd.DataFrame:
        df1_1 = df1[df1['金额（元）'].notna()].copy()
        return df1_1[df1_1['金额（元）'] > 0]

    def build_df1_2(self, df1_1: pd.DataFrame) -> pd.DataFrame:
        df1_2 = df1_1[df1_1['已收货数量'].notna()].copy()
        df1_2 = df1_2[df1_2['已收货数量'] > 0]
        df1_2 = df1_2[df1_2['已收货数量'] >= df1_2['下单数量'] * 0.70]
        return df1_2[df1_2['款号'].notna()]

    def get_valid_date_range(self, style_df: pd.DataFrame):
        if len(style_df) == 0: return None, None, []
        style_df = style_df.sort_values('下单日期', ascending=False)
        all_orders = style_df[['下单日期', '订单编号']].drop_duplicates().sort_values('下单日期', ascending=False)

        recent_date = all_orders['下单日期'].iloc[0]
        recent_order = all_orders['订单编号'].iloc[0]
        valid_dates, valid_orders, current_date = [recent_date], [recent_order], recent_date

        for idx in range(1, len(all_orders)):
            prev_date = all_orders['下单日期'].iloc[idx]
            prev_order = all_orders['订单编号'].iloc[idx]
            months_diff = (current_date.year - prev_date.year) * 12 + (current_date.month - prev_date.month)
            if months_diff > 2: break
            valid_dates.append(prev_date)
            valid_orders.append(prev_order)
            current_date = prev_date

        start_date, end_date = min(valid_dates), max(valid_dates)
        if (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month) > 6:
            cutoff_date = end_date - timedelta(days=180)
            valid_dates_filtered = [d for d in valid_dates if d >= cutoff_date]
            valid_orders = [all_orders[all_orders['下单日期'] == d]['订单编号'].iloc[0] for d in valid_dates_filtered]
            start_date = min(valid_dates_filtered) if valid_dates_filtered else end_date

        return start_date, end_date, valid_orders

    def lookup_order_quantities(self, order_no: str, color_code: str, product_name: str):
        order_df = self.data['订单表']
        if '颜色名称' not in order_df.columns:
            order_df = order_df.copy()
            order_df['颜色名称'] = order_df['颜色'].apply(self.extract_color_name)
            order_df['印花色号'] = order_df['颜色'].apply(self.extract_print_code)

        if self.is_solid(product_name):
            matched_orders = order_df[
                (order_df['订单编号'] == order_no) & (order_df['颜色名称'] == self.extract_color_name(color_code))]
        elif self.is_print(product_name):
            matched_orders = order_df[
                (order_df['订单编号'] == order_no) & (order_df['印花色号'] == self.extract_print_code(color_code))]
        else:
            matched_orders = order_df[(order_df['订单编号'] == order_no) & (order_df['颜色'] == color_code)]

        if len(matched_orders) > 0: return matched_orders['下单数量'].sum(), matched_orders['已收货数量'].sum()
        return 0.0, 0.0

    def build_df2(self, df1_2: pd.DataFrame, style_no: str) -> pd.DataFrame:
        style_df = df1_2[df1_2['款号'] == style_no].copy()
        if len(style_df) == 0: return pd.DataFrame()

        start_date, end_date, valid_orders = self.get_valid_date_range(style_df)
        if start_date is None or len(valid_orders) == 0: return pd.DataFrame()

        df2_raw = style_df[style_df['订单编号'].isin(valid_orders)].copy()
        grouped = df2_raw.groupby(['订单编号', '颜色&色号', '产品名称', '布料名称', '款号'], as_index=False).agg({
            '下单日期': 'first', '布料送货数量（匹）': 'sum', '金额（元）': 'sum'
        })

        grouped['下单数量'], grouped['已收货数量'] = 0.0, 0.0
        for idx, row in grouped.iterrows():
            order_qty, received_qty = self.lookup_order_quantities(row['订单编号'], row['颜色&色号'], row['产品名称'])
            grouped.at[idx, '下单数量'] = order_qty
            grouped.at[idx, '已收货数量'] = received_qty

        return grouped

    def calculate_cost(self, df2: pd.DataFrame):
        if len(df2) == 0: return {'K1': 0.0, 'K1-s': 0.0, 'K1-p': 0.0, 'Q1': 0.0}

        C1, N1, M1 = df2['金额（元）'].sum(), df2['布料送货数量（匹）'].sum(), df2['已收货数量'].sum()
        Q1 = M1 / N1 if N1 > 0 else 0.0
        K1 = C1 / N1 / Q1 if N1 > 0 and Q1 > 0 else 0.0

        df2_s = df2[df2['产品名称'].apply(self.is_solid)]
        K1_s = (df2_s['金额（元）'].sum() / df2_s['布料送货数量（匹）'].sum() / (
                    df2_s['已收货数量'].sum() / df2_s['布料送货数量（匹）'].sum())) if len(df2_s) > 0 and df2_s[
            '布料送货数量（匹）'].sum() > 0 and (df2_s['已收货数量'].sum() / df2_s['布料送货数量（匹）'].sum()) > 0 else 0.0

        df2_p = df2[df2['产品名称'].apply(self.is_print)]
        K1_p = (df2_p['金额（元）'].sum() / df2_p['布料送货数量（匹）'].sum() / (
                    df2_p['已收货数量'].sum() / df2_p['布料送货数量（匹）'].sum())) if len(df2_p) > 0 and df2_p[
            '布料送货数量（匹）'].sum() > 0 and (df2_p['已收货数量'].sum() / df2_p['布料送货数量（匹）'].sum()) > 0 else 0.0

        return {'K1': round(K1, 2), 'K1-s': round(K1_s, 2), 'K1-p': round(K1_p, 2), 'Q1': round(Q1, 0)}

    def get_valid_style_numbers(self) -> list:
        style_info = self.data['款号信息表']
        return style_info[style_info['产品定位'] != '淘汰款']['款号'].tolist()

    def calculate(self):
        self.load_data()
        df1 = self.build_df1()
        df1_1 = self.build_df1_1(df1)
        df1_2 = self.build_df1_2(df1_1)

        target_styles = self.get_valid_style_numbers()

        results = []
        for style_no in target_styles:
            df2 = self.build_df2(df1_2, style_no)
            if len(df2) == 0: continue

            cost_results = self.calculate_cost(df2)
            results.append({
                '款号': str(style_no),
                '成衣数Q1': cost_results['Q1'],
                '布料费用K1': cost_results['K1'],
                '纯色布料费用K1-s': cost_results['K1-s'],
                '印花布料费用K1-p': cost_results['K1-p']
            })

        return results


# 这是一个对外的统一接口，主程序只需要调用这个函数就行了
def run_fabric_calculation(file):
    calculator = FabricCostCalculator(file)
    results = calculator.calculate()
    if results:
        # 整理成表格返回
        df_results = pd.DataFrame(results)
        df_results = df_results[['款号', '成衣数Q1', '布料费用K1', '纯色布料费用K1-s', '印花布料费用K1-p']]
        return df_results
    return None