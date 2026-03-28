import pandas as pd
from datetime import timedelta
import logging

# 设置基础日志，替代 Coze 的 logger
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ShippingCostCalculator:
    """运费费用计算器"""

    def __init__(self, excel_file):
        # 直接接收网页传过来的文件对象
        self.excel_file = excel_file
        self.logger = logger
        self.data = {}
        self.df1 = None
        self.df2 = None
        self.df3 = None
        self.df4 = None

    def load_data(self):
        self.logger.info("正在加载 Excel 文件...")
        try:
            sheets = ['发货日报表', 'FBA头程', 'SKU-款号映射', '款号信息表']
            for sheet in sheets:
                self.data[sheet] = pd.read_excel(self.excel_file, sheet_name=sheet, engine='openpyxl')
            self.logger.info("数据加载成功")
        except Exception as e:
            raise Exception(
                f"加载 Excel 文件失败，请确认工作表是否包含【发货日报表, FBA头程, SKU-款号映射, 款号信息表】: {str(e)}")

    def build_df1(self):
        fba_df = self.data['FBA头程'].copy()
        fba_df = fba_df[(fba_df['总金额'].notna()) & (fba_df['总金额'] > 0)]

        fba_df['发货日期'] = pd.to_datetime(fba_df['发货日期'], errors='coerce')
        fba_df = fba_df.dropna(subset=['发货日期'])

        if len(fba_df) == 0:
            raise Exception("FBA头程 表格中没有有效的发货日期数据")

        last_date = fba_df.sort_values('发货日期')['发货日期'].iloc[-1]
        one_year_ago = last_date - timedelta(days=365)

        self.df1 = fba_df[fba_df['发货日期'] >= one_year_ago].copy()

    def build_df2(self):
        daily_df = self.data['发货日报表'].copy()

        sku_mapping = dict(zip(self.data['SKU-款号映射']['SKU'], self.data['SKU-款号映射']['款号']))
        daily_df['款号'] = daily_df['SKU编码'].map(sku_mapping)

        order_amount = dict(zip(self.df1['发货订单编号'], self.df1['总金额']))
        daily_df['金额'] = daily_df['发货订单编号'].map(order_amount)

        style_info = self.data['款号信息表']
        eliminated_styles = style_info[style_info['产品定位'] == '淘汰款']['款号'].tolist()

        self.df2 = daily_df[
            (~daily_df['款号'].isin(eliminated_styles)) &
            (daily_df['款号'].notna()) &
            (daily_df['金额'].notna())
            ].copy()

    def build_df3_and_df4(self):
        self.df3 = self.df2.groupby(['发货订单编号', '款号'], as_index=False).agg({
            '发货数量': 'sum',
            '金额': 'first'
        })

        order_total_qty = self.df3.groupby('发货订单编号')['发货数量'].transform('sum')
        self.df4 = self.df3.copy()
        self.df4['数量占比'] = self.df4['发货数量'] / order_total_qty

    def calculate_final(self):
        results = []
        grouped = self.df4.groupby('款号')

        for style_no, group in grouped:
            numerator = (group['金额'] * group['数量占比']).sum()
            denominator = group['发货数量'].sum()
            shipping_cost = round(numerator / denominator, 2) if denominator > 0 else 0

            results.append({
                '款号': style_no,
                '运费': shipping_cost
            })
        return results


# 对外统一接口
def run_shipping_calculation(file):
    calculator = ShippingCostCalculator(file)
    calculator.load_data()
    calculator.build_df1()
    calculator.build_df2()
    calculator.build_df3_and_df4()

    final_results = calculator.calculate_final()

    if final_results:
        df_costs = pd.DataFrame(final_results)
        df4_detail = calculator.df4
        # 因为运费计算原本会输出两个 Sheet 页，所以我们用字典把它们包起来返回
        return {
            "运费结果": df_costs,
            "df4明细数据": df4_detail
        }
    return None