import plotly.express as px
import pandas as pd


def plot_keyword_trends(df, keyword, width=400, height=350):
    """
    绘制单个关键词的流量趋势折线图
    - df: 包含该关键词所有年份、月份数据的DataFrame，必须含‘年份’, ‘月份’, ‘流量占比’列
    - keyword: 当前关键词名称（字符串）
    - width, height: 图表大小自定义参数

    逻辑说明：
    - x轴显示“月份”数字（01、02 ...）
    - y轴显示“流量占比”，小数0~1格式
    - 多条线按照不同‘年份’分组
    """
    # 确保月份是字符串，且排好顺序，避免默认按数字排序导致线错乱
    df['月份'] = df['月份'].astype(str).str.zfill(2)
    df = df.sort_values('月份')

    fig = px.line(
        df,
        x='月份',
        y='流量占比',
        color='年份',
        title=f"关键词：{keyword}",
        markers=True
    )

    fig.update_layout(
        xaxis_title='月份',
        yaxis_title='流量占比',
        yaxis_tickformat='.2%',
        xaxis_type='category',  # 按类别而不是连续数轴显示
        margin=dict(l=30, r=30, t=40, b=30),
        width=width,
        height=height
    )

    return fig