import plotly.express as px


def create_keyword_trend_fig(kw_data, keyword_name):
    """
    绘制单个关键词的流量趋势折线图

    参数:
    kw_data (DataFrame): 该关键词对应的原始数据，包含 '月份', '流量占比', '年份' 等列
    keyword_name (str): 关键词名称（用于显示在图表标题上）

    返回:
    fig (plotly.graph_objs._figure.Figure): Plotly 图表对象
    """
    # 确保 X 轴按月份从小到大排序，防止折线前后乱飞
    kw_data = kw_data.sort_values(by="月份")

    # 核心绘图逻辑：X轴=月份，Y轴=流量占比，按照"年份"画出不同的线 (color)
    fig = px.line(
        kw_data,
        x="月份",
        y="流量占比",
        color="年份",  # 有多少个年份，就会自动生成多少条不同颜色的线
        title=keyword_name,
        markers=True  # 在数据点上显示圆点
    )

    # 优化图表样式
    fig.update_layout(
        xaxis_title="月份",
        yaxis_title="流量占比",
        yaxis_tickformat='.2%',  # Y轴显示为百分比 (例如 12.34%)
        xaxis_type='category',  # 确保 X 轴被当做"分类/文本"处理，否则数字01和03中间会自动补个空白的02
        margin=dict(l=20, r=20, t=40, b=20),
        height=350,
        legend_title="年份"  # 图例的标题
    )

    return fig

# 如果以后有排料计算图、退货饼图等，可以继续在这个文件里定义：
# def create_return_pie_chart(...):
#     pass