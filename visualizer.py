import plotly.express as px


def create_keyword_trend_fig(kw_data, keyword_name):
    """
    绘制单个关键词的流量占比趋势折线图

    参数:
    kw_data (DataFrame): 单个关键词的数据，需包含 '月份', '年份', '流量占比'
    keyword_name (str): 关键词名称
    """
    # 保证月份是字符串，避免 1 和 01 混乱
    kw_data = kw_data.copy()
    kw_data['月份'] = kw_data['月份'].astype(str).str.zfill(2)
    kw_data['年份'] = kw_data['年份'].astype(str)

    # 排序，保证折线连接顺序正常
    kw_data = kw_data.sort_values(['年份', '月份'])

    fig = px.line(
        kw_data,
        x='月份',
        y='流量占比',
        color='年份',
        markers=True,
        title=keyword_name
    )

    fig.update_layout(
        xaxis_title='月份',
        yaxis_title='流量占比',
        yaxis_tickformat='.2%',
        xaxis_type='category',
        margin=dict(l=20, r=20, t=40, b=20),
        height=360,
        legend_title='年份'
    )

    return fig