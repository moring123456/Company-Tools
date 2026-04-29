import plotly.express as px

def create_keyword_trend_fig(kw_data, keyword_name):
    """
    绘制单个关键词的流量占比趋势折线图
    确保x轴月份严格从01到12升序排列
    """
    kw_data = kw_data.copy()
    # 确保月份字符串格式，补零
    kw_data['月份'] = kw_data['月份'].astype(str).str.zfill(2)
    kw_data['年份'] = kw_data['年份'].astype(str)

    # 排序确保数据点顺序正确
    kw_data = kw_data.sort_values(['年份', '月份'])

    fig = px.line(
        kw_data,
        x='月份',
        y='流量占比',
        color='年份',
        markers=True,
        title=keyword_name
    )

    # 固定月份顺序，橫坐标类别顺序严格从01到12
    month_order = [f"{m:02d}" for m in range(1, 13)]

    fig.update_xaxes(
        categoryorder='array',
        categoryarray=month_order,
        title_text='月份'
    )

    fig.update_layout(
        yaxis_title='流量占比',
        yaxis_tickformat='.2%',
        margin=dict(l=20, r=20, t=40, b=20),
        height=360,
        legend_title='年份'
    )

    return fig