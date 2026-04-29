import pandas as pd
import plotly.express as px


def fill_missing_months_dynamic(df):
    """
    针对输入的 df（某关键词所有年份的数据），
    逐年份动态补齐该年份min月到max月之间的所有月份。

    输入df含字段：'年份'、'月份'、'流量占比' 等
    返回补齐后的df
    """
    # 准备一个新的空DataFrame用来存结果
    dfs = []
    for year in df['年份'].unique():
        sub_df = df[df['年份'] == year].copy()

        min_month = int(sub_df['月份'].min())
        max_month = int(sub_df['月份'].max())

        full_month_range = list(range(min_month, max_month + 1))

        # 构造完整索引
        full_index = pd.MultiIndex.from_product(
            [[year], full_month_range],
            names=['年份', '月份']
        )

        sub_df = sub_df.set_index(['年份', '月份']).reindex(full_index).reset_index()

        # 如果关键词列存在，填充它（因为它会丢失）
        if '关键词 (数据来源于西柚找词)' in df.columns:
            sub_df['关键词 (数据来源于西柚找词)'] = df['关键词 (数据来源于西柚找词)'].iloc[0]

        dfs.append(sub_df)

    df_filled = pd.concat(dfs, ignore_index=True)
    return df_filled


def create_keyword_trend_fig(kw_data, keyword_name):
    """
    单关键词绘图，动态补齐月份，X轴月份从最小月到最大月，不多余输出。
    """
    kw_data = kw_data.copy()

    # 转数值，去无效，整数月份
    kw_data['月份'] = pd.to_numeric(kw_data['月份'], errors='coerce')
    kw_data['年份'] = pd.to_numeric(kw_data['年份'], errors='coerce')
    kw_data = kw_data.dropna(subset=['月份', '年份'])
    kw_data['月份'] = kw_data['月份'].astype(int)
    kw_data['年份'] = kw_data['年份'].astype(int)

    # 先按年份分组，逐年份补齐月份区间
    kw_data_filled = fill_missing_months_dynamic(kw_data)

    kw_data_filled = kw_data_filled.sort_values(['年份', '月份'])

    # 整理X轴tickvals合并所有年份的月份范围
    min_month = kw_data['月份'].min()
    max_month = kw_data['月份'].max()
    tickvals = list(range(min_month, max_month + 1))

    fig = px.line(
        kw_data_filled,
        x='月份',
        y='流量占比',
        color='年份',
        markers=True,
        title=keyword_name,
        category_orders={"月份": tickvals},
    )

    fig.update_xaxes(
        type='category',
        tickmode='array',
        tickvals=tickvals,
        ticktext=[str(m) for m in tickvals],
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