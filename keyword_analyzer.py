import pandas as pd
import logging
import re

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def _parse_time_from_filename(filename: str):
    """
    从文件名中解析年月信息
    支持格式：
    - 202501.xlsx
    - 2025-01.xlsx
    - 2025_01.xlsx
    """
    time_value = filename.replace('.xlsx', '').replace('.xls', '')
    clean_time_str = re.sub(r'\D', '', time_value)

    if len(clean_time_str) >= 6:
        year = clean_time_str[:4]
        month = clean_time_str[4:6]
        return year, month, f"{year}{month}"
    else:
        return "未知年份", clean_time_str, clean_time_str


def run_keyword_analysis(uploaded_files, threshold_rank: int):
    """
    处理多个西柚搜索词Excel文件：
    1. 合并所有文件
    2. 按“年份-月份”分别统计流量占比Top N关键词
    3. 取所有月份Top N关键词并集
    4. 从原始数据中筛选出这些关键词，形成 df2
    """
    all_data = []

    for file in uploaded_files:
        filename = file.name
        year, month, time_key = _parse_time_from_filename(filename)

        try:
            df = pd.read_excel(file)

            required_columns = [
                '关键词 (数据来源于西柚找词)',
                '流量',
                '流量占比',
                '周平均搜索量'
            ]

            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"文件 {filename} 缺少必需的列: {missing_columns}")

            df_filtered = df[required_columns].copy()
            df_filtered['年份'] = year
            df_filtered['月份'] = month
            df_filtered['时间'] = time_key

            all_data.append(df_filtered)

        except Exception as e:
            raise Exception(f"读取文件 {filename} 失败: {str(e)}")

    if not all_data:
        raise ValueError("没有成功读取任何数据")

    # 合并所有文件的数据，形成原始 df
    df = pd.concat(all_data, ignore_index=True)

    # 清洗“流量占比”字段为数值型
    if df['流量占比'].dtype == object:
        df['流量占比'] = df['流量占比'].astype(str).str.replace('%', '', regex=False)
    df['流量占比'] = pd.to_numeric(df['流量占比'], errors='coerce')

    # 如果流量占比是 0~100，统一转成 0~1
    if df['流量占比'].max(skipna=True) is not None and df['流量占比'].max(skipna=True) > 1:
        df['流量占比'] = df['流量占比'] / 100.0

    # 清洗数值列
    df['流量'] = pd.to_numeric(df['流量'], errors='coerce')
    df['周平均搜索量'] = pd.to_numeric(df['周平均搜索量'], errors='coerce')

    # 统计每个“年份-月份”中流量占比 Top N 的关键词
    q_set = set()

    grouped = df.groupby(['年份', '月份'], dropna=False)

    for (year, month), group in grouped:
        # 按流量占比降序，取前 threshold_rank 个关键词
        top_keywords = (
            group.sort_values('流量占比', ascending=False)
                 .head(threshold_rank)['关键词 (数据来源于西柚找词)']
                 .dropna()
                 .astype(str)
                 .tolist()
        )
        q_set.update(top_keywords)

    if not q_set:
        raise ValueError("未能提取到任何Top关键词，请检查数据内容")

    # 形成 df2：保留 q_set 中关键词对应的全部时间数据
    df2 = df[df['关键词 (数据来源于西柚找词)'].astype(str).isin(q_set)].copy()

    # 排序：先年份，再月份，再流量占比
    df2 = df2.sort_values(['关键词 (数据来源于西柚找词)', '年份', '月份'], ascending=[True, True, True])

    return df2