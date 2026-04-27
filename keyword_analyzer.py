import pandas as pd
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def run_keyword_analysis(uploaded_files, num_rank: int):
    """
    处理网页上传的多个西柚搜索词Excel文件
    """
    all_data = []

    for file in uploaded_files:
        # 获取文件名作为时间戳 (预期格式: YYYY-MM.xlsx)
        filename = file.name
        time_value = filename.replace('.xlsx', '').replace('.xls', '')

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
            df_filtered['时间'] = time_value

            # 解析年份和月份，用于折线图绘制
            # 假设文件名为 2025-01
            if '-' in time_value:
                parts = time_value.split('-')
                df_filtered['年份'] = parts[0]
                df_filtered['月份'] = parts[1]
            else:
                df_filtered['年份'] = "未知年份"
                df_filtered['月份'] = time_value

            all_data.append(df_filtered)

        except Exception as e:
            raise Exception(f"读取文件 {filename} 失败: {str(e)}")

    if not all_data:
        raise ValueError("没有成功读取任何数据")

    # 合并所有数据
    df = pd.concat(all_data, ignore_index=True)

    # 清洗"流量占比"字段，确保其为数值类型（处理可能带有 % 号的字符串）
    if df['流量占比'].dtype == object:
        df['流量占比'] = df['流量占比'].astype(str).str.replace('%', '', regex=False)
        df['流量占比'] = pd.to_numeric(df['流量占比'], errors='coerce') / 100.0

    # 计算每个关键词的流量总和
    keyword_traffic_sum = df.groupby('关键词 (数据来源于西柚找词)')['流量'].sum().reset_index()
    keyword_traffic_sum.columns = ['关键词', '流量总和']

    # 按流量总和降序排序，提取前 Num_Rank 名
    keyword_traffic_sum = keyword_traffic_sum.sort_values('流量总和', ascending=False)
    top_keywords = keyword_traffic_sum.head(num_rank)

    # 获取这 Num_Rank 个关键词的所有原始数据
    df1 = df[df['关键词 (数据来源于西柚找词)'].isin(top_keywords['关键词'])].copy()

    # 按关键词和流量总和排序，确保数据按照排名先后出现
    df1 = df1.merge(top_keywords, left_on='关键词 (数据来源于西柚找词)', right_on='关键词', how='left')
    df1 = df1.sort_values(['流量总和', '时间'], ascending=[False, True])
    df1 = df1.drop(columns=['关键词'])

    return df1