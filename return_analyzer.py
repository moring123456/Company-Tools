import pandas as pd
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def detect_encoding_from_file(uploaded_file) -> str:
    """自动检测 Streamlit 上传文件的编码"""
    bytes_data = uploaded_file.getvalue()
    encodings = ['utf-8', 'windows-1252', 'gbk', 'gb18030', 'utf-8-sig', 'latin-1']
    for enc in encodings:
        try:
            bytes_data.decode(enc)
            return enc
        except (UnicodeDecodeError, UnicodeError):
            continue
    return 'latin-1'


def determine_month(df: pd.DataFrame, date_column: str = 'purchase-date') -> str:
    df_temp = df.copy()
    df_temp['month_str'] = df_temp[date_column].astype(str).str[:7]
    month_counts = df_temp['month_str'].value_counts()
    if month_counts.empty:
        return 'Unknown'
    max_count = month_counts.max()
    max_months = month_counts[month_counts == max_count].index.tolist()
    if len(max_months) > 1:
        max_months.sort(reverse=True)
    return max_months[0]


def preprocess_order_files(order_files) -> pd.DataFrame:
    all_dfs = []
    for file in order_files:
        encoding = detect_encoding_from_file(file)
        # 读取文件
        df = pd.read_csv(file, sep='\t', encoding=encoding, dtype=str)

        required_columns = ['amazon-order-id', 'purchase-date', 'sku', 'quantity']
        available_columns = [col for col in required_columns if col in df.columns]

        if len(available_columns) < len(required_columns):
            logger.warning(f"文件 {file.name} 缺少必要列，跳过处理")
            continue

        df = df[required_columns].copy()
        month = determine_month(df, 'purchase-date')
        df['month'] = month
        all_dfs.append(df)

    if not all_dfs:
        raise ValueError("没有找到有效的 order 销售数据文件")

    df1 = pd.concat(all_dfs, ignore_index=True)
    return df1


def preprocess_return_files(return_files) -> pd.DataFrame:
    all_dfs = []
    for file in return_files:
        encoding = detect_encoding_from_file(file)
        df = pd.read_csv(file, sep='\t', encoding=encoding, dtype=str)
        all_dfs.append(df)

    if not all_dfs:
        raise ValueError("没有找到有效的 return 退货数据文件")

    df_combined = pd.concat(all_dfs, ignore_index=True)

    required_columns = ['order-id', 'sku', 'quantity', 'reason']
    available_columns = [col for col in required_columns if col in df_combined.columns]

    if len(available_columns) < len(required_columns):
        missing = set(required_columns) - set(available_columns)
        raise ValueError(f"退货数据缺少必要列: {missing}")

    df2 = df_combined[required_columns].copy()
    df2 = df2.rename(columns={
        'quantity': 'return-quantity',
        'reason': 'return-reason'
    })
    df2['return-quantity'] = pd.to_numeric(df2['return-quantity'], errors='coerce').fillna(0).astype(int)
    return df2


def load_sku_information(sku_file) -> tuple:
    xls = pd.ExcelFile(sku_file)
    if '映射表' not in xls.sheet_names:
        raise ValueError("SKU信息文件缺少'映射表'工作表")
    df_mapping = pd.read_excel(xls, sheet_name='映射表', dtype=str)

    if '款号信息表' not in xls.sheet_names:
        raise ValueError("SKU信息文件缺少'款号信息表'工作表")
    df_style_info = pd.read_excel(xls, sheet_name='款号信息表', dtype=str)

    return df_mapping, df_style_info


def merge_data(df1, df2, df_mapping, df_style_info) -> pd.DataFrame:
    df3 = df1.copy()

    df2_agg = df2.groupby(['order-id', 'sku']).agg({
        'return-quantity': 'sum',
        'return-reason': lambda x: '; '.join(x.dropna().unique()) if len(x.dropna().unique()) > 0 else ''
    }).reset_index()

    df3 = df3.merge(df2_agg, left_on=['amazon-order-id', 'sku'], right_on=['order-id', 'sku'], how='left')
    df3['return-quantity'] = df3['return-quantity'].fillna(0).astype(int)
    df3['return-reason'] = df3['return-reason'].fillna('')

    if 'order-id' in df3.columns:
        df3 = df3.drop(columns=['order-id'])

    df3['款号'] = ''
    df3['颜色'] = ''
    df3['尺码'] = ''
    df3['纯色/印花'] = ''

    if 'SKU' in df_mapping.columns and '款号' in df_mapping.columns:
        mapping_cols = ['SKU']
        for col in ['款号', '颜色', '尺码', '产品名称']:
            if col in df_mapping.columns:
                mapping_cols.append(col)

        df_mapping_subset = df_mapping[mapping_cols].copy().rename(columns={'SKU': 'sku'})
        df3 = df3.merge(df_mapping_subset, on='sku', how='left', suffixes=('', '_mapping'))

        for col in ['款号', '颜色', '尺码']:
            if f'{col}_mapping' in df3.columns:
                df3[col] = df3[f'{col}_mapping'].fillna('')
                df3 = df3.drop(columns=[f'{col}_mapping'])

        if '产品名称' in df3.columns:
            df3['纯色/印花'] = df3['产品名称'].fillna('').str[-2:]
            df3 = df3.drop(columns=['产品名称'])
        elif '产品名称_mapping' in df3.columns:
            df3['纯色/印花'] = df3['产品名称_mapping'].fillna('').str[-2:]
            df3 = df3.drop(columns=['产品名称_mapping'])

    df3[['款号', '颜色', '尺码', '纯色/印花']] = df3[['款号', '颜色', '尺码', '纯色/印花']].fillna('')

    if '退货分析标记' in df_style_info.columns and '款号' in df_style_info.columns:
        marked_styles = df_style_info[df_style_info['退货分析标记'] == '1']['款号'].tolist()
        if marked_styles:
            df3 = df3[df3['款号'].isin(marked_styles)]

    df3['quantity'] = pd.to_numeric(df3['quantity'], errors='coerce').fillna(0).astype(int)
    df3['return-quantity'] = pd.to_numeric(df3['return-quantity'], errors='coerce').fillna(0).astype(int)

    return df3


# 对外统一接口
def run_return_analysis(order_files, return_files, sku_file):
    df1 = preprocess_order_files(order_files)
    df2 = preprocess_return_files(return_files)
    df_mapping, df_style_info = load_sku_information(sku_file)
    df3 = merge_data(df1, df2, df_mapping, df_style_info)

    # 按照你的脚本逻辑，最终导出的是 df3（源数据）
    return df3