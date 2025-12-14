import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="JollyJupiter 分析平台", layout="wide")
st.title("JollyJupiter 分析平台")

# Sidebar
st.sidebar.header("操作選單")
uploaded_file = st.sidebar.file_uploader("請上傳 Excel 檔案", type=["xls", "xlsx"])
template = st.sidebar.radio(
    "請選擇分析模板",
    ("老師月統計", "學生出席明細", "下個月預算老師清單（不看出席）", "課文詞語墳充家長")
)

def clean_school_name(school):
    # 去除開頭的 _xxx_，只保留正式校名
    return re.sub(r'^_[^_]+_', '', str(school))

def fix_column_names(df):
    # 將「學栍姓名」欄位自動改為「學生姓名」
    if '學栍姓名' in df.columns and '學生姓名' not in df.columns:
        df = df.rename(columns={'學栍姓名': '學生姓名'})
    return df

def teacher_monthly_summary(df, month):
    df = df[df['學生出席狀況'] == '出席'].copy()
    df['月份'] = month
    result = (
        df.groupby(['老師', '月份'])
        .agg(
            學生人數=('學生姓名', 'nunique'),
            本月總收入=('單堂收費', 'sum')
        )
        .reset_index()
    )
    return result

def student_attendance_detail(df, month):
    df = df[df['學生出席狀況'] == '出席'].copy()
    df['月份'] = month
    result = df[['學生姓名', '老師', '上課日期', '單堂收費', '月份']]
    return result

def teacher_email_prep_template(df):
    df['學校清理後'] = df['學校'].apply(clean_school_name)
    # 年級翻譯
    p_map = {
        'P1': '一年級', 'P2': '二年級', 'P3': '三年級',
        'P4': '四年級', 'P5': '五年級', 'P6': '六年級'
    }
    # 只翻譯 P1~P6，其餘年級保持原樣
    df['年級翻譯'] = df['年級'].map(lambda x: p_map.get(x, x))
    group = df.groupby(['學校清理後', '年級翻譯', '老師'])
    result = group.agg(
        預計學生人數=('學生編號', 'nunique'),
        未交費學生人數=('學生編號', lambda x: x[df.loc[x.index, '欠數總額'] > 0].nunique())
    ).reset_index()
    result['老師電郵'] = ""
    # 重新命名欄位
    result = result.rename(columns={'學校清理後': '學校', '年級翻譯': '年級'})
    result = result[['學校', '年級', '預計學生人數', '未交費學生人數', '老師', '老師電郵']]
    return result

def parent_vocab_template(df):
    # 先清理學校名稱
    df['學校清理後'] = df['學校'].apply(clean_school_name)
    # 只保留 P1~P6 學生
    p_map = {
        'P1': '一年級', 'P2': '二年級', 'P3': '三年級',
        'P4': '四年級', 'P5': '五年級', 'P6': '六年級'
    }
    # 只保留年級為 P1~P6 的資料
    df = df[df['年級'].isin(p_map.keys())].copy()
    # 年級翻譯
    df['年級'] = df['年級'].map(p_map)
    cols = ['學生編號', '學生姓名', '學校清理後', '年級', '家長電郵']
    result = df[[col for col in cols if col in df.columns]].copy()
    result = result.rename(columns={'學校清理後': '學校'})
    # 只在這個模板做去重
    result = result.drop_duplicates(subset=['學生編號', '學生姓名', '學校', '年級', '家長電郵'])
    return result
if uploaded_file:
    # 嘗試抓月份資訊，若沒有也不影響新模板
    try:
        info = pd.read_excel(uploaded_file, header=None, nrows=10)
        period_str = info.iloc[3, 2]
        if pd.isna(period_str):
            period_str = info.iloc[3, 1]
        month = str(period_str)[:7]
    except Exception:
        month = ""
    # 主要資料
    df = pd.read_excel(uploaded_file, header=5)
    df = fix_column_names(df)

    # Dashboard summary
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("總學生數", df['學生編號'].nunique() if '學生編號' in df.columns else "-")
    with col2:
        st.metric("總老師數", df['老師'].nunique() if '老師' in df.columns else "-")
    with col3:
        st.metric("總年級數", df['年級'].nunique() if '年級' in df.columns else "-")

    # 分析結果
    if template == "老師月統計":
        result = teacher_monthly_summary(df, month)
    elif template == "學生出席明細":
        result = student_attendance_detail(df, month)
    elif template == "下個月預算老師清單（不看出席）":
        result = teacher_email_prep_template(df)
    elif template == "課文詞語墳充家長":
        result = parent_vocab_template(df)

    st.dataframe(result)
    st.download_button("下載結果 Excel", result.to_csv(index=False).encode('utf-8-sig'), file_name=f"{template}.csv")