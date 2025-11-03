import streamlit as st
import pandas as pd
from excel_utils import append_many, read_last_date_str_from_template_bytes
from auth import verify, PWD_DB

st.set_page_config(page_title="拉力值紀錄 (雲端版)", layout="wide")

# ---- Auth ----
if "authed_user" not in st.session_state:
    st.session_state.authed_user = None

def login_form():
    st.title("登入 / Login")
    with st.form("login"):
        col1, col2 = st.columns(2)
        with col1:
            username = st.selectbox("帳號 (Username)", options=sorted(PWD_DB.keys()))
        with col2:
            password = st.text_input("密碼 (Password)", type="password")
        submitted = st.form_submit_button("登入")
    if submitted:
        if verify(username, password):
            st.session_state.authed_user = username
            st.success(f"歡迎，{username}")
            st.rerun()
        else:
            st.error("帳號或密碼錯誤")
    st.stop()

if st.session_state.authed_user is None:
    login_form()

# ---- Main App ----
st.sidebar.write(f"已登入：**{st.session_state.authed_user}**")
st.title("拉力值紀錄 (Pull Test Data Sheet) — 雲端版")

uploaded_file = st.file_uploader("上傳 Excel 範本 (.xlsx/.xlsm)", type=["xlsx","xlsm"])
sheet_name = st.text_input("工作表名稱（找不到會退回第一個可見表）", value="Data")

if uploaded_file is None:
    st.info("請先上傳範本檔案。")
    st.stop()

template_bytes = uploaded_file.read()

# 顯示最後一筆日期
try:
    last_date = read_last_date_str_from_template_bytes(template_bytes, sheet_name=sheet_name)
    st.caption(f"最後一筆日期：{last_date or '—'}")
except Exception as e:
    st.warning(f"讀取最後一筆日期失敗：{e}")

# 準備 12 列空白輸入
rows = []
for _ in range(12):
    rows.append({
        "Date (YYYY/MM/DD-LL)": "",
        "Value_1 (P1-1)": "",
        "Value_2 (P1-2)": "",
        "Value_3 (P1-3)": "",
        "Value_4 (P2-1)": "",
        "Value_5 (P2-2)": "",
        "Value_6 (P2-3)": "",
        "WO No.（工單號碼）": "",
    })
df = pd.DataFrame(rows)

st.write("請在下表輸入 12 筆資料（可留白的日期列會自動略過）：")
edited = st.data_editor(df, num_rows="fixed", use_container_width=True, key="input_table")

if st.button("新增紀錄並下載新檔"):
    # 組裝列
    rows_to_add = []
    for _, row in edited.iterrows():
        rows_to_add.append({
            "date": str(row["Date (YYYY/MM/DD-LL)"]).strip(),
            "values": [
                str(row["Value_1 (P1-1)"]).strip(),
                str(row["Value_2 (P1-2)"]).strip(),
                str(row["Value_3 (P1-3)"]).strip(),
                str(row["Value_4 (P2-1)"]).strip(),
                str(row["Value_5 (P2-2)"]).strip(),
                str(row["Value_6 (P2-3)"]).strip(),
            ],
            "owner": str(row["WO No.（工單號碼）"]).strip(),
        })
    try:
        out_bytes, used_sheet = append_many(template_bytes, rows_to_add, sheet_name=sheet_name)
        base = uploaded_file.name.rsplit(".", 1)[0]
        out_name = f"{base}-out.xlsx"
        st.success(f"已完成（Sheet: {used_sheet}）。")
        st.download_button(
            "下載處理後檔案",
            data=out_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(str(e))
