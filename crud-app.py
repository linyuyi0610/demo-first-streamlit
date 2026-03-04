import streamlit as st
import gspread
import pandas as pd


# ==========================================
# 1. 建立 Google Sheets 連線
# ==========================================
@st.cache_resource
def init_connection():
    credentials = dict(st.secrets["gcp_service_account"])
    gc = gspread.service_account_from_dict(credentials)
    return gc


gc = init_connection()

# ==========================================
# 2. 開啟指定的試算表與工作表
# ==========================================
SHEET_INPUT = "https://docs.google.com/spreadsheets/d/1O8oAoKowNwUpJekiuc8gKdJ7OUVqhxT3xBKk3lzqaW4/edit?pli=1&gid=0#gid=0"
WORKSHEET_NAME = "工作表1"

try:
    if SHEET_INPUT.startswith("http://") or SHEET_INPUT.startswith("https://"):
        sh = gc.open_by_url(SHEET_INPUT)
    else:
        sh = gc.open(SHEET_INPUT)
    worksheet = sh.worksheet(WORKSHEET_NAME)
except Exception as e:
    st.error(
        f"無法開啟試算表，請確認名稱/網址是否正確，"
        f"且服務帳號 ({gc.auth.signer_email}) 已被加入共用編輯者！\n錯誤訊息：{e}"
    )
    st.stop()

st.title("📊 Google Sheets 讀寫測試儀表板")

# ==========================================
# 用 session_state 暫存操作結果訊息
# ==========================================
if "message" in st.session_state:
    msg = st.session_state.pop("message")
    if msg["type"] == "success":
        st.success(msg["text"])
    elif msg["type"] == "error":
        st.error(msg["text"])

# ==========================================
# 3. 讀取資料 (Read) — 動態偵測欄位名稱
# ==========================================
st.header("1️⃣ 目前資料列表")

# 先讀取標題列，去除前後空白
headers = [h.strip() for h in worksheet.row_values(1)]

if len(headers) < 2:
    st.error("工作表的第一列至少需要兩個欄位標題（例如：姓名, 數量）")
    st.stop()

COL_NAME = headers[0]  # 第一欄標題（例如「姓名」）
COL_QTY = headers[1]   # 第二欄標題（例如「數量」）

st.caption(f"偵測到的欄位標題：**{COL_NAME}** / **{COL_QTY}**")

data = worksheet.get_all_records()

if data:
    # 統一去除 key 的前後空白，避免隱藏空格造成 KeyError
    data = [{k.strip(): v for k, v in row.items()} for row in data]

    df = pd.DataFrame(data)
    df.insert(0, "試算表列數", range(2, len(data) + 2))
    st.dataframe(df, use_container_width=True)
else:
    st.info("目前工作表中沒有資料。")

st.divider()

# ==========================================
# 4. 新增資料 (Create)
# ==========================================
st.header("2️⃣ 新增資料")

with st.form("add_data_form", clear_on_submit=True):
    col1 = st.text_input(COL_NAME, key="add_name")
    col2 = st.number_input(COL_QTY, min_value=0, value=1, key="add_qty")

    submitted = st.form_submit_button("寫入 Google Sheet")

    if submitted:
        if col1.strip() == "":
            st.warning(f"請填寫{COL_NAME}！")
        else:
            try:
                with st.spinner("正在寫入資料中..."):
                    worksheet.append_row([col1, int(col2)])
                st.session_state["message"] = {
                    "type": "success",
                    "text": "資料已成功寫入！",
                }
            except Exception as e:
                st.session_state["message"] = {
                    "type": "error",
                    "text": f"寫入失敗：{e}",
                }
            st.rerun()

st.divider()

# 只有在有資料的時候，才顯示修改與刪除的區塊
if data:
    row_options = {
        f"第 {i + 2} 列: {row.get(COL_NAME, '(空)')}": i + 2
        for i, row in enumerate(data)
    }

    col_update, col_delete = st.columns(2)

    # ==========================================
    # 5. 修改資料 (Update)
    # ==========================================
    with col_update:
        st.header("3️⃣ 修改資料")

        selected_option_update = st.selectbox(
            "選擇要修改的資料",
            options=list(row_options.keys()),
            key="update_select",
        )
        selected_row_update = row_options[selected_option_update]
        current_data = data[selected_row_update - 2]

        with st.form("update_data_form"):
            new_name = st.text_input(
                f"新{COL_NAME}", value=str(current_data.get(COL_NAME, ""))
            )
            new_qty = st.number_input(
                f"新{COL_QTY}",
                min_value=0,
                value=int(current_data.get(COL_QTY, 0)),
            )
            update_submitted = st.form_submit_button("更新資料")

            if update_submitted:
                if new_name.strip() == "":
                    st.warning(f"請填寫{COL_NAME}！")
                else:
                    try:
                        with st.spinner("正在更新資料中..."):
                            worksheet.update(
                                f"A{selected_row_update}:B{selected_row_update}",
                                [[new_name, int(new_qty)]],
                            )
                        st.session_state["message"] = {
                            "type": "success",
                            "text": "資料已成功更新！",
                        }
                    except Exception as e:
                        st.session_state["message"] = {
                            "type": "error",
                            "text": f"更新失敗：{e}",
                        }
                    st.rerun()

    # ==========================================
    # 6. 刪除資料 (Delete)
    # ==========================================
    with col_delete:
        st.header("4️⃣ 刪除資料")

        selected_option_del = st.selectbox(
            "選擇要刪除的資料",
            options=list(row_options.keys()),
            key="delete_select",
        )
        selected_row_del = row_options[selected_option_del]

        st.write(f"⚠️ 即將刪除：**{selected_option_del}**")

        if st.button("🗑️ 確認刪除這筆資料", type="primary"):
            try:
                with st.spinner("正在刪除資料中..."):
                    worksheet.delete_rows(selected_row_del)
                st.session_state["message"] = {
                    "type": "success",
                    "text": "資料已成功刪除！",
                }
            except Exception as e:
                st.session_state["message"] = {
                    "type": "error",
                    "text": f"刪除失敗：{e}",
                }
            st.rerun()
