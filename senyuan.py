import streamlit as st
import pandas as pd
import io

# 檔案路徑
# EXCEL_PATH = r"D:\\OneDrive - 華新麗華股份有限公司\\料號比對最終資料\\歐億_比對結果.xlsx"
EXCEL_PATH = "森元_比對結果.xlsx"
# 標題
st.set_page_config(page_title="森元料號比對系統", layout="wide")
st.title("森元料號比對系統")

# 讀取資料
@st.cache_data
def load_data():
    return pd.read_excel(EXCEL_PATH, engine="openpyxl")

df = load_data()

# 篩選開關
show_filter = st.toggle("啟用篩選功能", value=False)

# 預設顯示所有資料
filtered_df = df.copy()

if show_filter:
    st.sidebar.header("🔍 查詢模式")
    mode = st.sidebar.radio("選擇查詢方式", ["依森元料號/品名規格", "依線種/電壓/尺寸/尺寸單位/芯數/顏色"])

    if mode == "依森元料號/品名規格":
        col1, col2 = st.columns(2)

        # 取得所有料號和品名的選項
        part_numbers = [""] + df['森元料號'].dropna().unique().tolist()
        product_names = [""] + df["品名規格"].dropna().unique().tolist()

        with col1:
            selected_part_number = st.selectbox("📌選擇料號", options=part_numbers, key="part_number_select")

        with col2:
            # 如果選擇了料號，自動找到對應的品名
            if selected_part_number:
                matched_names = df[df["森元料號"] == selected_part_number]["品名規格"].dropna().unique()
                if len(matched_names) > 0:
                    # 如果有對應的品名，只顯示對應的品名選項
                    name_options = [""] + matched_names.tolist()
                    default_index = 1 if len(matched_names) == 1 else 0  # 如果只有一個品名，自動選擇
                else:
                    name_options = [""]
                    default_index = 0
            else:
                # 如果沒有選擇料號，顯示所有品名選項
                name_options = product_names
                default_index = 0

            selected_name = st.selectbox("📌選擇品名", options=name_options, key="product_name_select", index=default_index)

        # 自動搜尋邏輯（不需要按鈕）
        if selected_part_number:
            filtered_df = df[df["森元料號"] == selected_part_number]
        elif selected_name:
            filtered_df = df[df["品名規格"] == selected_name]

    elif mode == "依線種/電壓/尺寸/尺寸單位/芯數/顏色":
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        line_type = col1.multiselect("線種", options=df["產品名稱(線種)"].dropna().unique())
        voltage = col2.multiselect("電壓", options=df["電壓"].dropna().unique())
        product_type = col3.multiselect("尺寸", options=df["尺寸"].dropna().unique())
        size = col4.multiselect("尺寸單位", options=df["單位"].dropna().unique())
        cores = col5.multiselect("芯數", options=df["芯數"].dropna().unique())
        color = col6.multiselect("顏色", options=df["顏色"].dropna().unique())

        # 自動搜尋邏輯（不需要按鈕）
        if line_type or voltage or product_type or size or color:
            filtered_df = df.copy()
            if line_type:
                filtered_df = filtered_df[filtered_df["產品名稱(線種)"].isin(line_type)]
            if voltage:
                filtered_df = filtered_df[filtered_df["電壓"].isin(voltage)]
            if product_type:
                filtered_df = filtered_df[filtered_df["尺寸"].isin(product_type)]
            if size:
                filtered_df = filtered_df[filtered_df["單位"].isin(size)]
            if cores:
                filtered_df = filtered_df[filtered_df["芯數"].isin(cores)]
            if color:
                filtered_df = filtered_df[filtered_df["顏色"].isin(color)]

# 查無資料提示
if filtered_df.empty:
    st.warning("查無符合資料，請重新選擇條件")
else:
    st.success(f"共找到 {len(filtered_df)} 筆資料")

    # 顯示資料（篩選後或全部）
    with st.expander("📋 顯示查詢結果", expanded=True):
        st.dataframe(filtered_df, use_container_width=True)

    # 匯出Excel按鈕
    def to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='篩選結果')
        return output.getvalue()

    excel_bytes = to_excel(filtered_df)
    st.download_button(
        label="📥 匯出為 Excel",
        data=excel_bytes,
        file_name="篩選結果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # 列印按鈕
    if st.button("🖨️ 列印畫面"):
        st.info("請使用瀏覽器的列印功能（Ctrl+P 或 Command+P）進行列印")

