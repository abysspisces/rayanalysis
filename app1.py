import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import re
import math
from io import BytesIO

st.set_page_config(page_title="Excel 匹配汇总工具", layout="wide")


def clean_columns(df):
    df.columns = df.columns.astype(str).str.strip()
    return df


def normalize_text(value):
    if pd.isna(value):
        return ""
    text = str(value)
    text = text.replace("\u00A0", " ")
    text = text.replace("\u2007", " ")
    text = text.replace("\u202F", " ")
    return text.strip()


def normalize_id(value):
    text = normalize_text(value)
    if text == "":
        return ""
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"\.0+$", "", text)
    return text


def parse_amount(value):
    if pd.isna(value):
        return math.nan

    text = str(value).strip()
    if text == "":
        return math.nan

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()

    text = text.replace("\u00A0", "")
    text = text.replace(" ", "")
    text = re.sub(r"[€$£¥]", "", text)
    text = re.sub(r"[^0-9,.\-]", "", text)

    if text in ("", "-", ".", ","):
        return math.nan

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "")
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text and "." not in text:
        parts = text.split(",")
        if len(parts) == 2 and len(parts[1]) in (1, 2):
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")

    try:
        number = float(text)
        if negative:
            number = -number
        return number
    except Exception:
        return math.nan


def to_excel_download(data_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in data_dict.items():
            safe_sheet_name = str(sheet_name)[:31]
            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
    output.seek(0)
    return output


st.title("表1/表2 匹配汇总工具")
st.write("上传两个 Excel 文件后，系统会按 Characteristic = Id 进行匹配，并汇总 Discount Code 对应的 Subtotal。")

col1, col2 = st.columns(2)

with col1:
    file1 = st.file_uploader("上传表1：TT DE联盟客出单情况", type=["xlsx"])

with col2:
    file2 = st.file_uploader("上传表2：DE 商城后台订单情况", type=["xlsx"])

if file1 and file2:
    try:
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        df1 = clean_columns(df1)
        df2 = clean_columns(df2)

        st.subheader("原始数据预览")
        tab1, tab2 = st.tabs(["表1预览", "表2预览"])

        with tab1:
            st.write("表1列名：", list(df1.columns))
            st.dataframe(df1.head(10), use_container_width=True)

        with tab2:
            st.write("表2列名：", list(df2.columns))
            st.dataframe(df2.head(10), use_container_width=True)

        required_df1 = ["Characteristic"]
        required_df2 = ["Id", "Financial Status", "Subtotal", "Discount Code"]

        missing_df1 = [c for c in required_df1 if c not in df1.columns]
        missing_df2 = [c for c in required_df2 if c not in df2.columns]

        if missing_df1:
            st.error("表1缺少必要列：" + str(missing_df1))
            st.stop()

        if missing_df2:
            st.error("表2缺少必要列：" + str(missing_df2))
            st.stop()

        df1["Characteristic_norm"] = df1["Characteristic"].apply(normalize_id)
        df2["Id_norm"] = df2["Id"].apply(normalize_id)
        df2["Financial Status_norm"] = df2["Financial Status"].apply(lambda x: normalize_text(x).lower())
        df2["Subtotal_num"] = df2["Subtotal"].apply(parse_amount)
        df2["Discount Code_norm"] = df2["Discount Code"].apply(normalize_text)

        # 只保留指定状态
        valid_status = ["paid", "partially refunded"]
        df2_filtered = df2[df2["Financial Status_norm"].isin(valid_status)].copy()

        # 匹配
        merged = pd.merge(
            df1[["Characteristic", "Characteristic_norm"]],
            df2_filtered[["Id", "Id_norm", "Financial Status", "Subtotal", "Subtotal_num", "Discount Code"]],
            left_on="Characteristic_norm",
            right_on="Id_norm",
            how="inner"
        )

        # 明细结果：按你的要求，仅保留对应的 Characteristic 字段
        result_detail = merged[["Characteristic"]].copy()

        # 汇总结果：按 Discount Code 汇总 Subtotal
        summary = (
            merged.groupby("Discount Code", dropna=False, as_index=False)
            .agg(
                Characteristic_Count=("Characteristic", "count"),
                Subtotal_Sum=("Subtotal_num", "sum")
            )
            .sort_values(by="Subtotal_Sum", ascending=False)
        )

        summary["Discount Code"] = summary["Discount Code"].fillna("空值/无折扣码")

        # 未匹配 Characteristic
        unmatched = df1[
            ~df1["Characteristic_norm"].isin(df2_filtered["Id_norm"])
        ][["Characteristic"]].drop_duplicates()

        discount_code_count = summary["Discount Code"].fillna("").astype(str)
        discount_code_count = (discount_code_count != "").sum()

        st.subheader("处理统计")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("表1总行数", len(df1))
        c2.metric("表2总行数", len(df2))
        c3.metric("表2筛选后", len(df2_filtered))
        c4.metric("匹配成功", len(result_detail))
        c5.metric("Discount Code 数量", int(discount_code_count))

        st.subheader("匹配后的明细结果（仅保留 Characteristic）")
        st.dataframe(result_detail, use_container_width=True)

        st.subheader("按 Discount Code 汇总结果")
        st.dataframe(summary, use_container_width=True)

        with st.expander("查看未匹配的 Characteristic"):
            st.dataframe(unmatched, use_container_width=True)

        st.subheader("可视化展示：Discount Code 对应的 Subtotal 汇总值")

        if len(summary) > 0:
            fig, ax = plt.subplots(figsize=(12, 6))
            ax.bar(summary["Discount Code"].astype(str), summary["Subtotal_Sum"], color="skyblue")
            ax.set_title("Discount Code 对应的 Subtotal 汇总值")
            ax.set_xlabel("Discount Code")
            ax.set_ylabel("Subtotal Sum")
            plt.xticks(rotation=45, ha="right")
            plt.tight_layout()
            st.pyplot(fig)
        else:
            st.warning("没有可展示的汇总数据。")

        excel_file = to_excel_download(
            {
                "matched_characteristic": result_detail,
                "discount_summary": summary,
                "unmatched_characteristic": unmatched
            }
        )

        st.download_button(
            label="下载处理结果 Excel",
            data=excel_file,
            file_name="processed_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error("处理出错：" + str(e))
else:
    st.info("请先上传两个 Excel 文件。")
