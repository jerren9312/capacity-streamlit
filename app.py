# -*- coding: utf-8 -*-
"""
Created on Mon Aug  4 01:22:22 2025

@author: bookn
"""

import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="产能计算工具", layout="centered")

st.title("📊 产能计算工具（支持 Excel / CSV）")
st.markdown("请上传包含机器产能数据的 `.xlsx` 或 `.csv` 文件，我们将自动处理并生成汇总表。")

uploaded_file = st.file_uploader("选择文件", type=["xlsx", "csv"])

if uploaded_file:
    file_type = uploaded_file.name.split(".")[-1]
    
    try:
        if file_type == "xlsx":
            df = pd.read_excel(uploaded_file, sheet_name=0)
        else:
            df = pd.read_csv(uploaded_file)

        df.columns = df.columns.str.replace(r'\n', '', regex=True).str.strip()
        df = df.rename(columns={
            df.columns[0]: 'date',
            df.columns[1]: 'machine',
            df.columns[3]: 'manufacturer',
            df.columns[5]: 'thickness',
            df.columns[10]: 'large_rolls',
            df.columns[12]: 'antistatic'
        })

        df = df[df['date'].notnull() & df['machine'].astype(str).str.contains('HJ', na=False)]
        df['large_rolls'] = pd.to_numeric(df['large_rolls'], errors='coerce').fillna(0)
        df['thickness'] = pd.to_numeric(df['thickness'], errors='coerce').fillna(0)
        df['antistatic'] = df['antistatic'].astype(str).str.upper().str.strip()
        df['manufacturer'] = df['manufacturer'].astype(str).str.strip()

        def compute_multiplier(row):
            if row['machine'] == '4HJ':
                return 1
            if row['antistatic'] == 'Y' and row['manufacturer'] == 'ZJ' and row['thickness'] == 21:
                return 3
            if row['antistatic'] == 'N' and row['manufacturer'] == 'JT' and row['thickness'] in [18, 20]:
                return 2
            if row['antistatic'] == 'Y':
                return 1.5
            return 1

        df['multiplier'] = df.apply(compute_multiplier, axis=1)
        df['capacity'] = df['large_rolls'] * df['multiplier']

        pivot_df = df.pivot_table(index='date', columns='machine', values='capacity', aggfunc='sum').fillna(0).reset_index()
        pivot_df['date'] = pd.to_datetime(pivot_df['date'], origin='1899-12-30', unit='D', errors='ignore')
        pivot_df['date'] = pd.to_datetime(pivot_df['date'], errors='coerce').dt.strftime('%Y-%m-%d')

        st.success("✅ 处理完成！你可以下载下方结果表格：")
        st.dataframe(pivot_df)

        # Download link
        csv = pivot_df.to_csv(index=False, encoding='utf-8-sig')
        st.download_button("📥 下载产能汇总 CSV", csv, file_name="output_capacity.csv", mime="text/csv")

    except Exception as e:
        st.error(f"❌ 处理失败：{e}")
