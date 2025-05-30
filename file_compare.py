import streamlit as st
import pandas as pd
from fuzzywuzzy import fuzz, process
from io import BytesIO
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode

st.set_page_config(layout="wide")
st.title("Multi-Excel Comparator with Fuzzy Matching")

uploaded_files = st.sidebar.file_uploader("Upload 2 or more Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files and len(uploaded_files) >= 2:
    file_names = [f.name for f in uploaded_files]
    threshold = st.slider("Fuzzy Match Threshold", 70, 100, 90)

    temp_df = pd.read_excel(uploaded_files[0])
    compare_cols = st.multiselect("Columns to Compare", temp_df.columns.tolist(), default=[temp_df.columns[0]])

    def consolidated_presence_matrix(files, filenames, columns, threshold):
        file_data = [pd.read_excel(f) for f in files]
        all_values = set()
        value_counts = {}

        for idx, df in enumerate(file_data):
            df = df.dropna(subset=columns)
            df['__combined__'] = df[columns].astype(str).agg(' - '.join, axis=1)
            value_counts[filenames[idx]] = df['__combined__'].value_counts()
            all_values.update(df['__combined__'].unique())

        all_values = sorted(list(all_values))
        summary_data = {'Unique Value': all_values}

        for filename in filenames:
            summary_data[filename] = [value_counts.get(filename, {}).get(v, 0) for v in all_values]

        summary_df = pd.DataFrame(summary_data)
        st.subheader("Summary: Count of Each Unique Value Across Files")

        gb = GridOptionsBuilder.from_dataframe(summary_df)
        gb.configure_default_column(filter=True, sortable=True, resizable=True, editable=True, groupable=True)
        gb.configure_grid_options(domLayout='autoHeight', suppressHorizontalScroll=True, suppressColumnVirtualisation=True)
        gb.configure_auto_height(autoHeight=True)
        for col in summary_df.columns:
            gb.configure_column(col, autoSizeColumns=True, suppressSizeToFit=True)
        grid_options = gb.build()

        summary_response = AgGrid(
            summary_df,
            gridOptions=grid_options,
            enable_enterprise_modules=True,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            allow_unsafe_jscode=True,
            editable=True,
            return_mode='AS_INPUT'
        )

        filtered_summary = summary_response['data']
        buffer_summary = BytesIO()
        pd.DataFrame(filtered_summary).to_excel(buffer_summary, index=False)
        buffer_summary.seek(0)
        st.download_button(
            label="Download Filtered Summary",
            data=buffer_summary,
            file_name="filtered_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Comparison matrix
        key_sets = [set(vc.index) for vc in value_counts.values()]
        matrix = []

        for key in all_values:
            presence_row = {"Value": key}
            for idx, key_set in enumerate(key_sets):
                match, score = process.extractOne(key, list(key_set), scorer=fuzz.token_sort_ratio)
                presence_row[filenames[idx]] = "Yes" if score >= threshold else "No"
            matrix.append(presence_row)

        return pd.DataFrame(matrix)

    if compare_cols:
        st.subheader("Consolidated Presence Matrix")
        matrix_df = consolidated_presence_matrix(uploaded_files, file_names, compare_cols, threshold)
        gb = GridOptionsBuilder.from_dataframe(matrix_df)
        gb.configure_default_column(filter=True, sortable=True, resizable=True, editable=True, groupable=True)
        gb.configure_grid_options(domLayout='autoHeight', suppressHorizontalScroll=True, suppressColumnVirtualisation=True)
        gb.configure_auto_height(autoHeight=True)
        for col in matrix_df.columns:
            gb.configure_column(col, autoSizeColumns=True, suppressSizeToFit=True)
        grid_options = gb.build()

        matrix_response = AgGrid(
            matrix_df,
            gridOptions=grid_options,
            enable_enterprise_modules=True,
            update_mode=GridUpdateMode.MODEL_CHANGED,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            allow_unsafe_jscode=True,
            editable=True,
            return_mode='AS_INPUT'
        )

        filtered_matrix = matrix_response['data']
        buffer_matrix = BytesIO()
        pd.DataFrame(filtered_matrix).to_excel(buffer_matrix, index=False)
        buffer_matrix.seek(0)
        st.download_button(
            label="Download Filtered Presence Matrix",
            data=buffer_matrix,
            file_name="filtered_presence_matrix.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Please upload at least two Excel files to begin.")
