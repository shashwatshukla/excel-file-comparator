import streamlit as st
import pandas as pd

def main():
    st.set_page_config(layout="wide")
    st.title("Excel Sheet Comparator")
    st.subheader("This app compares two Excel sheets and shows rows that are present in one sheet but not in the other.")
    # Upload two Excel files
    st.sidebar.header("Upload Excel Files")
    file1 = st.sidebar.file_uploader("Upload first Excel file", type=["xlsx"])
    file2 = st.sidebar.file_uploader("Upload second Excel file", type=["xlsx"])

    # Check if files are uploaded
    if file1 is not None and file2 is not None:
        excel_file1 = pd.ExcelFile(file1)
        excel_file2 = pd.ExcelFile(file2)

        # Select sheets for comparison
        sheet_names1 = excel_file1.sheet_names
        sheet_names2 = excel_file2.sheet_names
        sheet1 = st.sidebar.selectbox("Select Sheet in file1 to check", sheet_names1)
        sheet2 = st.sidebar.selectbox("Select Sheet in file2 to check", sheet_names2)

        # Display selected sheets
        st.sidebar.write(f"Sheet in file1 to check: {sheet1}")
        st.sidebar.write(f"Sheet in file2 to check: {sheet2}")

        # Read selected sheets into dataframes
        df1 = pd.read_excel(excel_file1, sheet1)
        df2 = pd.read_excel(excel_file2, sheet2)

        # Get selected columns for comparison
        selected_columns_df1 = st.multiselect(f"Select columns for comparison in {sheet1}", df1.columns, key="df1")
        selected_columns_df2 = st.multiselect(f"Select columns for comparison in {sheet2}", df2.columns, key="df2")
        if selected_columns_df2 and selected_columns_df2:
        # Find rows present in sheet1 but not in sheet2
            rows_only_in_sheet1 = find_rows_only_in_sheet(df1[selected_columns_df1], df2[selected_columns_df2])

            # Find rows present in sheet2 but not in sheet1
            rows_only_in_sheet2 = find_rows_only_in_sheet(df2[selected_columns_df2], df1[selected_columns_df1])

            # Display the results
            st.header(f"Rows present in {sheet1} but not in {sheet2}:")
            st.dataframe(rows_only_in_sheet1)

            st.header(f"Rows present in {sheet2} but not in {sheet1}:")
            st.dataframe(rows_only_in_sheet2)

            # Export non-matching rows to Excel
            if st.button("Export Non-Matching Rows"):
                export_non_matching_rows(rows_only_in_sheet1, f"{sheet1}_not_in_{sheet2}")
                export_non_matching_rows(rows_only_in_sheet2, f"{sheet2}_not_in_{sheet1}")

def find_rows_only_in_sheet(sheet1, sheet2):
    # Find rows present in sheet1 but not in sheet2
    merged_df = pd.merge(sheet1, sheet2, how='left', indicator=True)
    rows_only_in_sheet = merged_df[merged_df['_merge'] == 'left_only'].drop('_merge', axis=1)
    return rows_only_in_sheet

def export_non_matching_rows(non_matching_rows, output_filename):
    # Save non-matching rows to Excel
    output_filename += ".xlsx"
    non_matching_rows.to_excel(output_filename, index=False)
    st.success(f"Non-matching rows exported to {output_filename}")

if __name__ == "__main__":
    main()
