import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile

st.set_page_config(layout="wide")

st.title("Excel Processor")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

st.caption("Required columns: HUB, Location, Zone/COC, Owner, Customer Code, Customer Name, Order No, Invoice No, WF_TaskID, Shap Hrs., Performed Hrs, Billed Hrs, Variance, Excess Paid, Excess billing, Short billing, Short / Missing Roster, Training & OJT, Complimentary Hrs.")

if st.button("Run"):
    if uploaded_file is not None:
        st.write("Reading file...")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.read())
            file_path = tmp.name

        Bangalore = pd.read_excel(file_path, sheet_name='Bangalore', header=1)
        Chennai = pd.read_excel(file_path, sheet_name='Chennai', header=1)
        Hyderabad = pd.read_excel(file_path, sheet_name='Hyderabad', header=1)
        Mumbai = pd.read_excel(file_path, sheet_name='Mumbai', header=1)
        Kolkata = pd.read_excel(file_path, sheet_name='Kolkata', header=1)
        NCR = pd.read_excel(file_path, sheet_name='NCR', header=1)

        dataframes = {
            'Bangalore': Bangalore,
            'Chennai': Chennai,
            'Hyderabad': Hyderabad,
            'Mumbai': Mumbai,
            'Kolkata': Kolkata,
            'NCR': NCR
        }

        desired_columns = [
            'HUB','Location','Zone/COC','Owner','Customer Code','Customer Name',
            'Order No','Invoice No','WF_TaskID','Shap Hrs.','Performed Hrs',
            'Billed Hrs','Variance','Excess Paid','Excess billing','Short billing',
            'Short / Missing Roster','Training & OJT','Complimentary Hrs.'
        ]

        desired_columns = [col.lower() for col in desired_columns]

        st.write("Processing sheets...")

        for name, df in dataframes.items():
            df.columns = df.columns.str.strip().str.lower()

            existing_cols = [col for col in desired_columns if col in df.columns]
            df = df[existing_cols]

            df.columns = df.columns.str.title()
            df = df.copy()

            df['Customer Code'] = (
                df['Customer Code']
                .fillna('')
                .astype(str)
                .str.replace(r'\.0$', '', regex=True)
                .str.strip()
            )

            df = df[df['Customer Code'] != '']
            df.loc[:, 'Loc'] = None
            df.loc[df.index[0], 'Loc'] = name

            cols = ['Loc'] + [col for col in df.columns if col != 'Loc']
            df = df[cols]

            

            cols_to_fix = cols_to_fix = [ 'Excess Paid', 'Excess Billing', 'Short Billing', 'Short / Missing Roster', 'Training & Ojt', 'Complimentary Hrs.' ]
            df[cols_to_fix] = df[cols_to_fix].fillna(0).astype(int)

            # Your existing filter line
            df = df[df[cols_to_fix].sum(axis=1) == 0]
            # 1. Convert back to object/string so they can hold blanks
            df[cols_to_fix] = df[cols_to_fix].astype(str)
            
            # 2. Replace the '0' strings with actual blanks
            for col in cols_to_fix:
                df[col] = df[col].astype(str).replace(['0', '0.0', 'nan', 'None'], '')
            dataframes[name] = df

        st.write("Combining data...")

        all_cities_df = pd.concat(dataframes.values(), ignore_index=True)

        st.write("Writing output...")

        output_path = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name

        wb = load_workbook(file_path)
        ws = wb.create_sheet(title='All Cities')

        for r_idx, row in enumerate(all_cities_df.itertuples(index=False), start=2):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        for c_idx, col_name in enumerate(all_cities_df.columns, start=1):
            ws.cell(row=1, column=c_idx, value=col_name)

        wb.save(output_path)

        st.write("Done")

        with open(output_path, "rb") as f:
            st.download_button("Download Output", f, file_name="final_output.xlsx")

    else:
        st.write("Please upload a file")
