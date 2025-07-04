import streamlit as st
import pandas as pd
import os
import re
import tempfile
import pdfplumber
from io import BytesIO

st.set_page_config(page_title="BOE Extractor Tool", layout="wide")
st.title("📄 PDF to Excel BOE Extractor")

uploaded_files = st.file_uploader("Upload PDF files", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    output_summary = []

    with st.spinner("Preparing files..."):

        # Temp folders
        temp_dir = tempfile.TemporaryDirectory()
        pdf_folder = os.path.join(temp_dir.name, "PDF")
        os.makedirs(pdf_folder, exist_ok=True)

        format1_results = []
        format2_results = []

        # Save PDFs to temp dir
        for file in uploaded_files:
            with open(os.path.join(pdf_folder, file.name), "wb") as f:
                f.write(file.read())

        def extract_tables_from_pdf(pdf_path):
            tables = []
            with pdfplumber.open(pdf_path) as pdf:
                for i, page in enumerate(pdf.pages):
                    page_tables = page.extract_tables()
                    for table in page_tables:
                        df = pd.DataFrame(table)
                        df['Page'] = i + 1
                        tables.append(df)
            return tables

        def process_format1(file_path):
            try:
                xls = pd.ExcelFile(file_path, engine='openpyxl')
                df1 = pd.read_excel(file_path, sheet_name=0, header=None, engine='openpyxl')
            except:
                return []

            be_no = hawb = country = None

            for row in range(len(df1) - 2):
                for col in range(len(df1.columns)):
                    if str(df1.iat[row, col]).strip() == "BE No":
                        be_no = df1.iat[row + 2, col]
                        break

            for row in range(len(df1)):
                for col in range(len(df1.columns) - 5):
                    if str(df1.iat[row, col]).strip() == "13.COUNTRY OF ORIGIN":
                        country = df1.iat[row, col + 5]
                        break

            for row in range(len(df1) - 1):
                for col in range(len(df1.columns)):
                    if str(df1.iat[row, col]).strip() == "8.HAWB NO":
                        hawb = df1.iat[row + 1, col]
                        break

            results = []

            for sheet_name in xls.sheet_names:
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    found = any("INVOICE & VALUATION DETAILS" in str(cell).upper() for cell in df.values.flatten())
                    if not found:
                        continue

                    invoice = None
                    for row in range(len(df) - 1):
                        for col in range(len(df.columns)):
                            if str(df.iat[row, col]).strip().upper() == "2.INVOICE NO. & DT.":
                                invoice = df.iat[row + 1, col]
                                break
                        if invoice:
                            break

                    desc = None
                    for row in range(len(df) - 1):
                        for col in range(len(df.columns)):
                            if "DESCRIPTION" in str(df.iat[row, col]).strip().upper():
                                val = df.iat[row + 1, col]
                                val = re.sub(r'^\d+\s*', '', str(val))
                                desc = re.sub(r'^S[\n\s]+', '', val)
                                break
                        if desc:
                            break

                    # Extract supplier name & address from 2nd sheet
                    supplier = supplier_address = None
                    if sheet_name == xls.sheet_names[1]:  # 2nd sheet (index 1)
                        for row in range(len(df) - 2):  # -2 to allow 2 lines after
                            for col in range(len(df.columns)):
                                if str(df.iat[row, col]).strip().upper() == "3.SUPPLIER NAME & ADDRESS":
                                    supplier = df.iat[row + 1, col]
                                    supplier_address = df.iat[row + 2, col] if row + 2 < len(df) else 'Not found'
                                    break
                            if supplier:
                                break

                    results.append({
                        "BE No": be_no,
                        "Invoice No": invoice,
                        "ORIGIN COUNTRY*": country,
                        "GOODS DESCRIPTION": desc,
                        "HAWB NO": hawb,
                        "SUPPLIER NAME": supplier if supplier else 'Not found',
                        "SUPPLIER ADDRESS": supplier_address if supplier_address else 'Not found'
                    })
                except:
                    continue
            return results

        def process_format2(file_path):
            try:
                xls = pd.ExcelFile(file_path, engine='openpyxl')
            except:
                return []

            be_no = hawb = origin = invoice = None
            invoice_found = False
            descriptions = []

            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, header=None)
                all_cells = df.astype(str).values.flatten()
                if any('DetailsOfInvoice-' in c for c in all_cells):
                    invoice_found = True
                    break

            for sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, header=None)
                for _, row in df.iterrows():
                    for i, val in enumerate(row.values):
                        if isinstance(val, str):
                            if 'CBEXIV_' in val:
                                m = re.search(r'_(\d+)$', val)
                                if m:
                                    be_no = '14' + m.group(1)
                            elif 'HouseAirwayBill(HAWB)' in val and i + 1 < len(row):
                                hawb = row[i + 1]
                            elif 'CountryofOrigin:' in val and i + 1 < len(row):
                                origin = row[i + 1]
                            elif invoice_found and 'InvoiceNumber:' in val and i + 1 < len(row):
                                if not invoice:
                                    invoice = row[i + 1]
                            elif invoice_found and 'ItemDescription:' in val and i + 1 < len(row):
                                cleaned = re.sub(r'^\d+\s*', '', str(row[i + 1]))
                                descriptions.append(cleaned)

            supplier = supplier_address = None
            if len(xls.sheet_names) > 2:
                df_supplier = pd.read_excel(xls, sheet_name=xls.sheet_names[2], header=None)
                
                # Find supplier name as before
                for row in range(len(df_supplier) - 2):
                    for col in range(len(df_supplier.columns) - 1):
                        if str(df_supplier.iat[row, col]).strip().lower() == "name:":
                            supplier = df_supplier.iat[row, col + 1]
                            break
                    if supplier:
                        break

                # Find supplier address by searching for "Address:"
                address_row = address_col = None
                for r in range(len(df_supplier)):
                    for c in range(len(df_supplier.columns)):
                        if str(df_supplier.iat[r, c]).strip().lower() == "address:":
                            address_row, address_col = r, c
                            break
                    if address_row is not None:
                        break

                if address_row is not None and address_col + 1 < len(df_supplier.columns):
                    supplier_address = df_supplier.iat[address_row, address_col + 1]
                else:
                    supplier_address = 'Not found'


            return [{
                'BE No': be_no if be_no else 'Not found',
                'Invoice No': invoice if invoice else ('Not found' if invoice_found else 'N/A'),
                'ORIGIN COUNTRY*': origin if origin else 'Not found',
                'GOODS DESCRIPTION': '; '.join(descriptions) if descriptions else ('Not found' if invoice_found else 'N/A'),
                'HAWB NO': hawb if hawb else 'Not found',
                'SUPPLIER NAME': supplier if supplier else 'Not found',
                'SUPPLIER ADDRESS': supplier_address if supplier_address else 'Not found'
            }]

        file_list = os.listdir(pdf_folder)
        total_files = len(file_list)
        progress_bar = st.progress(0)
        status_text = st.empty()

        for i, filename in enumerate(file_list):
            pdf_path = os.path.join(pdf_folder, filename)
            status_text.text(f"📄 Processing file: {filename} ({i+1}/{total_files})")

            tables = extract_tables_from_pdf(pdf_path)

            excel_buf = BytesIO()
            with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
                for idx, df in enumerate(tables):
                    df.to_excel(writer, sheet_name=f"Table_{idx+1}", index=False)

            excel_buf.seek(0)

            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_xlsx:
                temp_xlsx.write(excel_buf.read())
                temp_xlsx_path = temp_xlsx.name

            if "_" in filename:
                format2_results.extend(process_format2(temp_xlsx_path))
            else:
                format1_results.extend(process_format1(temp_xlsx_path))

            progress_bar.progress((i + 1) / total_files)

        progress_bar.empty()
        status_text.text("✅ All files processed!")

        df_combined = pd.DataFrame(format1_results + format2_results)

        def agg_func(series):
            if series.dtype == object:
                return "; ".join(series.dropna().astype(str).unique())
            else:
                return series.iloc[0]

        df_grouped = df_combined.groupby('Invoice No', dropna=False).agg(agg_func).reset_index()

        st.success("🎉 Processing complete!")
        st.dataframe(df_grouped)

        output = BytesIO()
        df_grouped.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="📥 Download in Excel",
            data=output,
            file_name="Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
