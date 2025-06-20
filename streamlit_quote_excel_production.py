import streamlit as st
import pandas as pd
import numpy as np
import io

# --- DATA CLEANING FUNCTION ---
def process_data(file):
    df = pd.read_csv(file)
    df = df[['Quote#','QuoteDate','Customer','PartNumber','Qty','Price','Extension','Margin%','MatlCode','V','LTA Flag']]

    df['Price'] = df['Price'].replace({'\$': '', ',': ''}, regex=True)
    df['Extension'] = df['Extension'].replace({'\$': '', ',': ''}, regex=True)
    df['Qty'] = df['Qty'].replace({',': ''}, regex=True)
    df['Qty'] = pd.to_numeric(df['Qty'])
    df['Margin'] = pd.to_numeric(df['Margin%']) / 100
    df.drop(columns=['Margin%'], inplace=True)
    df['Price'] = pd.to_numeric(df['Price'])
    df['Extension'] = pd.to_numeric(df['Extension'])
    df['Cost'] = df['Price'] * (1 - df['Margin'])
    return df

# --- STANDARD EXCEL EXPORT ---
def create_standard_excel(df, title_text):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    worksheet = workbook.add_worksheet()
    writer.sheets['Sheet1'] = worksheet

    worksheet.hide_gridlines(2)

    # Formats
    boldfill = workbook.add_format({'bold': True, 'bg_color': '#CCECFF', 'border': 1})
    header = workbook.add_format({'bold': True, 'border': 1})
    currency_format = workbook.add_format({'num_format': '$#,##0.00'})
    percent_format = workbook.add_format({'num_format': '0%'})
    fill_currency = workbook.add_format({'bg_color': '#CCECFF', 'border': 1, 'num_format': '$#,##0.00'})
    fill_percent = workbook.add_format({'bg_color': '#CCECFF', 'border': 1, 'num_format': '0%'})

    worksheet.write('A1', title_text, boldfill)

    start_row = 3
    start_col = 0
    df.to_excel(writer, sheet_name='Sheet1', startrow=start_row, startcol=start_col, index=False)

    cost_col = df.columns.get_loc('Cost') + start_col + 1
    price_col = df.columns.get_loc('Price') + start_col + 1
    qty_col = df.columns.get_loc('Qty') + start_col + 1
    total_cost = len(df.columns) + start_col
    total_price = total_cost + 1
    margin_col = df.columns.get_loc('Margin') + start_col + 1
    ext_col = df.columns.get_loc("Extension") + start_col

    worksheet.set_column(cost_col - 1, cost_col - 1, None, currency_format)
    worksheet.set_column(price_col - 1, price_col - 1, None, currency_format)
    worksheet.set_column(total_cost, total_cost, None, currency_format)
    worksheet.set_column(total_price, total_price, None, currency_format)
    worksheet.set_column(ext_col, ext_col, None, currency_format)
    worksheet.set_column(margin_col - 1, margin_col - 1, None, percent_format)

    worksheet.write(start_row, total_cost, 'Total Cost', header)
    worksheet.write(start_row, total_price, 'Total Price', header)

    for i, row in enumerate(range(start_row + 1, start_row + 1 + len(df))):
        worksheet.write_formula(
            row, price_col-1,
            f'=${chr(64+cost_col)}{row + 1}/(1 - ${chr(64+margin_col)}{row + 1})'
        )
        worksheet.write_formula(
            row, ext_col,
            f'=${chr(64+price_col)}{row + 1}*${chr(64+qty_col)}{row + 1}'
        )
        worksheet.write_formula(
            row, total_cost,
            f'=${chr(64+cost_col)}{row + 1}*${chr(64+qty_col)}{row + 1}'
        )
        worksheet.write_formula(
            row, total_price,
            f'=${chr(64+price_col)}{row + 1}*${chr(64+qty_col)}{row + 1}'
        )

    ext_cost_formula = f'=SUBTOTAL(9, ${chr(64+total_cost + 1)}{start_row + 2}:${chr(64+total_cost + 1)}{start_row + 1 + len(df)})'
    ext_price_formula = f'=SUBTOTAL(9, ${chr(64+total_price + 1)}{start_row + 2}:${chr(64+total_price + 1)}{start_row + 1 + len(df)})'
    gm_formula = f'=(R2-Q2)/R2'

    worksheet.write('Q1', 'Ext Cost', boldfill)
    worksheet.write_formula('Q2', ext_cost_formula, fill_currency)
    worksheet.write('R1', 'Ext Price', boldfill)
    worksheet.write_formula('R2', ext_price_formula, fill_currency)
    worksheet.write('S1', 'GM', boldfill)
    worksheet.write_formula('S2', gm_formula, fill_percent)

    # Summary sheet
    summary = workbook.add_worksheet('summary')
    border = workbook.add_format({'border': 2})
    boldborderfill = workbook.add_format({'bold': True, 'bg_color': '#DDD9C4', 'border': 2})

    summary.write('A1', 'Customer Name', border)
    summary.write('A2', 'Item Types', border)
    summary.write('A3', 'Bid Type: (AdHoc / LTA/ Renewal)', border)
    summary.write('A4', 'Existing or New business', border)
    summary.write('A5', 'Total Lines (Individual P/Ns)', border)
    summary.write('A6', 'Value', border)
    summary.write('A7', 'Gross Margin', border)
    summary.write('A8', 'Price Protection', border)
    summary.write('A9', 'Quote IDs', border)

    for i in range(1, 10):
        summary.write(f'B{i}', '', boldborderfill if i == 1 else border)

    writer.close()
    output.seek(0)
    return output

# --- INFLATION EXCEL EXPORT ---
def create_inflation_excel(df, title_text, compound_periods, rate):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    worksheet = workbook.add_worksheet()
    writer.sheets['Sheet1'] = worksheet

    worksheet.hide_gridlines(2)

    # Formats
    boldfill = workbook.add_format({'bold': True, 'bg_color': '#CCECFF', 'border': 1})
    fill = workbook.add_format({'bg_color': '#CCECFF', 'border': 1})
    header = workbook.add_format({'bold': True, 'border': 1})
    currency_format = workbook.add_format({'num_format': '$#,##0.00'})
    percent_format = workbook.add_format({'num_format': '0%'})
    fill_currency = workbook.add_format({'bg_color': '#CCECFF', 'border': 1, 'num_format': '$#,##0.00'})
    fill_percent = workbook.add_format({'bg_color': '#CCECFF', 'border': 1, 'num_format': '0%'})

    worksheet.write('A1', title_text, boldfill)
    start_row = 3
    start_col = 0
    df.to_excel(writer, sheet_name='Sheet1', startrow=start_row, index=False)

    # Add formulas for "Future Cost" and "Future Price" in terms of 0 based column numbers
    cost_col = df.columns.get_loc('Cost') + start_col
    price_col = df.columns.get_loc('Price') + start_col
    qty_col = df.columns.get_loc('Qty') + start_col
    future_cost_formula_col = len(df.columns) + start_col
    future_price_formula_col = future_cost_formula_col + 1
    cost_ext_formula_col = future_price_formula_col + 1
    price_ext_formula_col = cost_ext_formula_col + 1
    today_margin_formula_col = price_ext_formula_col + 1
    margin_col = df.columns.get_loc('Margin') + start_col
    ext_col = df.columns.get_loc("Extension") + start_col


    #currency formats
    worksheet.set_column(cost_col, cost_col, None, currency_format)
    worksheet.set_column(price_col, price_col, None, currency_format)
    worksheet.set_column(future_cost_formula_col, future_cost_formula_col, None, currency_format)
    worksheet.set_column(future_price_formula_col, future_price_formula_col, None, currency_format)
    worksheet.set_column(ext_col, ext_col, None, currency_format)
    worksheet.set_column(cost_ext_formula_col, cost_ext_formula_col, None, currency_format)
    worksheet.set_column(price_ext_formula_col, price_ext_formula_col, None, currency_format)

    #percentage formats
    worksheet.set_column(margin_col, margin_col, None, percent_format)
    worksheet.set_column(today_margin_formula_col, today_margin_formula_col, None, percent_format)


    worksheet.write('N1', 'Inflation Periods', boldfill)
    worksheet.write('O1', compound_periods, fill)
    worksheet.write('N2', 'Rate', boldfill)
    worksheet.write('O2', rate, fill_percent)

    #new columns
    worksheet.write(start_row, future_cost_formula_col, 'Future Cost', header)
    worksheet.write(start_row, future_price_formula_col, 'Future Price', header)
    worksheet.write(start_row, cost_ext_formula_col, 'Cost Ext', header)
    worksheet.write(start_row, price_ext_formula_col, 'Price Ext', header)
    worksheet.write(start_row, today_margin_formula_col, "Today's Margin", header)


    for i, row in enumerate(range(start_row + 1, start_row + 1 + len(df))):  # Start writing formulas after headers
        worksheet.write_formula(
            row, price_col,
            f'=${chr(65+cost_col)}{row + 1}/(1 - ${chr(65+margin_col)}{row + 1})'  
        )
        worksheet.write_formula(
            row, ext_col,
            f'=${chr(65+price_col)}{row + 1}*${chr(65+qty_col)}{row + 1}'  
        )
        worksheet.write_formula(
            row, future_cost_formula_col,
            f'=${chr(65+cost_col)}{row + 1}*(1+$O$2)^$O$1'  
        )
        worksheet.write_formula(
            row, future_price_formula_col,
            f'=${chr(65+price_col)}{row + 1}*(1+$O$2)^$O$1' 
        )
        worksheet.write_formula(
            row, cost_ext_formula_col,
            f'=${chr(65+future_cost_formula_col)}{row + 1}*${chr(65+qty_col)}{row + 1}'  
        )
        worksheet.write_formula(
            row, price_ext_formula_col,
            f'=${chr(65+future_price_formula_col)}{row + 1}*${chr(65+qty_col)}{row + 1}'  
        )
        worksheet.write_formula(
            row, today_margin_formula_col,
            f'=(${chr(65+future_price_formula_col)}{row + 1}-${chr(65+cost_col)}{row + 1})/${chr(65+future_price_formula_col)}{row + 1}'  
        )


    # Total and GM
    ext_cost_formula = f'=SUBTOTAL(9, ${chr(64+cost_ext_formula_col + 1)}{start_row + 2}:${chr(64+cost_ext_formula_col + 1)}{start_row + 1 + len(df)})'
    ext_price_formula = f'=SUBTOTAL(9, ${chr(64+price_ext_formula_col + 1)}{start_row + 2}:${chr(64+price_ext_formula_col + 1)}{start_row + 1 + len(df)})'
    gm_formula = f'=(R2-Q2)/R2'

    worksheet.write('Q1', 'Ext Cost', boldfill)
    worksheet.write_formula('Q2', ext_cost_formula, fill_currency)

    worksheet.write('R1', 'Ext Price', boldfill)
    worksheet.write_formula('R2', ext_price_formula, fill_currency)

    worksheet.write('S1', 'GM', boldfill)
    worksheet.write_formula('S2', gm_formula, fill_percent)


    # Summary Sheet
    summary = workbook.add_worksheet("summary")
    border = workbook.add_format({'border': 2})
    boldborderfill = workbook.add_format({'bold': True, 'bg_color': '#DDD9C4', 'border': 2})

    for i, label in enumerate([
        'Customer Name', 'Item Types', 'Bid Type: (AdHoc / LTA/ Renewal)', 'Existing or New business',
        'Total Lines (Individual P/Ns)', 'Value', 'Gross Margin', 'Price Protection', 'Quote IDs'
    ], 1):
        summary.write(f'A{i}', label, border)
        summary.write(f'B{i}', '', boldborderfill if i == 1 else border)

    writer.close()
    output.seek(0)
    return output


# --- STREAMLIT UI ---
st.title("Quote Summary Processor")

# Step 1: Upload CSV file
uploaded_file = st.file_uploader("Upload CSV File", type=["csv"])

# Only show other options if file is uploaded
if uploaded_file:
    st.success("File uploaded successfully!")

    # Step 2: Additional inputs appear AFTER file upload
    title_text = st.text_input("Enter a Title (this will be placed in cell A1)", value="")
    file_name = st.text_input("Name the output file", value="")

    # Step 3: Excel format option
    excel_type = st.radio("Choose Excel Format:", ["Standard", "Inflation-adjusted"])

    compound_periods = None
    rate = None

    # Step 4: If Inflation-adjusted is selected, show extra inputs
    if excel_type == "Inflation-adjusted":
        compound_periods_input = st.text_input("Compound Periods", value="")
        rate_input = st.text_input("Annual Rate (e.g. 0.04 for 4%)", value="")

        try:
            compound_periods = int(compound_periods_input) if compound_periods_input else None
        except ValueError:
            st.error("Compound Periods must be a whole number.")

        if rate_input:
            if '.' not in rate_input:
                st.error("Rate must be a decimal (e.g., 0.04 for 4%). Whole numbers are not allowed.")
            else:
                try:
                    rate_value = float(rate_input)
                    if 0 < rate_value < 1:
                        rate = rate_value
                    else:
                        st.error("Rate must be between 0 and 1 (exclusive).")
                except ValueError:
                    st.error("Rate must be a valid decimal number (e.g., 0.04 for 4%).")

    # Step 5: Run processing + download logic if all fields are filled
    if file_name.strip() and title_text.strip():
        df = process_data(uploaded_file)
        st.subheader("Data Preview")
        st.dataframe(df.head())

        if excel_type == "Standard":
            excel_bytes = create_standard_excel(df, title_text)
        else:
            if compound_periods is None or rate is None:
                st.warning("Please provide valid values for Compound Periods and Annual Rate.")
                excel_bytes = None
            else:
                excel_bytes = create_inflation_excel(df, title_text, compound_periods, rate)

        if excel_bytes:
            st.download_button(
                label="Download Excel File",
                data=excel_bytes,
                file_name=f"{file_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("Please enter both a title and a file name.")
