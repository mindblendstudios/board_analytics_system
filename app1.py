import streamlit as st
import pandas as pd
from io import BytesIO

# ------------------------
# Sample Excel generation
# ------------------------
def generate_sample_excel():
    data = {
        'Year': [2021, 2022, 2023],
        'Sales': [1000000, 1200000, 1350000],
        'Operating_Income': [200000, 240000, 270000],
        'Operating_Margin': [0.20, 0.20, 0.20],
        'Operating_Depreciation_Amortization': [50000, 60000, 65000],
        'Non_Operating_Costs': [10000, 12000, 13000],
        'Operating_Cash_Flow': [180000, 210000, 230000],
        'Free_Cash_Flow': [150000, 175000, 190000],
        'EBIT': [200000, 240000, 270000],
        'Capital_Employed': [1000000, 1100000, 1150000],
    }
    df = pd.DataFrame(data)
    return df

def get_excel_download_link(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='SampleData')
    buffer.seek(0)
    return buffer

# ------------------------
# KPI Calculation
# ------------------------
def calculate_kpis(df):
    # Normalize column names
    df.columns = df.columns.str.strip().str.lower()

    # Dynamic calculations only if required columns exist
    if all(col in df.columns for col in ['operating_income', 'operating_depreciation_amortization', 'non_operating_costs']):
        df['ebitda'] = df['operating_income'] + df['operating_depreciation_amortization'] + df['non_operating_costs']

    if 'free_cash_flow' in df.columns and 'operating_cash_flow' in df.columns:
        df['free_cash_flow_conversion'] = (df['free_cash_flow'] / df['operating_cash_flow']) * 100

    if 'ebit' in df.columns and 'capital_employed' in df.columns:
        df['roce'] = (df['ebit'] / df['capital_employed']) * 100

    return df

# ------------------------
# Main Streamlit App
# ------------------------
def main():
    st.title("📊 Company Financial Report KPI Calculation & Analysis")

    # Sample Excel download
    st.subheader("⬇ Download Sample Excel Template")
    sample_df = generate_sample_excel()
    sample_excel = get_excel_download_link(sample_df)
    st.download_button(
        label="Download Sample Excel File",
        data=sample_excel,
        file_name="sample_kpi_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # File upload
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        # Normalize column names
        df.columns = df.columns.str.strip().str.lower()

        st.write("### Uploaded Data Preview")
        st.dataframe(df)

        df_kpis = calculate_kpis(df)

        st.write("### Calculated KPIs")
        st.dataframe(df_kpis)

        # Dynamic KPI charting
        numeric_cols = df_kpis.select_dtypes(include='number').columns.tolist()
        if len(numeric_cols) > 1:
            x_col = st.selectbox("Select X-axis column", options=numeric_cols, index=0)
            y_col = st.selectbox("Select KPI to plot", options=[c for c in numeric_cols if c != x_col])
            st.line_chart(df_kpis.set_index(x_col)[y_col])

if __name__ == "__main__":
    main()
