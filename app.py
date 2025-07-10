import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(page_title="Driver-Based Forecasting", layout="wide")

st.title("📈 Driver-Based Financial Forecasting Tool")
st.markdown("Upload an Excel file with assumptions for Base, Best, and Worst scenarios.")

# File uploader
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

def calculate_forecast_from_row(row):
    revenue_local = row['Units Sold'] * row['Price per Unit']
    revenue_group = revenue_local * row['FX Rate']
    cogs = revenue_local * row['COGS %']
    gross_margin = revenue_local - cogs
    operating_profit = gross_margin - row['Operating Expenses']
    gross_margin_pct = gross_margin / revenue_local if revenue_local else 0
    operating_margin_pct = operating_profit / revenue_local if revenue_local else 0

    return {
        'Scenario': row['Scenario'],
        'Revenue (Local)': revenue_local,
        'Revenue (Group)': revenue_group,
        'COGS': cogs,
        'Gross Margin': gross_margin,
        'Gross Margin %': gross_margin_pct * 100,
        'Operating Profit (EBIT)': operating_profit,
        'Operating Margin %': operating_margin_pct * 100,
        'Net Income': operating_profit
    }

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        forecast_results = [calculate_forecast_from_row(row) for _, row in df.iterrows()]
        results_df = pd.DataFrame(forecast_results).set_index('Scenario')

        # Formatting for display
        display_df = results_df.copy()
        for col in display_df.columns:
            if '%' in col:
                display_df[col] = display_df[col].map('{:.1f}%'.format)
            else:
                display_df[col] = display_df[col].map('{:,.2f}'.format)

        # Sidebar: Scenario dropdown
        st.sidebar.title("Scenario Selection")
        scenario_selected = st.sidebar.selectbox("Choose a Scenario", results_df.index.tolist())

        # Charts
        st.subheader(f"📊 Financial Overview - {scenario_selected} Scenario")
        selected_row = results_df.loc[scenario_selected]

        fig = go.Figure()
        fig.add_trace(go.Bar(x=['Revenue (Local)', 'Revenue (Group)', 'COGS',
                                'Gross Margin', 'Operating Profit (EBIT)', 'Net Income'],
                             y=[selected_row['Revenue (Local)'], selected_row['Revenue (Group)'], selected_row['COGS'],
                                selected_row['Gross Margin'], selected_row['Operating Profit (EBIT)'], selected_row['Net Income']],
                             marker_color='teal'))

        fig.update_layout(title=f"Financial Metrics for {scenario_selected}", yaxis_title="Amount", template="plotly_white")
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("📋 Full Forecast Table")
        st.dataframe(display_df)

        # Download button
        def convert_df_to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer)
            processed_data = output.getvalue()
            return processed_data

        excel_download = convert_df_to_excel(results_df.reset_index())
        st.download_button(label="📥 Download Results as Excel",
                           data=excel_download,
                           file_name='forecast_results.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("Please upload a valid Excel file with columns: Scenario, Units Sold, Price per Unit, FX Rate, COGS %, Operating Expenses.")
