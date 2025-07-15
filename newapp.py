import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from io import BytesIO

def get_sample_template():
    data = {
        'Scenario': ['Base', 'Base', 'Best', 'Worst'],
        'Product': ['Product A', 'Product B', 'Product C', 'Product D'],
        'Region': ['Asia', 'Europe', 'North America', 'South America'],
        'Year': [2024, 2024, 2024, 2024],
        'Units Sold': [1000, 1200, 800, 700],
        'Price per Unit': [25, 27, 20, 22],
        'FX Rate': [1.0, 1.1, 0.9, 1.2],
        'COGS %': [0.60, 0.65, 0.62, 0.66],
        'Operating Expenses': [20000, 22000, 18000, 17000],
        'Depreciation': [2500, 2600, 2400, 2200],
        'Tax Rate': [0.25, 0.25, 0.25, 0.25],
    }
    return pd.DataFrame(data)


st.set_page_config(page_title="Driver-Based Financial Forecasting", layout="wide")

st.title("📈 Driver-Based Financial Forecasting Tool")
st.markdown(
    "Upload an Excel file with assumptions for Base, Best, and Worst scenarios, including Product, Region, and Year breakdown. You can also override key assumptions manually."
)

# ---- Download Sample Template Button ----
st.markdown("📋 **Need a template? Download the sample Excel file below:**")
sample_df = get_sample_template()
excel_io = BytesIO()
with pd.ExcelWriter(excel_io) as writer:
    sample_df.to_excel(writer, index=False)
excel_io.seek(0)

st.download_button(
    label="📥 Download Sample Excel Template",
    data=excel_io,
    file_name="driver_forecast_template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


st.markdown("---")

# File uploader
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

def calculate_forecast_from_row(row):
    revenue_local = row['Units Sold'] * row['Price per Unit']
    revenue_group = revenue_local * row['FX Rate']
    cogs = revenue_local * row['COGS %']
    gross_margin = revenue_local - cogs
    ebitda = gross_margin - row['Operating Expenses'] + row.get('Depreciation', 0)
    operating_profit = ebitda - row.get('Depreciation', 0)
    tax = operating_profit * row.get('Tax Rate', 0.25)
    net_income = operating_profit - tax

    gross_margin_pct = gross_margin / revenue_local if revenue_local else 0
    operating_margin_pct = operating_profit / revenue_local if revenue_local else 0
    net_margin_pct = net_income / revenue_local if revenue_local else 0

    cash_flow = net_income + row.get('Depreciation', 0)

    return {
        'Scenario': row['Scenario'],
        'Product': row['Product'],
        'Region': row.get('Region', 'Unknown'),
        'Year': row.get('Year', 'N/A'),
        'Revenue (Local)': revenue_local,
        'Revenue (Group)': revenue_group,
        'COGS': cogs,
        'Gross Margin': gross_margin,
        'Gross Margin %': gross_margin_pct * 100,
        'EBITDA': ebitda,
        'Operating Profit (EBIT)': operating_profit,
        'Operating Margin %': operating_margin_pct * 100,
        'Tax': tax,
        'Net Income': net_income,
        'Net Margin %': net_margin_pct * 100,
        'Cash Flow': cash_flow,
        'Operating Expenses': row['Operating Expenses'],
        'Depreciation': row['Depreciation']
    }

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()  # Remove hidden spaces

        # Add missing optional columns if needed
        if 'Tax Rate' not in df.columns:
            df['Tax Rate'] = 0.25
        if 'Depreciation' not in df.columns:
            df['Depreciation'] = 0

        # User overrides
        st.sidebar.markdown("### Optional Overrides")
        override_units = st.sidebar.number_input("Override Units Sold", min_value=0, value=None, step=100)
        override_price = st.sidebar.number_input("Override Price per Unit", min_value=0.0, value=None, step=0.1)
        override_fx = st.sidebar.number_input("Override FX Rate", min_value=0.0, value=None, step=0.01)

        if override_units:
            df['Units Sold'] = override_units
        if override_price:
            df['Price per Unit'] = override_price
        if override_fx:
            df['FX Rate'] = override_fx

        forecast_results = [calculate_forecast_from_row(row) for _, row in df.iterrows()]
        results_df = pd.DataFrame(forecast_results)
        results_df.columns = results_df.columns.str.strip()

        st.sidebar.title("Scenario & Filter Options")
        scenario_selected = st.sidebar.selectbox("Choose a Scenario", sorted(results_df['Scenario'].unique()))
        region_selected = st.sidebar.multiselect("Filter by Region", sorted(results_df['Region'].unique()), default=sorted(results_df['Region'].unique()))
        year_selected = st.sidebar.multiselect("Filter by Year", sorted(results_df['Year'].unique()), default=sorted(results_df['Year'].unique()))

        filtered_df = results_df[
            (results_df['Scenario'] == scenario_selected) &
            (results_df['Region'].isin(region_selected)) &
            (results_df['Year'].isin(year_selected))
        ]

        for col in ['Revenue (Group)', 'COGS', 'Operating Expenses', 'Depreciation', 'Tax']:
            if col not in filtered_df.columns:
                filtered_df[col] = 0

        agg = filtered_df[['Revenue (Group)', 'COGS', 'Operating Expenses', 'Depreciation', 'Tax']].sum()
        operating_profit = agg['Revenue (Group)'] - agg['COGS'] - agg['Operating Expenses']
        net_income = operating_profit - agg['Tax']

        st.subheader(f"📊 Revenue by Product and Region - {scenario_selected} Scenario")
        fig_bar = go.Figure()
        for region in filtered_df['Region'].unique():
            region_df = filtered_df[filtered_df['Region'] == region].groupby('Product')['Revenue (Group)'].sum().reset_index()
            fig_bar.add_trace(go.Bar(x=region_df['Product'], y=region_df['Revenue (Group)'], name=region))
        fig_bar.update_layout(barmode='stack', title="Revenue by Product (Stacked by Region)", yaxis_title="Amount", template="plotly_white")
        st.plotly_chart(fig_bar, use_container_width=True)

        st.subheader(f"🧁 Revenue Split by Product - {scenario_selected} Scenario")
        fig_pie = go.Figure(data=[go.Pie(labels=filtered_df['Product'], values=filtered_df['Revenue (Group)'], hole=0.3)])
        fig_pie.update_layout(title="Revenue Split by Product", template="plotly_white")
        st.plotly_chart(fig_pie, use_container_width=True)

        st.subheader(f"💰 Net Income and Cash Flow by Region - {scenario_selected} Scenario")
        region_grouped = filtered_df.groupby('Region').agg({'Net Income': 'sum', 'Cash Flow': 'sum'}).reset_index()
        fig_region = go.Figure()
        fig_region.add_trace(go.Bar(x=region_grouped['Region'], y=region_grouped['Net Income'], name='Net Income', marker_color='green'))
        fig_region.add_trace(go.Bar(x=region_grouped['Region'], y=region_grouped['Cash Flow'], name='Cash Flow', marker_color='blue'))
        fig_region.update_layout(barmode='group', title="Net Income and Cash Flow by Region", template="plotly_white")
        st.plotly_chart(fig_region, use_container_width=True)

        st.subheader("📉 Revenue to Net Income Walk (Waterfall Chart)")
        waterfall_fig = go.Figure(go.Waterfall(
            name="20", orientation="v",
            measure=["absolute", "relative", "relative", "relative", "relative", "total"],
            x=["Revenue", "-COGS", "-Opex", "-Depreciation", "-Tax", "Net Income"],
            y=[agg['Revenue (Group)'], -agg['COGS'], -agg['Operating Expenses'], -agg['Depreciation'], -agg['Tax'], net_income],
            connector={"line": {"color": "rgb(63, 63, 63)"}}
        ))
        waterfall_fig.update_layout(title="Waterfall Chart: Revenue to Net Income", template="plotly_white")
        st.plotly_chart(waterfall_fig, use_container_width=True)

        st.subheader("📋 Forecast Details by Product")
        display_df = filtered_df.set_index('Product').copy()
        for col in display_df.columns:
            if '%' in col:
                display_df[col] = display_df[col].map('{:.1f}%'.format)
            elif col not in ['Scenario', 'Region', 'Year']:
                display_df[col] = display_df[col].map('{:,.2f}'.format)
        st.dataframe(display_df)

        st.subheader("📎 Margin Summary")
        margin_summary = filtered_df.groupby('Product')[['Gross Margin %', 'Operating Margin %', 'Net Margin %']].mean().reset_index()
        for col in ['Gross Margin %', 'Operating Margin %', 'Net Margin %']:
            margin_summary[col] = margin_summary[col].map('{:.1f}%'.format)
        st.dataframe(margin_summary, use_container_width=True)

        def convert_df_to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        excel_download = convert_df_to_excel(results_df)
        st.download_button(label="📥 Download Results as Excel", data=excel_download, file_name='forecast_results_by_product.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
else:
    st.info("Please upload an Excel file with columns: Scenario, Product, Region, Year, Units Sold, Price per Unit, FX Rate, COGS %, Operating Expenses. Optional: Depreciation, Tax Rate")
