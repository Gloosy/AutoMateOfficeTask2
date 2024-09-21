import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Sales Reports", page_icon="barr", layout="wide")

# Custom CSS for new font and styling
custom_css = """
    <style>
    html, body, [class*="css"] {
        font-family: 'Helvetica Neue', sans-serif;  
        background-color: #F4F6F9; 
        color: #2D2D2D; 
    }
    .sidebar .sidebar-content {
        background-color: #0056b3; 
        color: white;
        border-right: 2px solid #004494;  
    }
    .css-1d391kg {
        background-color: #FFFFFF; 
        border-radius: 12px; 
        padding: 15px; 
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
    }
    .stButton>button {
        background-color: #007BFF; 
        color: white;
        border-radius: 8px;
        border: none; 
        padding: 12px 20px; 
        transition: background-color 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #0056b3; 
    }
    h1, h2, h3, h4, h5, h6 {
        color: #343A40; 
        margin: 12px 0; 
        font-weight: bold;
    }
    .stMarkdown {
        font-size: 17px; 
        line-height: 1.5; 
    }
    .stTextInput>div>input {
        border-radius: 5px; 
        border: 1px solid #007BFF; 
        padding: 10px; 
    }
    </style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

@st.cache_data
def get_data_from_excel():
    try:
        df = pd.read_excel(
            io="supermarkt_sales.xlsx",
            engine="openpyxl",
            sheet_name="Sales",
            skiprows=3,
            usecols="B:R",
            nrows=1000,
        )
        df["hour"] = pd.to_datetime(df["Time"], format="%H:%M:%S").dt.hour
        df["Date"] = pd.to_datetime(df["Date"])  # Ensure Date is in datetime format
        return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame()  # Return an empty DataFrame

df = get_data_from_excel()

# SIDEBAR
st.sidebar.header("Please Filter Here:")
city = st.sidebar.multiselect("Select the City:", options=df["City"].unique(), default=df["City"].unique())
customer_type = st.sidebar.multiselect("Select the Customer Type:", options=df["Customer_type"].unique(), default=df["Customer_type"].unique())
gender = st.sidebar.multiselect("Select the Gender:", options=df["Gender"].unique(), default=df["Gender"].unique())

# Adding a date range filter
start_date = st.sidebar.date_input("Start Date", value=df["Date"].min())
end_date = st.sidebar.date_input("End Date", value=df["Date"].max())
df_selection = df.query("City == @city & Customer_type == @customer_type & Gender == @gender & Date >= @start_date & Date <= @end_date")

if df_selection.empty:
    st.warning("No data available based on the current filter settings!")
    st.stop()

# Title with new style
st.markdown("<h1 style='text-align: center; color: #000; font-size: 50px; font-weight: bold;'>Sales Reports Monthly</h1>", unsafe_allow_html=True)
st.markdown("##")

# KPI's
total_sales = int(df_selection["Total"].sum())
average_rating = round(df_selection["Rating"].mean(), 1)
star_rating = ":star:" * int(round(average_rating, 0))
average_sale_by_transaction = round(df_selection["Total"].mean(), 2)
total_transactions = df_selection.shape[0]
total_customers = df_selection["CustomerID"].nunique() if 'CustomerID' in df.columns else "N/A"

# Customize KPI display
kpi_columns = st.columns(5)
kpi_data = [
    ("Total Sales:", f"US $ {total_sales:,}"),
    ("Average Rating:", f"{average_rating} {star_rating}"),
    ("Average Sales Per Transaction:", f"US $ {average_sale_by_transaction}"),
    ("Total Transactions:", total_transactions),
    ("Total Customers:", total_customers),
]

for col, (title, value) in zip(kpi_columns, kpi_data):
    with col:
        st.subheader(title)
        st.subheader(value)

st.markdown("---")

# SALES BY PRODUCT LINE [BAR CHART]
sales_by_product_line = df_selection.groupby(by=["Product line"])[["Total"]].sum().sort_values(by="Total")
fig_product_sales = px.bar(
    sales_by_product_line,
    x="Total",
    y=sales_by_product_line.index,
    orientation="h",
    title="<b>Sales by Product Line</b>",
    color_discrete_sequence=["#007BFF"] * len(sales_by_product_line),
    template="plotly_white",
)
fig_product_sales.update_layout(
    title_font=dict(family="Helvetica Neue", size=20, color="#007BFF"),
    font=dict(family="Helvetica Neue", size=14),
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=dict(showgrid=False),
)

# SALES BY HOUR [BAR CHART]
sales_by_hour = df_selection.groupby(by=["hour"])[["Total"]].sum()
fig_hourly_sales = px.bar(
    sales_by_hour,
    x=sales_by_hour.index,
    y="Total",
    title="<b>Sales by Hour</b>",
    color_discrete_sequence=["#FFC107"] * len(sales_by_hour),
    template="plotly_white",
)
fig_hourly_sales.update_layout(
    title_font=dict(family="Helvetica Neue", size=20, color="#FFC107"),
    font=dict(family="Helvetica Neue", size=14),
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=dict(tickmode="linear"),
    yaxis=dict(showgrid=False),
)

# SALES BY CUSTOMER TYPE [PIE CHART]
sales_by_customer_type = df_selection.groupby(by=["Customer_type"])[["Total"]].sum().reset_index()
fig_customer_type = px.pie(
    sales_by_customer_type,
    names='Customer_type',
    values='Total',
    title="<b>Sales Distribution by Customer Type</b>",
    color_discrete_sequence=px.colors.qualitative.Pastel,
)
fig_customer_type.update_traces(textinfo='percent+label')

# SALES TREND OVER TIME [LINE CHART]
sales_trend = df_selection.groupby(by=["Date"])[["Total"]].sum().reset_index()
fig_sales_trend = px.line(
    sales_trend,
    x="Date",
    y="Total",
    title="<b>Sales Trend Over Time</b>",
    markers=True,
    template="plotly_white"
)
fig_sales_trend.update_layout(
    title_font=dict(family="Helvetica Neue", size=20, color="#28A745"),
    font=dict(family="Helvetica Neue", size=14),
    plot_bgcolor="rgba(0,0,0,0)",
)

# CUSTOMER DEMOGRAPHICS [BAR CHART]
customer_demographics = df_selection.groupby(by=["Gender"])[["CustomerID"]].nunique().reset_index() if 'CustomerID' in df.columns else df_selection.groupby(by=["Gender"])[["Gender"]].size().reset_index(name='Total Customers')
customer_demographics.columns = ['Gender', 'Total Customers']

fig_demographics = px.bar(
    customer_demographics,
    x='Gender',
    y='Total Customers',
    title="<b>Customer Demographics</b>",
    color='Total Customers',
    color_continuous_scale=px.colors.sequential.Viridis,
)
fig_demographics.update_layout(
    title_font=dict(family="Helvetica Neue", size=20, color="#6F42C1"),
    font=dict(family="Helvetica Neue", size=14),
    plot_bgcolor="rgba(0,0,0,0)",
)

# SHOWING THE GRAPHS
left_column, right_column = st.columns(2)
with left_column:
    st.plotly_chart(fig_product_sales, use_container_width=True)
    st.plotly_chart(fig_hourly_sales, use_container_width=True)
with right_column:
    st.plotly_chart(fig_customer_type, use_container_width=True)
    st.plotly_chart(fig_sales_trend, use_container_width=True)
st.plotly_chart(fig_demographics, use_container_width=True)
