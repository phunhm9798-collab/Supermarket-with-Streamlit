# Import necessary modules
import pandas as pd
import plotly.express as px
import streamlit as st


def get_data_from_excel():
    df = pd.read_excel(
        io='Supermarket.xlsx',
        engine='openpyxl',
        sheet_name='Sales',
        skiprows=3,
        usecols='B:R',
        nrows=1000,
    )

# Add 'hour' column to dataframe
    df["hour"] = pd.to_datetime(df["Time"], format="%H:%M:%S").dt.hour
    return df


df = get_data_from_excel()

st.set_page_config(page_title='Sales Dashboard',
                   page_icon=':bar_chart:',
                   layout='wide')

st.dataframe(df)

# Sidebar
st.sidebar.header('Please Filter Here:')
city = st.sidebar.multiselect(
    'Select the City',
    options=df['City'].unique(),
    default=df['City'].unique()
)

customer_type = st.sidebar.multiselect(
    "Select the customer type: ",
    options=df['Customer_type'].unique(),
    default=df['Customer_type'].unique()
)

gender = st.sidebar.multiselect(
    "Select the Gender : ",
    options=df['Gender'].unique(),
    default=df['Gender'].unique()
)

df_selection = df.query(
    'City == @city & Customer_type == @customer_type & Gender == @gender'
)

# Mainpage
st.title('bar_chart : Sales Dashboard')
st.markdown('##')

# Top KPI's
total_sales = int(df_selection['Total'].sum())
average_rating = round(df_selection['Rating'].mean(), 1)
star_rating = ':star:' * int(round(average_rating, 0))
average_sales_by_transaction = round(df_selection['Total'].mean(), 2)

left_column, middle_column, right_column = st.columns(3)
with left_column:
    st.subheader('Total Sales: ')
    st.subheader(f'US $ {total_sales:,}')
with middle_column:
    st.subheader('Average ating:')
    st.subheader(f'{average_rating} {star_rating}')
with right_column:
    st.subheader("Average Sales Per Transaction:")
    st.subheader(f'US $ {average_sales_by_transaction}')

st.markdown('---')

# Sales by product line bar chart
sale_by_product_line = df_selection.groupby(by=["Product line"])[
    ["Total"]].sum().sort_values(by="Total")

fig_product_sales = px.bar(
    sale_by_product_line,
    x='Total',
    y=sale_by_product_line.index,
    orientation='h',
    title='<b> Sales by product Line</b>',
    color_discrete_sequence=['#0083B8'] * len(sale_by_product_line),
    template='plotly_white'
)

st.plotly_chart(fig_product_sales)

# Sale by hour bar chart
sale_by_hour = df_selection.groupby(by=["hour"])[["Total"]].sum()
fig_hourly_sales = px.bar(
    sale_by_hour,
    x=sale_by_hour.index,
    y='Total',
    title='<b>Sales by hour</b>',
    color_discrete_sequence=['#0083B8'] * len(sale_by_product_line),
    template='plotly_white'
)
st.plotly_chart(fig_hourly_sales)
