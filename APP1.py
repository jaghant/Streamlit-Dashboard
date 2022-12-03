from pathlib import Path  # Python Standard Library
import plotly.express as px # pip install plotly-express
import pandas as pd  # pip install pandas
import xlwings as xw  # pip install xlwings
import streamlit as st # pip install streamlit
from streamlit_option_menu import option_menu
import pickle

# import pickle
# from pathlib import Path
# import streamlit_authenticator as stauth # pip install streamlit-authenticator

st.set_page_config(page_title="My dashboard",
                   page_icon=":dart:",
                   layout="wide"
)

    
# -----READ EXCEL ------
@st.cache             
def get_data_from_excel():
            df = pd.read_excel(
                io='supermarkt_sales.xlsx',
                engine='openpyxl',
                sheet_name='Sales',
                skiprows=3,
                usecols='B:R',
                nrows=1000,
            )

        # Add 'hour' column to dataframe
            df["hour"] = pd.to_datetime(df["Time"], format = "%H:%M:%S").dt.hour
            return df
df = get_data_from_excel()

        # this_dir = Path(__file__).parent if '__file__' in locals() else Path.cwd()
        # wb_file_path = this_dir / 'supermarkt_sales.xlsx'

        # wb = xw.Book(wb_file_path)
        # sht = wb.sheets['Sales']
        # rng = sht.range('B4:R1004')

        #df = rng.options(pd.DataFrame, index=False, header=True).value
        #print(df)

        #st.dataframe(df) 

        # ----SIDEBAR-----
# authenticator.logout("Logout", "sidebar")
# st.sidebar.title(f"Welcome {name}")
st.title(":dart: Dashboard")
st.sidebar.title(":dart: My Dashboard")

st.sidebar.header("Please Filter Here:")
city = st.sidebar.multiselect(
                "Select the City:",
                options = df["City"].unique(),
                default = df["City"].unique()
                )

customer_type = st.sidebar.multiselect(
                "Select the Customer Type:",
                options = df["Customer_type"].unique(),
                default = df["Customer_type"].unique()
                )

gender = st.sidebar.multiselect(
                "Select the Gender:",
                options = df["Gender"].unique(),
                default = df["Gender"].unique()
                )   

branch = st.sidebar.multiselect(
                "Select the Branch:",
                options = df["Branch"].unique(),
                default = df["Branch"].unique()
        )

# product = st.sidebar.multiselect(
#                 "Select the Product:",
#                 options = df["Product_line"].unique(),
#                 default = df["Product_line"].unique()
#         )

payment = st.sidebar.multiselect(
                "Select the Payment:",
                options = df["Payment"].unique(),
                default = df["Payment"].unique()
        )
agree = st.sidebar.checkbox('I agree')

if agree:
    st.sidebar.write(':sunglasses::sunglasses::sunglasses:')

df_selection = df.query(
            "City == @city & Customer_type == @customer_type & Gender == @gender & Branch == @branch & Payment == @payment"
        )


# st.dataframe(df_selection)

        #----- MAINPAGE -----


st.markdown("------------------------------------------------------------------------------")

        # TOP KPI'set

total_sales = int(df_selection["Total"].sum())
average_rating = round(df_selection["Rating"].mean(), 1)
star_rating = ":star:" * int(round(average_rating, 0))
average_sale_by_transaction = round(df_selection["Total"].mean(), 2)

left_column , middle_column, right_column = st.columns(3)
with left_column:
            st.subheader("Total Sales:")
            st.subheader(f"US $ {total_sales:,}")
with middle_column:
            st.subheader("Average Rating:")
            st.subheader(f"{average_rating} {star_rating}")
with right_column:
            st.subheader("Average Sales Per Transaction:")
            st.subheader(f"US $ {average_sale_by_transaction}")
            
              
            
#------------------ SALES BY PRODUCT LINE [BAR CHART] ---------------
st.markdown("---------------------------------------------------------------------------")  
sales_by_product_line = (
            df_selection.groupby(by=["Product_line"]).sum()[["Total"]].sort_values(by="Total")
            
        )        
fig_product_sales = px.bar(
            sales_by_product_line,
            x="Total",
            y=sales_by_product_line.index,
            orientation="h",
            title="<b>Sales by Product Line</b>",
            color_discrete_sequence = ["#0083b8"] * len(sales_by_product_line),
            template="plotly_white",
        )
fig_product_sales.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis=(dict(showgrid=False))
        )
        # st.plotly_chart(fig_product_sales)

# ------------------- SALES BY HOUR [BAR CHART] ----------------

sales_by_hour = df_selection.groupby(by=["hour"]).sum()[["Total"]]
fig_hourly_sales = px.bar(
            sales_by_hour,
            x=sales_by_hour.index,
            y="Total",
            title="<b>Sales by hour</b>",
            color_discrete_sequence=['#0083B8'] * len(sales_by_hour),
            template="plotly_white",
        )
fig_hourly_sales.update_layout(
            xaxis=dict(tickmode="linear"),
            plot_bgcolor="rgba(0,0,0,0)",
            yaxis=(dict(showgrid=False))
        )

        # st.plotly_chart(fig_hourly_sales)
        


left_column, right_column = st.columns(2)
left_column.plotly_chart(fig_hourly_sales, use_container_width=True)
right_column.plotly_chart(fig_product_sales, use_container_width=True)
st.markdown("--------------------------------------------------------------------------------") 
# --------- sales by branch [pie chart] ----------------

sales_by_gender = pd.read_excel("supermarkt_sales.xlsx",
                                usecols="C:K",
                                )
sales_by_gender.dropna(inplace=True)

left_column, right_column = st.columns(2)
with left_column:
        pie_chart = px.pie(sales_by_gender,
                        title="Sales by Gender",
                        values="Total",
                        names="Gender")
        st.plotly_chart(pie_chart) 

# --------------sales by customer type [pie chart]
sales_by_customer_type = pd.read_excel("supermarkt_sales.xlsx",
                                usecols="D:J",
                                )
sales_by_customer_type.dropna(inplace=True)

with right_column:
        pie_chart = px.pie(sales_by_customer_type,
                        title="Sales by Customer Type",
                        values="Total",
                        names="Customer_type",
                        )
        st.plotly_chart(pie_chart)                               

        # --------HIDE STREAMLIT STYLE-------

# hide_st_style = """
#                     <style>
#                     #MainMenu {visibility: hidden;}
#                     footer {visibility: hidden;}
#                     header {visiblity: hidden:}
#                     </style>
#                     """
# st.markdown(hide_st_style, unsafe_allow_html=True)