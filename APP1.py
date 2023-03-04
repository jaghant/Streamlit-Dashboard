from pathlib import Path  # Python Standard Library
import plotly.express as px # pip install plotly-express
import pandas as pd  # pip install pandas
import xlwings as xw  # pip install xlwings
import streamlit as st # pip install streamlit
from streamlit_option_menu import option_menu
import pickle
import time
import numpy as np
import datetime
from PIL import Image
from st_aggrid import AgGrid



st.set_page_config(page_title="My dashboard",
                   page_icon=":dart:",
                   layout="wide"
)

    
# -----READ EXCEL ------
@st.cache             
def get_data_from_excel():
            df         = pd.read_excel(
            io         = 'supermarkt_sales.xlsx',
            engine     = 'openpyxl',
            sheet_name = 'Sales',
            skiprows   = 3,
            usecols    = 'B:G',
            nrows      = 1000,
            )

# ------------------------------- Add 'hour' column to dataframe----------------------------------------------------
            df["hour"] = pd.to_datetime(df["Time"], format = "%H:%M:%S").dt.hour
            return df
    
df = get_data_from_excel()

   

# ----------------------------------------------- SIDEBAR -----------------------------------------------
st.title("🎥 People Counting Dashboard")
with st.sidebar:
     image = Image.open('logo.png')
     st.image(image,  width = 200)
# st.sidebar.title("🎥 PCD")

# col1, col2 = st.columns(2)
# df["Date"] = pd.to_datetime(df["Date"]).dt.date
# with col1:
#         start_date = st.date_input("Start date", value=df["Date"].min())
# with col2:
#         end_date = st.date_input("End date", value=df["Date"].max())
# data = df.loc[df["Date"].between(start_date, end_date)]

df_selection = df
# st.dataframe(df_selection)

# --------------------------------------------- MAINPAGE ----------------------------------------------------


# st.markdown("------------------------------------------------------------------------------")

#----------------------------------------------- TOP KPI'set-------------------------------------------------

total        = int(df_selection["Total"].sum())
people_in  = int(df_selection["People_In"].sum())
people_out = int(df_selection["People_Out"].sum())
people_in_per =  people_in / total * 100
format_people_in_per = ("{:.1f}".format(people_in_per))
people_out_per = people_out / total * 100
format_people_out_per = ("{:.1f}".format(people_out_per))



with open('style1.css') as f:
    st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)
    
left_column , middle_column, right_column = st.columns(3)
with left_column:
            st.subheader("👩‍👩‍👦‍👦Total")
            st.subheader(f"{total}")            
with middle_column:
            st.subheader(":two_men_holding_hands:People In")
            st.subheader(f"{people_in} [{format_people_in_per}%]")
with right_column:
            st.subheader(":two_men_holding_hands:People Out")
            st.subheader(f"{people_out} [{format_people_out_per}%]")

#----------------------------------- Date By People IN [BAR CHART] -------------------------------------------

st.markdown("---------------------------------------------------------------------------")  
people_in_groupby = df_selection.groupby(df_selection['Date'].dt.strftime('%B'))['People_In'].sum().sort_index()
  
fig_people_in = px.line(
            people_in_groupby,
            x = people_in_groupby.index,
            y = "People_In",
            orientation = "v",
            title = "<b>Month by People In</b>",
            color_discrete_sequence = ["#12A999"] * len(people_in_groupby),
            template = "plotly_white",
            markers="**",
        )
fig_people_in.update_layout(
            plot_bgcolor = "rgba(0,0,0,0)",
            xaxis        = (dict(showgrid=False)),
            yaxis        = (dict(showgrid=False))
        )

        
#----------------------------------- Date By People Out [BAR CHART] -------------------------------------------


   
people_out_groupby = df_selection.groupby(df_selection['Date'].dt.strftime('%B'))['People_Out'].sum().sort_index()

fig_people_out = px.bar(
            people_out_groupby,
            x = people_out_groupby.index,
            y = "People_Out",
            orientation = "v",
            title = "<b>Month by People Out</b>",
            color_discrete_sequence = ["#12A999"] * len(people_out_groupby),
            template = "plotly_white",
        )
fig_people_out.update_layout(
            plot_bgcolor = "rgba(0,0,0,0)",
            xaxis        = (dict(showgrid=False)),
            yaxis        = (dict(showgrid=False)),
        
        )

         
col1 , col2 = st.columns(2)
with col1:
        st.write(fig_people_in)
with col2:
        st.write(fig_people_out)
st.markdown("---------------------------------------------------------------------------")   
#----------------------------------- Date By Total [BAR CHART] -------------------------------------------


total_groupby = df_selection.groupby(df_selection['Date'].dt.strftime('%B'))['Total'].sum().sort_index()

fig_total = px.bar(
            total_groupby,
            x = total_groupby.index,
            y = "Total",
            orientation = "v",
            title = "<b>Month by Total</b>",
            color_discrete_sequence = ["#12A999"] * len(total_groupby),
            template = "plotly_white",
        )
fig_total.update_layout(
            plot_bgcolor = "rgba(0,0,0,0)",
            xaxis        = (dict(showgrid=False)),
            yaxis        = (dict(showgrid=False)),
        
        )
st.write(fig_total)

st.markdown("--------------------------------------------------------------------------------") 

        # --------HIDE STREAMLIT STYLE-------

hide_st_style = """
                    <style>
                    #MainMenu {visibility: hidden;}
                    footer {visibility: hidden;}
                    header {visiblity: hidden:}
                    </style>
                    """
st.markdown(hide_st_style, unsafe_allow_html=True)

# year_option = df["Date"].tolist()
# year = st.selectbox("choose it?", year_option, 0)
# data = df[df["Date"]==year]

# fig = px.scatter(data, y="Total", x="Product_line",color_discrete_sequence=['#0BFCE2'])
# fig.update_layout(width=1000, height=800, plot_bgcolor="rgba(0,0,0,0)", xaxis=dict(showgrid=False), yaxis=dict(showgrid=False))
# st.write(fig)
st.title("📋Report")
# start_date = st.date_input("Fron Date", value = df["Date"].min())
# end_date   = st.date_input("End Date", value = df["Date"].max())
# filtered_data = df.loc[(df["Date"] >= start_date) & (df["Date"] <= end_date)]
# period = st.selectbox("Select Period:", get_all_periods())
AgGrid(df)

