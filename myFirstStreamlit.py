import pandas as pd
import streamlit as st

st.title("Netimpact Day -1 Report 😄")
rawFile = st.file_uploader("Upload Excel", type=[".xlsx"])
if rawFile is not None:
    df = pd.ExcelFile(rawFile,engine='openpyxl')



    # """   Loading the EMP List and Raw Data   """

    emp_data = df.parse('Emp LIst')
    data = df.parse('Raw')
    try:
        data.drop(columns=['Tagging','Date','Time','OLMS ID','Unique','LOB'], inplace=True)
    except:
        pass

    
    
   
    # """ Creating custom Columns """

    data['Tagging'] = data['Type_1'] + '>>' + data['Area'] + '>>' + data['Sub Area']
    data['Date'] = data['Created time'].astype(str).str.split(' ').str[0]
    data['Time'] = data['Created time'].astype(str).str.split(' ').str[1]


    # """   Performing the marge    """

    data = data.merge(emp_data[['OLMS ID','Name','Adjustment LOB']],left_on='Agent Name.', right_on='Name' , how= 'left')

    
    # """ Drop the Name column, (Duplicate)"""

    data.drop('Name',inplace=True,axis=1)

   
   
    # """ Creating the Unique column (Primary Key)"""

    data['Unique'] = data['Date'].astype(str) + data['OLMS ID'].astype(str)


    # """
    # TOP 10
    # """
    Top_10 = data.pivot_table(
        index='Type_1',
        columns='Date',
        aggfunc='size',
        fill_value=0).assign(total=lambda df: df.sum(axis=1)).sort_values(by='total', ascending=False)


    # """tAGGING"""
    tagging = data.pivot_table(
        index='Tagging',
        columns='Date',
        aggfunc='size',
        fill_value=0).assign(total=lambda df: df.sum(axis=1)).sort_values(by='total', ascending=False)


    # """ Main Productivity """
    mainProductivity = data.pivot_table(
        index='Unique',
        columns='Date',
        aggfunc='size',
        fill_value=0).assign(total=lambda df: df.sum(axis=1)).sort_values(by='total', ascending=False)

    # """ad prod"""
    adProd = data.pivot_table(
        index='Agent Name.',
        columns='Date',
        aggfunc='size',
        fill_value=0).assign(total=lambda df: df.sum(axis=1)).sort_values(by='total', ascending=False)
    

    st.header("Representation of Data 📊")
    st.subheader("Top 10")
    st.dataframe(Top_10,use_container_width=True)
    st.subheader("Tagging")
    st.dataframe(tagging, use_container_width=True)
    st.subheader("Main Productivity")
    st.dataframe(mainProductivity, use_container_width=True)
    st.subheader("Agent Productivity")
    st.dataframe(adProd, use_container_width=True)  
    st.write("Thank you for using the Netimpact Day -1 Report tool! 😎")  
