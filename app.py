import streamlit as st
import pandas as pd
import os, sys
current_path = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_path)
from stat_functions import *

uploaded_file = st.file_uploader("Upload Files", type=["csv", "xlsx"])
msg = st.empty()

if not uploaded_file:
    msg.info("Please upload a file containing raw sample data")
    st.stop()

if uploaded_file:
    msg.info("Reading File...")
    #check file type and read into pandas
    try: 
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        msg.warning("Error reading file. Please ensure it is a valid csv or excel file")
        st.error(e)
        st.stop()
    msg.write("")
    # st.write(df.head())

    #get column names and verify we have what we need.
    required_cols = ['LOC_ID', 'LOC_TYPE', 'LOC_SUBTYPE', 'CHEMICAL_NAME', 'REPORT_RESULT_VALUE', 'REPORT_RESULT_UNIT', 'SAMPLE_DATE']
    col_names = df.columns
    if not set(required_cols).issubset(col_names):
        msg.warning("Missing required columns:", set(required_cols) - set(col_names))
        st.stop()

    df['SAMPLE_DATE'] = pd.to_datetime(df['SAMPLE_DATE'])#, format='%m/%d/%Y')
    df['SAMPLE_YEAR'] = df['SAMPLE_DATE'].dt.year

    #get unique metals
    metals = df['CHEMICAL_NAME'].unique()
    metals.sort()

    msg.info('Currently testing for a single chemical. Select one from the dropdown below.')
    chemical = st.selectbox("Chemicals:", [' '] + list(metals))
    if not chemical or chemical == ' ':
        st.stop()
    if chemical:
        #get alumnium data
        metal_data = df[df['CHEMICAL_NAME'] == chemical]
        metal_name = metal_data['CHEMICAL_NAME'].unique()[0]
        try: 
            stats = get_metal_stats(metal_data, metal_name)
            to_sheet(metal_data, stats)

            #lets also show a scatter plot
            scatter(metal_data, metal_name, True)
            st.write('Sample count:', len(metal_data))
            st.image(f'files/{metal_name}.png')
        except Exception as e:
            msg.warning("Error processing data. Please ensure the data is valid")
            st.write(metal_data)
            st.error(e)
            st.stop()
    msg.info("Ready for download. Click the download button at the bottom.")

with open ("files/output.xlsx", "rb") as f:
    st.download_button("Download", f, "stats.xlsx")