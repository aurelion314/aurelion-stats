import streamlit as st
import pandas as pd
import os, sys
current_path = os.path.dirname(os.path.abspath(__file__))
sys.path.append(current_path)
from stat_functions import *

uploaded_file = st.file_uploader("Upload Files", type=["csv", "xlsx"])
msg = st.empty()
button = st.empty()

#we need a files folder. Make if it doesn't exist
try: 
    if not os.path.exists('files'):
        os.makedirs('files')
except Exception as e:
    st.warning('Error creating files folder.')
    st.error(e)
    st.stop()

if not uploaded_file:
    msg.info("Please upload a file containing raw sample data")
    st.stop()

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
    msg.warning("Missing required columns: " + ", ".join(set(required_cols) - set(col_names)))
    st.stop()

df['SAMPLE_DATE'] = pd.to_datetime(df['SAMPLE_DATE'])#, format='%m/%d/%Y')
df['SAMPLE_YEAR'] = df['SAMPLE_DATE'].dt.year

#get unique metals
metals = df['CHEMICAL_NAME'].unique()
metals.sort()

msg.info('Select a chemical from the dropdown below.')
chemical = st.selectbox("Chemicals:", [' '] + ['All'] + list(metals))
if not chemical or chemical == ' ':
    st.stop()
msg.info('Calculating statistics..')
try: 
    if chemical == 'All':
        filename =  os.path.join(current_path, "files/output.xlsx")
        workbook = xlsxwriter.Workbook(filename)
        for i, metal in enumerate(list(metals)):
            msg.info(f"Calculating Statistics for {metal} ({i+1}/{len(metals)})")
            metal_data = df[df['CHEMICAL_NAME'] == metal]
            metal_name = metal_data['CHEMICAL_NAME'].unique()[0]
            stats = get_metal_stats(metal_data, metal_name)
            to_sheet(metal_data, stats, workbook)

        workbook.close()

    else:
        #get single data
        metal_data = df[df['CHEMICAL_NAME'] == chemical]
        metal_name = metal_data['CHEMICAL_NAME'].unique()[0]
        stats = get_metal_stats(metal_data, metal_name)
        to_sheet(metal_data, stats)

        st.write('Sample count:', len(metal_data))

        #lets also show a plot
        plot_type = st.selectbox("Plot Type:", ['Scatter', 'Line', 'Scatter by Site'])
        if plot_type == 'Scatter':
            scatter(metal_data, metal_name, True)
            st.image(f'files/{metal_name}.png')
        elif plot_type == 'Line':
            siteLineChart(metal_data, metal_name)
            st.image(f'files/{metal_name}_site_line.png')
        elif plot_type == 'Scatter by Site':
            siteScatter(metal_data, metal_name)
            st.image(f'files/{metal_name}_site_scatter.png')
        #filter colums to only the required cols
        metal_data = metal_data[required_cols]
        st.write(metal_data)
    
except Exception as e:
    msg.warning("Error processing data. Please ensure the data is valid")
    st.write(metal_data)
    st.error(e)
    st.stop()

msg.info("Ready for download.")

with open ("files/output.xlsx", "rb") as f:
    button.download_button("Download", f, "stats.xlsx")