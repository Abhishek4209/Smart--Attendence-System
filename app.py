import streamlit as st
import pandas as pd
from datetime import datetime
import time
ts = time.time()
date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")

filename = f"Attendance_{date}.csv"  
df = pd.read_csv(filename)

# Use Streamlit to display the DataFrame in an interactive web app
st.dataframe(df.style.highlight_max(axis=0))