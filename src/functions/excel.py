import pandas as pd
import io
import streamlit as st

def interactive_excel(file):
    df = pd.read_excel(file)
    st.write(df)
    st.dataframe(df)

def create_project_structure_template():
    data = {
        "Section": ["Introduction", "Objectives", "Timeline", "Budget", "Team Members"],
        "Description": ["", "", "", "", ""],
        "Status": ["Not Started"] * 5,
        "Assigned To": [""] * 5
    }
    df = pd.DataFrame(data)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Project Structure', index=False)
        writer.save()
    output.seek(0)
    
    return output
