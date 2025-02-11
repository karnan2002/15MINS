import streamlit as st
import pandas as pd
from sqlalchemy import create_engine
import openpyxl

def fetch_data(server_details_df):
    combined_data = []
    
    for index, row in server_details_df.iterrows():
        server_name = row['ServerName']
        connection_string = row['ConnectionString']
        
        try:
            engine = create_engine(connection_string)
            
            # Execute the first query
            df1 = pd.read_sql("SELECT STORE FROM CONTROL WHERE REG_NUM='001'", engine)
            
            # Execute the second query
            df2 = pd.read_sql("EXEC xp_dirtree 'E:\\DB_AUTOBACKUP', 1, 1", engine)
            
            # Check if any of the DataFrames are empty
            if not df1.empty or not df2.empty:
                df_combined = pd.concat([df for df in [df1, df2] if not df.empty], ignore_index=True)
                combined_data.append(df_combined)
        
        except Exception as e:
            st.error(f"Error for server {server_name}: {str(e)}")
            continue
    
    return combined_data

def main():
    st.title("Database Query Executor")
    st.subheader("Upload Site IDs File")
    
    uploaded_file = st.file_uploader("Upload Site IDs file", type=["xlsx", "txt", "csv"])
    
    if uploaded_file:
        if uploaded_file.name.endswith("xlsx"):
            server_details_df = pd.read_excel(uploaded_file)
        elif uploaded_file.name.endswith("csv"):
            server_details_df = pd.read_csv(uploaded_file)
        else:
            server_details_df = pd.read_table(uploaded_file)
        
        st.write("File loaded successfully!")
        
        if st.button("Fetch Data"):
            combined_data = fetch_data(server_details_df)
            
            if combined_data:
                final_df = pd.concat(combined_data, ignore_index=True)
                
                # Save to Excel
                output_filename = "11-02-2025_QC.xlsx"
                with pd.ExcelWriter(output_filename, engine='openpyxl') as excel_writer:
                    final_df.to_excel(excel_writer, sheet_name='Combined', index=False)
                
                st.success("Data successfully fetched and saved!")
                st.download_button(label="Download Excel File", data=open(output_filename, 'rb'), file_name=output_filename, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            else:
                st.warning("No data found to export.")

if __name__ == "__main__":
    main()
