import streamlit as st
import pandas as pd
import io
import base64

# Ensure required packages are installed
try:
    import openpyxl
except ImportError:
    st.error("Installing required packages...")
    import subprocess
    subprocess.check_call(["pip", "install", "openpyxl"])
    import openpyxl

st.set_page_config(
    page_title="Excel Duplicate Remover",
    page_icon="📊",
    layout="centered"
)

st.title("Excel Duplicate Remover")
st.write("Upload your Excel file to remove duplicates from all columns")

def process_excel(uploaded_file):
    try:
        # Read Excel file with explicit engine
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        
        # Create empty series for collecting all values
        all_values = pd.Series(dtype=str)
        
        # Process each column
        for col in df.columns:
            # Add non-null values from this column
            column_values = df[col].dropna().astype(str)
            all_values = pd.concat([all_values, column_values])
        
        # Remove duplicates
        unique_values = all_values.drop_duplicates().reset_index(drop=True)
        
        # Create output DataFrame
        output_df = pd.DataFrame(unique_values, columns=['Unique Values'])
        
        return output_df, {
            'total': len(all_values),
            'unique': len(unique_values),
            'removed': len(all_values) - len(unique_values)
        }
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        return None, None

def get_download_link(df):
    # Create download link
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    excel_data = output.getvalue()
    b64 = base64.b64encode(excel_data).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="unique_values.xlsx">📥 Download Excel File</a>'

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    # Show processing message
    with st.spinner('Processing your file...'):
        result_df, stats = process_excel(uploaded_file)
        
        if result_df is not None and stats is not None:
            # Show statistics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Values", stats['total'])
            with col2:
                st.metric("Unique Values", stats['unique'])
            with col3:
                st.metric("Duplicates Removed", stats['removed'])
            
            # Show preview
            st.subheader("Preview of Results")
            st.dataframe(result_df.head())
            
            # Download button
            st.markdown(get_download_link(result_df), unsafe_allow_html=True)
            
            # Success message
            st.success("✅ File processed successfully!")
else:
    st.info("👆 Upload an Excel file to start")

# Footer
st.markdown("---")
st.markdown("Made with ❤️ by Muhammad Ismail")
