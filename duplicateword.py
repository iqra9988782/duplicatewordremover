import streamlit as st
import pandas as pd
import io
import base64

def get_download_link(df, filename="unique_keywords.xlsx"):
    """Generate a download link for the processed file"""
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)
    b64 = base64.b64encode(buffer.read()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download Processed File</a>'

def process_file(uploaded_file):
    """Process the Excel file and remove duplicates"""
    try:
        df = pd.read_excel(uploaded_file)
        
        # Create empty series for collecting keywords
        all_keywords = pd.Series(dtype=str)
        
        # Collect stats for each column
        column_stats = {}
        
        # Process each column
        for col in df.columns:
            keywords_in_col = df[col].dropna().astype(str)
            column_stats[col] = len(keywords_in_col)
            all_keywords = pd.concat([all_keywords, keywords_in_col], ignore_index=True)
        
        # Remove duplicates
        unique_keywords = all_keywords.drop_duplicates().reset_index(drop=True)
        
        # Create result DataFrame
        result_df = pd.DataFrame(unique_keywords, columns=['Unique Keywords'])
        
        return result_df, {
            'total': len(all_keywords),
            'unique': len(unique_keywords),
            'duplicates': len(all_keywords) - len(unique_keywords),
            'column_stats': column_stats
        }
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        return None, None

# Set page config
st.set_page_config(
    page_title="Excel Duplicate Remover",
    page_icon="📊",
    layout="centered"
)

# Add custom CSS
st.markdown("""
    <style>
        .main {
            padding: 2rem;
        }
        .stButton>button {
            width: 100%;
        }
        .success-text {
            color: #28a745;
        }
        .info-box {
            padding: 1rem;
            border-radius: 0.5rem;
            background-color: #f8f9fa;
            margin: 1rem 0;
        }
    </style>
""", unsafe_allow_html=True)

# App header
st.title("📊 Excel Duplicate Remover")
st.markdown("""
    Upload your Excel file to remove duplicate values across all columns.
    The tool will:
    - Process all columns in your Excel file
    - Extract all unique values
    - Remove duplicates
    - Provide detailed statistics
    - Generate a downloadable Excel file with unique values
""")

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    with st.spinner('Processing your file...'):
        result_df, stats = process_file(uploaded_file)
        
        if result_df is not None:
            # Show statistics
            st.subheader("📈 Processing Statistics")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Total Keywords", stats['total'])
            with col2:
                st.metric("Unique Keywords", stats['unique'])
            with col3:
                st.metric("Duplicates Removed", stats['duplicates'])
            
            # Show column-wise statistics
            st.subheader("📊 Column-wise Statistics")
            for col, count in stats['column_stats'].items():
                st.text(f"{col}: {count} keywords")
            
            # Generate download link
            st.markdown("### ⬇️ Download Results")
            st.markdown(get_download_link(result_df), unsafe_allow_html=True)
            
            # Preview of results
            st.subheader("👀 Preview of Unique Keywords")
            st.dataframe(result_df.head(10))
            
            # Show success message
            st.success("✅ Processing completed successfully!")
else:
    st.info("👆 Upload an Excel file to get started")

# Footer
st.markdown("""
---
Made with ❤️ by [Muhammad Ismail]  
For support, contact: muhammadismailkpt@gmail.com
""")
