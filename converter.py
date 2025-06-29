#!/usr/bin/env python3
"""
Streamlit Multi-Format Data to Excel Converter
A local web app for converting various data formats (JSON, CSV, TSV, XML) to Excel with file upload functionality.

Run with: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import json
import io
import xml.etree.ElementTree as ET
from datetime import datetime
import base64
import yaml
import csv

# Page configuration
st.set_page_config(
    page_title="Multi-Format Data to Excel Converter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        text-align: center;
        color: #1f77b4;
        font-size: 3rem;
        margin-bottom: 2rem;
    }
    .upload-section {
        border: 2px dashed #1f77b4;
        border-radius: 10px;
        padding: 2rem;
        margin: 1rem 0;
        background-color: #f0f8ff;
    }
    .stats-container {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .format-info {
        background-color: #e7f3ff;
        border-left: 4px solid #1f77b4;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def detect_format_from_content(content):
    """Dynamically detect data format from content"""
    content = content.strip()
    
    # JSON detection
    if (content.startswith('[') and content.endswith(']')) or (content.startswith('{') and content.endswith('}')):
        try:
            json.loads(content)
            return "JSON"
        except:
            pass
    
    # XML detection
    if content.startswith('<') and content.endswith('>'):
        try:
            ET.fromstring(content)
            return "XML"
        except:
            pass
    
    # YAML detection (check for YAML patterns)
    if ('- ' in content and ':' in content) or (content.count(':') > content.count(',') and '\n' in content):
        try:
            yaml.safe_load(content)
            return "YAML"
        except:
            pass
    
    # CSV/TSV detection
    lines = content.split('\n')
    if len(lines) > 1:
        first_line = lines[0]
        comma_count = first_line.count(',')
        tab_count = first_line.count('\t')
        
        if tab_count > comma_count and tab_count > 0:
            return "TSV"
        elif comma_count > 0:
            return "CSV"
    
    return "Unknown"

def parse_json_file(file_content):
    """Parse JSON file content"""
    try:
        return json.loads(file_content), None
    except json.JSONDecodeError as e:
        return None, f"Invalid JSON format: {e}"

def parse_csv_file(file_content, delimiter=','):
    """Parse CSV file content"""
    try:
        df = pd.read_csv(io.StringIO(file_content), delimiter=delimiter)
        return df.to_dict('records'), None
    except Exception as e:
        return None, f"Error parsing CSV: {e}"

def parse_tsv_file(file_content):
    """Parse TSV file content"""
    return parse_csv_file(file_content, delimiter='\t')

def parse_xml_file(file_content):
    """Parse XML file content"""
    try:
        root = ET.fromstring(file_content)
        
        # Handle different XML structures
        data = []
        
        # Try to find repeating elements (common structure)
        children = list(root)
        if len(children) > 0:
            # If root has multiple children, treat each as a record
            for child in children:
                record = {}
                if len(list(child)) > 0:
                    # Child has sub-elements
                    for subchild in child:
                        record[subchild.tag] = subchild.text or ""
                else:
                    # Child is a simple element
                    record[child.tag] = child.text or ""
                data.append(record)
        else:
            # Single record case
            record = {}
            for child in root:
                record[child.tag] = child.text or ""
            data.append(record)
        
        return data, None
    except ET.ParseError as e:
        return None, f"Invalid XML format: {e}"
    except Exception as e:
        return None, f"Error parsing XML: {e}"

def parse_yaml_file(file_content):
    """Parse YAML file content"""
    try:
        data = yaml.safe_load(file_content)
        if isinstance(data, list):
            return data, None
        elif isinstance(data, dict):
            return [data], None
        else:
            return None, "YAML must contain a list of objects or a single object"
    except yaml.YAMLError as e:
        return None, f"Invalid YAML format: {e}"

def parse_text_file(file_content):
    """Parse plain text file - attempt to detect format"""
    file_content = file_content.strip()
    
    # Try JSON first
    if file_content.startswith('[') or file_content.startswith('{'):
        data, error = parse_json_file(file_content)
        if data is not None:
            return data, None
    
    # Try CSV/TSV
    lines = file_content.split('\n')
    if len(lines) > 1:
        # Check for common delimiters
        first_line = lines[0]
        if '\t' in first_line:
            return parse_tsv_file(file_content)
        elif ',' in first_line:
            return parse_csv_file(file_content)
    
    return None, "Unable to detect format. Please specify the file type."

def process_uploaded_file(uploaded_file):
    """Process uploaded file based on its type"""
    file_extension = uploaded_file.name.split('.')[-1].lower()
    
    try:
        # Read file content
        if file_extension in ['json', 'csv', 'tsv', 'txt', 'xml', 'yaml', 'yml']:
            file_content = uploaded_file.read().decode('utf-8')
        else:
            return None, f"Unsupported file type: {file_extension}"
        
        # Parse based on extension
        if file_extension == 'json':
            return parse_json_file(file_content)
        elif file_extension == 'csv':
            return parse_csv_file(file_content)
        elif file_extension == 'tsv':
            return parse_tsv_file(file_content)
        elif file_extension == 'xml':
            return parse_xml_file(file_content)
        elif file_extension in ['yaml', 'yml']:
            return parse_yaml_file(file_content)
        elif file_extension == 'txt':
            return parse_text_file(file_content)
        else:
            return None, f"Unsupported file extension: {file_extension}"
            
    except UnicodeDecodeError:
        return None, "Unable to decode file. Please ensure it's a text file with UTF-8 encoding."
    except Exception as e:
        return None, f"Error processing file: {e}"

def validate_data(data):
    """Validate data for Excel conversion"""
    if not isinstance(data, list):
        return False, "Data must be an array/list of objects"
    
    if len(data) == 0:
        return False, "Data array is empty"
    
    if not isinstance(data[0], dict):
        return False, "Data must contain objects/dictionaries"
    
    return True, "Data is valid"

def create_excel_file(df, filename="converted_data.xlsx"):
    """Create Excel file with formatting"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Data', index=False)
        
        # Get workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Data']
        
        # Header format
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#4472C4',
            'font_color': 'white',
            'border': 1
        })
        
        # Number format
        number_format = workbook.add_format({
            'num_format': '#,##0.00',
            'border': 1
        })
        
        # Write headers with formatting
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Auto-adjust column widths
        for col_num, col_name in enumerate(df.columns):
            max_length = max(
                len(str(col_name)),
                df[col_name].astype(str).str.len().max() if len(df) > 0 else 0
            )
            worksheet.set_column(col_num, col_num, min(max_length + 2, 50))
        
        # Apply number formatting to numeric columns
        for col_num, col_name in enumerate(df.columns):
            if df[col_name].dtype in ['int64', 'float64']:
                worksheet.set_column(col_num, col_num, None, number_format)
    
    output.seek(0)
    return output

def main():
    """Main Streamlit app"""
    
    # Header
    st.markdown('<h1 class="main-header">📊 Multi-Format Data to Excel Converter</h1>', unsafe_allow_html=True)
    st.markdown("Convert JSON, CSV, TSV, XML, YAML files to Excel format with formatting and statistics")
    st.markdown("---")
    
    # Sidebar
    st.sidebar.header("🛠️ Options")
    data_source = st.sidebar.radio(
        "Choose Data Source:",
        ["Upload Data File", "Paste Data"]
    )
    
    include_formatting = st.sidebar.checkbox("Include Excel Formatting", value=True)
    show_statistics = st.sidebar.checkbox("Show Data Statistics", value=True)
    
    # Supported formats info
    st.sidebar.markdown("### 📁 Supported Formats")
    st.sidebar.markdown("""
    - **JSON** (.json) - JavaScript Object Notation
    - **CSV** (.csv) - Comma Separated Values
    - **TSV** (.tsv) - Tab Separated Values
    - **XML** (.xml) - Extensible Markup Language
    - **YAML** (.yaml, .yml) - YAML Ain't Markup Language
    - **TXT** (.txt) - Plain text (auto-detect format)
    """)
    
    # Initialize session state
    if 'processed_data' not in st.session_state:
        st.session_state.processed_data = None
    if 'df' not in st.session_state:
        st.session_state.df = None
    if 'file_info' not in st.session_state:
        st.session_state.file_info = None
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📋 Data Input")
        
        data = None
        
        if data_source == "Upload Data File":
            st.markdown('<div class="upload-section">', unsafe_allow_html=True)
            uploaded_file = st.file_uploader(
                "Choose a data file",
                type=['json', 'csv', 'tsv', 'xml', 'yaml', 'yml', 'txt'],
                help="Upload a file containing structured data"
            )
            st.markdown('</div>', unsafe_allow_html=True)
            
            # Show format examples
            with st.expander("📄 View Supported Format Examples"):
                format_tab1, format_tab2, format_tab3, format_tab4 = st.tabs(["JSON", "CSV", "XML", "YAML"])
                
                with format_tab1:
                    st.code('''[
  {"name": "John", "age": 30, "city": "New York"},
  {"name": "Jane", "age": 25, "city": "Los Angeles"}
]''', language='json')
                
                with format_tab2:
                    st.code('''name,age,city
John,30,New York
Jane,25,Los Angeles''', language='csv')
                
                with format_tab3:
                    st.code('''<data>
  <record>
    <name>John</name>
    <age>30</age>
    <city>New York</city>
  </record>
  <record>
    <name>Jane</name>
    <age>25</age>
    <city>Los Angeles</city>
  </record>
</data>''', language='xml')
                
                with format_tab4:
                    st.code('''- name: John
  age: 30
  city: New York
- name: Jane
  age: 25
  city: Los Angeles''', language='yaml')
            
            if uploaded_file is not None:
                data, error = process_uploaded_file(uploaded_file)
                
                if data is not None:
                    st.success(f"✅ File uploaded successfully: {uploaded_file.name}")
                    st.session_state.file_info = {
                        'name': uploaded_file.name,
                        'size': uploaded_file.size,
                        'type': uploaded_file.name.split('.')[-1].upper(),
                        'detected_format': uploaded_file.name.split('.')[-1].upper()
                    }
                else:
                    st.error(f"❌ {error}")
        
        elif data_source == "Paste Data":
            data_text = st.text_area(
                "Paste your data here:",
                height=300,
                placeholder='Paste your data in any supported format (JSON, CSV, TSV, XML, YAML)...'
            )
            
            if data_text.strip():
                # Detect format dynamically
                detected_format = detect_format_from_content(data_text)
                
                # Show detected format
                if detected_format != "Unknown":
                    st.info(f"🔍 Detected format: **{detected_format}**")
                else:
                    st.warning("⚠️ Unable to detect format automatically. Trying multiple parsers...")
                
                # Try parsing with detected format first, then fallback to others
                data = None
                error = None
                
                if detected_format == "JSON":
                    data, error = parse_json_file(data_text)
                elif detected_format == "CSV":
                    data, error = parse_csv_file(data_text)
                elif detected_format == "TSV":
                    data, error = parse_tsv_file(data_text)
                elif detected_format == "XML":
                    data, error = parse_xml_file(data_text)
                elif detected_format == "YAML":
                    data, error = parse_yaml_file(data_text)
                else:
                    # Try all formats if detection failed
                    parsers = [
                        ("JSON", parse_json_file),
                        ("CSV", parse_csv_file),
                        ("TSV", parse_tsv_file),
                        ("XML", parse_xml_file),
                        ("YAML", parse_yaml_file)
                    ]
                    
                    for format_name, parser in parsers:
                        data, error = parser(data_text)
                        if data is not None:
                            st.info(f"✅ Successfully parsed as **{format_name}**")
                            break
                
                if data is not None:
                    st.success("✅ Data parsed successfully!")
                else:
                    st.error(f"❌ {error}")
        
        
        # Process data
        if data is not None:
            is_valid, message = validate_data(data)
            
            if is_valid:
                st.session_state.processed_data = data
                st.session_state.df = pd.DataFrame(data)
                
                st.markdown('<div class="success-box">✅ Data processed successfully!</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="error-box">❌ {message}</div>', unsafe_allow_html=True)
    
    with col2:
        st.header("📊 Data Info")
        
        # File info
        if st.session_state.file_info:
            st.markdown('<div class="format-info">', unsafe_allow_html=True)
            st.write(f"**📁 File:** {st.session_state.file_info['name']}")
            st.write(f"**📏 Size:** {st.session_state.file_info['size']:,} bytes")
            st.write(f"**🏷️ Format:** {st.session_state.file_info['detected_format']}")
            st.markdown('</div>', unsafe_allow_html=True)
        
        if st.session_state.df is not None:
            df = st.session_state.df
            
            # Statistics
            st.markdown('<div class="stats-container">', unsafe_allow_html=True)
            st.metric("📋 Total Records", len(df))
            st.metric("📊 Total Columns", len(df.columns))
            st.metric("💾 Memory Usage", f"{df.memory_usage(deep=True).sum() / 1024:.2f} KB")
            st.markdown('</div>', unsafe_allow_html=True)
            
            if show_statistics:
                st.subheader("📈 Column Info")
                col_info = []
                for col in df.columns:
                    col_info.append({
                        "Column": col,
                        "Type": str(df[col].dtype),
                        "Non-Null": df[col].count(),
                        "Null": df[col].isnull().sum()
                    })
                st.dataframe(pd.DataFrame(col_info), use_container_width=True)
    
    # Data preview and download section
    if st.session_state.df is not None:
        st.header("👀 Data Preview")
        
        # Show preview with pagination
        preview_rows = st.slider("Preview rows:", 5, min(100, len(st.session_state.df)), 10)
        st.dataframe(st.session_state.df.head(preview_rows), use_container_width=True)
        
        if len(st.session_state.df) > preview_rows:
            st.info(f"Showing {preview_rows} of {len(st.session_state.df)} total rows")
        
        # Numeric summary
        numeric_cols = st.session_state.df.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0 and show_statistics:
            st.subheader("🔢 Numeric Summary")
            st.dataframe(st.session_state.df[numeric_cols].describe(), use_container_width=True)
        
        # Download section
        st.header("📥 Download Excel File")
        
        col1, col2, col3 = st.columns([1, 1, 1])
        
        with col1:
            filename = st.text_input("Filename:", value=f"converted_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        with col2:
            if st.button("🚀 Generate Excel File", type="primary"):
                try:
                    excel_data = create_excel_file(st.session_state.df, filename)
                    st.session_state.excel_data = excel_data
                    st.success("✅ Excel file generated successfully!")
                except Exception as e:
                    st.error(f"❌ Error generating Excel file: {e}")
        
        with col3:
            if 'excel_data' in st.session_state:
                st.download_button(
                    label="📥 Download Excel",
                    data=st.session_state.excel_data.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    # Footer
    st.markdown("---")
    st.markdown(
        "💡 **Tips:** Upload data files or paste formatted data. "
        "The app automatically detects and validates your data format, then creates beautifully formatted Excel files with statistics."
    )

if __name__ == "__main__":
    main()