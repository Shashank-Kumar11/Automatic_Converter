#!/usr/bin/env python3
"""
Streamlit Dynamic Data Format Converter
A local web app for converting between various data formats with user-selectable input and output formats.

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
    page_title="Dynamic Data Format Converter",
    page_icon="🔄",
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
    .format-selector {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border-radius: 10px;
        padding: 1rem;
        margin: 1rem 0;
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
    .conversion-flow {
        background: linear-gradient(90deg, #4CAF50, #2196F3);
        color: white;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Input format parsers
def parse_json_data(content):
    """Parse JSON data"""
    try:
        return json.loads(content), None
    except json.JSONDecodeError as e:
        return None, f"Invalid JSON format: {e}"

def parse_csv_data(content, delimiter=','):
    """Parse CSV data"""
    try:
        df = pd.read_csv(io.StringIO(content), delimiter=delimiter)
        return df.to_dict('records'), None
    except Exception as e:
        return None, f"Error parsing CSV: {e}"

def parse_tsv_data(content):
    """Parse TSV data"""
    return parse_csv_data(content, delimiter='\t')

def parse_xml_data(content):
    """Parse XML data"""
    try:
        root = ET.fromstring(content)
        data = []
        
        # Handle different XML structures
        children = list(root)
        if len(children) > 0:
            for child in children:
                record = {}
                if len(list(child)) > 0:
                    for subchild in child:
                        record[subchild.tag] = subchild.text or ""
                else:
                    record[child.tag] = child.text or ""
                data.append(record)
        else:
            record = {}
            for child in root:
                record[child.tag] = child.text or ""
            data.append(record)
        
        return data, None
    except ET.ParseError as e:
        return None, f"Invalid XML format: {e}"

def parse_yaml_data(content):
    """Parse YAML data"""
    try:
        data = yaml.safe_load(content)
        if isinstance(data, list):
            return data, None
        elif isinstance(data, dict):
            return [data], None
        else:
            return None, "YAML must contain a list of objects or a single object"
    except yaml.YAMLError as e:
        return None, f"Invalid YAML format: {e}"

# Output format generators
def generate_json_output(data):
    """Generate JSON output"""
    return json.dumps(data, indent=2, ensure_ascii=False), "application/json", ".json"

def generate_csv_output(df):
    """Generate CSV output"""
    output = io.StringIO()
    df.to_csv(output, index=False)
    return output.getvalue(), "text/csv", ".csv"

def generate_tsv_output(df):
    """Generate TSV output"""
    output = io.StringIO()
    df.to_csv(output, index=False, sep='\t')
    return output.getvalue(), "text/tab-separated-values", ".tsv"

def generate_xml_output(data):
    """Generate XML output"""
    root = ET.Element("data")
    
    for i, record in enumerate(data):
        record_elem = ET.SubElement(root, f"record_{i+1}")
        for key, value in record.items():
            field_elem = ET.SubElement(record_elem, str(key).replace(' ', '_'))
            field_elem.text = str(value) if value is not None else ""
    
    # Convert to string with proper formatting
    rough_string = ET.tostring(root, encoding='unicode')
    try:
        import xml.dom.minidom
        dom = xml.dom.minidom.parseString(rough_string)
        return dom.toprettyxml(indent="  "), "application/xml", ".xml"
    except:
        return rough_string, "application/xml", ".xml"

def generate_yaml_output(data):
    """Generate YAML output"""
    return yaml.dump(data, default_flow_style=False, allow_unicode=True), "application/x-yaml", ".yaml"

def generate_excel_output(df):
    """Generate Excel output with formatting"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Data', index=False)
        
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
    return output.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".xlsx"

def generate_html_output(df):
    """Generate HTML table output"""
    html_string = df.to_html(index=False, classes='table table-striped table-bordered', escape=False)
    
    # Add Bootstrap styling
    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Data Export</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body {{ padding: 20px; }}
            .table {{ margin-top: 20px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Exported Data</h1>
            <p>Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
            {html_string}
        </div>
    </body>
    </html>
    """
    return full_html, "text/html", ".html"

def process_data_with_format(content, input_format):
    """Process data based on selected input format"""
    parsers = {
        'JSON': parse_json_data,
        'CSV': parse_csv_data,
        'TSV': parse_tsv_data,
        'XML': parse_xml_data,
        'YAML': parse_yaml_data
    }
    
    parser = parsers.get(input_format)
    if parser:
        return parser(content)
    else:
        return None, f"Unsupported input format: {input_format}"

def convert_to_output_format(data, df, output_format):
    """Convert data to selected output format"""
    generators = {
        'JSON': lambda: generate_json_output(data),
        'CSV': lambda: generate_csv_output(df),
        'TSV': lambda: generate_tsv_output(df),
        'XML': lambda: generate_xml_output(data),
        'YAML': lambda: generate_yaml_output(data),
        'Excel': lambda: generate_excel_output(df),
        'HTML': lambda: generate_html_output(df)
    }
    
    generator = generators.get(output_format)
    if generator:
        return generator()
    else:
        return None, None, None

def main():
    """Main Streamlit app"""
    
    # Header
    st.markdown('<h1 class="main-header">🔄 Dynamic Data Format Converter</h1>', unsafe_allow_html=True)
    st.markdown("**Convert between JSON, CSV, TSV, XML, YAML, Excel, and HTML formats with full user control**")
    st.markdown("---")
    
    # Initialize session state
    if 'processed_data' not in st.session_state:
        st.session_state.processed_data = None
    if 'df' not in st.session_state:
        st.session_state.df = None
    
    # Sidebar - Format Selection
    st.sidebar.header("🎛️ Format Configuration")
    
    # Input format selection
    st.sidebar.markdown("### 📥 Input Format")
    input_format = st.sidebar.selectbox(
        "Select input data format:",
        ["JSON", "CSV", "TSV", "XML", "YAML"],
        help="Choose the format of your input data"
    )
    
    # Output format selection
    st.sidebar.markdown("### 📤 Output Format")
    output_format = st.sidebar.selectbox(
        "Select output format:",
        ["Excel", "JSON", "CSV", "TSV", "XML", "YAML", "HTML"],
        help="Choose the desired output format"
    )
    
    # Show conversion flow
    st.sidebar.markdown('<div class="conversion-flow">', unsafe_allow_html=True)
    st.sidebar.markdown(f"**{input_format}** ➡️ **{output_format}**")
    st.sidebar.markdown('</div>', unsafe_allow_html=True)
    
    # Data source selection
    st.sidebar.markdown("### 📋 Data Source")
    data_source = st.sidebar.radio(
        "Choose data source:",
        ["Upload File", "Paste Data"]
    )
    
    # Additional options
    st.sidebar.markdown("### ⚙️ Options")
    show_preview = st.sidebar.checkbox("Show data preview", value=True)
    show_statistics = st.sidebar.checkbox("Show statistics", value=True)
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📋 Data Input")
        
        # Format information
        st.markdown('<div class="format-selector">', unsafe_allow_html=True)
        st.markdown(f"**Input Format:** {input_format} ➡️ **Output Format:** {output_format}")
        st.markdown('</div>', unsafe_allow_html=True)
        
        data = None
        
        if data_source == "Upload File":
            st.markdown('<div class="upload-section">', unsafe_allow_html=True)
            
            # Dynamic file extensions based on input format
            file_extensions = {
                'JSON': ['json'],
                'CSV': ['csv'],
                'TSV': ['tsv', 'txt'],
                'XML': ['xml'],
                'YAML': ['yaml', 'yml']
            }
            
            uploaded_file = st.file_uploader(
                f"Upload {input_format} file:",
                type=file_extensions.get(input_format, ['txt']),
                help=f"Upload a file in {input_format} format"
            )
            st.markdown('</div>', unsafe_allow_html=True)
            
            if uploaded_file is not None:
                try:
                    file_content = uploaded_file.read().decode('utf-8')
                    data, error = process_data_with_format(file_content, input_format)
                    
                    if data is not None:
                        st.success(f"✅ {input_format} file uploaded and parsed successfully!")
                        st.session_state.file_info = {
                            'name': uploaded_file.name,
                            'size': uploaded_file.size,
                            'input_format': input_format,
                            'output_format': output_format
                        }
                    else:
                        st.error(f"❌ {error}")
                        
                except UnicodeDecodeError:
                    st.error("❌ Unable to decode file. Please ensure it's a text file with UTF-8 encoding.")
                except Exception as e:
                    st.error(f"❌ Error processing file: {e}")
        
        elif data_source == "Paste Data":
            st.markdown(f"**Paste your {input_format} data below:**")
            
            # Format-specific placeholders
            placeholders = {
                'JSON': '[{"key": "value"}, {"key": "value"}]',
                'CSV': 'name,age,city\nJohn,30,New York\nJane,25,Los Angeles',
                'TSV': 'name\tage\tcity\nJohn\t30\tNew York\nJane\t25\tLos Angeles',
                'XML': '<data><record><name>John</name><age>30</age></record></data>',
                'YAML': '- name: John\n  age: 30\n- name: Jane\n  age: 25'
            }
            
            data_text = st.text_area(
                f"{input_format} Data:",
                height=300,
                placeholder=placeholders.get(input_format, 'Enter your data here...')
            )
            
            if data_text.strip():
                data, error = process_data_with_format(data_text, input_format)
                
                if data is not None:
                    st.success(f"✅ {input_format} data parsed successfully!")
                else:
                    st.error(f"❌ {error}")
        
        # Process and validate data
        if data is not None:
            if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
                st.session_state.processed_data = data
                st.session_state.df = pd.DataFrame(data)
                st.markdown('<div class="success-box">✅ Data processed and ready for conversion!</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="error-box">❌ Data must be a list of objects/dictionaries</div>', unsafe_allow_html=True)
    
    with col2:
        st.header("📊 Data Info")
        
        # File info
        if hasattr(st.session_state, 'file_info') and st.session_state.file_info:
            st.markdown('<div class="format-info">', unsafe_allow_html=True)
            st.write(f"**📁 File:** {st.session_state.file_info['name']}")
            st.write(f"**📏 Size:** {st.session_state.file_info['size']:,} bytes")
            st.write(f"**📥 Input:** {st.session_state.file_info['input_format']}")
            st.write(f"**📤 Output:** {st.session_state.file_info['output_format']}")
            st.markdown('</div>', unsafe_allow_html=True)
        
        if st.session_state.df is not None:
            df = st.session_state.df
            
            # Statistics
            st.markdown('<div class="stats-container">', unsafe_allow_html=True)
            st.metric("📋 Records", len(df))
            st.metric("📊 Columns", len(df.columns))
            st.metric("💾 Memory", f"{df.memory_usage(deep=True).sum() / 1024:.1f} KB")
            st.markdown('</div>', unsafe_allow_html=True)
            
            if show_statistics:
                st.subheader("📈 Column Details")
                col_info = []
                for col in df.columns:
                    col_info.append({
                        "Column": col,
                        "Type": str(df[col].dtype),
                        "Non-Null": df[col].count(),
                        "Null": df[col].isnull().sum()
                    })
                st.dataframe(pd.DataFrame(col_info), use_container_width=True, hide_index=True)
    
    # Data preview and conversion section
    if st.session_state.df is not None:
        if show_preview:
            st.header("👀 Data Preview")
            preview_rows = st.slider("Preview rows:", 5, min(50, len(st.session_state.df)), 10)
            st.dataframe(st.session_state.df.head(preview_rows), use_container_width=True)
            
            if len(st.session_state.df) > preview_rows:
                st.info(f"Showing {preview_rows} of {len(st.session_state.df)} total rows")
        
        # Conversion and download section
        st.header(f"🔄 Convert to {output_format}")
        
        col1, col2, col3 = st.columns([1, 1, 1])
        
        with col1:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            default_filename = f"converted_data_{timestamp}"
            filename = st.text_input("Filename (without extension):", value=default_filename)
        
        with col2:
            if st.button(f"🚀 Convert to {output_format}", type="primary"):
                try:
                    output_data, mime_type, extension = convert_to_output_format(
                        st.session_state.processed_data, 
                        st.session_state.df, 
                        output_format
                    )
                    
                    if output_data is not None:
                        st.session_state.converted_data = output_data
                        st.session_state.mime_type = mime_type
                        st.session_state.file_extension = extension
                        st.session_state.output_filename = f"{filename}{extension}"
                        st.success(f"✅ Successfully converted to {output_format}!")
                    else:
                        st.error(f"❌ Error converting to {output_format}")
                        
                except Exception as e:
                    st.error(f"❌ Conversion error: {e}")
        
        with col3:
            if hasattr(st.session_state, 'converted_data'):
                st.download_button(
                    label=f"📥 Download {output_format}",
                    data=st.session_state.converted_data,
                    file_name=st.session_state.output_filename,
                    mime=st.session_state.mime_type
                )
        
        # Show converted data preview for text formats
        if hasattr(st.session_state, 'converted_data') and output_format in ['JSON', 'CSV', 'TSV', 'XML', 'YAML', 'HTML']:
            st.subheader(f"📄 {output_format} Output Preview")
            if output_format == 'HTML':
                st.components.v1.html(st.session_state.converted_data, height=400, scrolling=True)
            else:
                # Show first 2000 characters for text formats
                preview_text = str(st.session_state.converted_data)[:2000]
                if len(str(st.session_state.converted_data)) > 2000:
                    preview_text += "\n... (truncated)"
                st.code(preview_text, language=output_format.lower())
    
    # Footer
    st.markdown("---")
    st.markdown(
        "💡 **How to use:** Select input format → Choose output format → Upload file or paste data → Convert → Download! "
        f"Currently converting **{input_format}** to **{output_format}**."
    )

if __name__ == "__main__":
    main()