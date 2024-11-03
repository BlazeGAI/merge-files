import streamlit as st
import zipfile
import os
import shutil

# Dictionary of file types and their extensions
FILE_TYPES = {
    'Excel Files': ['.xlsx', '.xls'],
    'Word Files': ['.docx', '.doc'],
    'PowerPoint Files': ['.pptx', '.ppt']
}

def process_zip_file(uploaded_zip, selected_file_type):
    # Get the file extensions for the selected type
    target_extensions = FILE_TYPES[selected_file_type]
    
    # Create a temporary directory to store files
    if not os.path.exists('temp_files'):
        os.makedirs('temp_files')
    
    # Create a directory for the final zip file
    if not os.path.exists('output'):
        os.makedirs('output')
    
    # Read and extract the uploaded zip file
    with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
        # Get list of all files in zip
        file_list = zip_ref.namelist()
        
        # Extract only files with target extensions
        for file in file_list:
            if any(file.lower().endswith(ext) for ext in target_extensions):
                zip_ref.extract(file, 'temp_files')
    
    # Create new zip file with extracted files
    output_zip_path = os.path.join('output', f'extracted_{selected_file_type.lower().replace(" ", "_")}.zip')
    with zipfile.ZipFile(output_zip_path, 'w') as zipf:
        for root, dirs, files in os.walk('temp_files'):
            for file in files:
                if any(file.lower().endswith(ext) for ext in target_extensions):
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, os.path.basename(file_path))
    
    return output_zip_path

# Streamlit app interface
st.title('Office File Extractor')
st.write('Upload a zip file and select the type of files you want to extract')

# File type selector
selected_type = st.selectbox(
    'What type of files do you want to extract?',
    list(FILE_TYPES.keys())
)

# File uploader
uploaded_file = st.file_uploader("Choose a ZIP file", type="zip")

if uploaded_file is not None:
    if st.button('Extract Files'):
        try:
            # Process the zip file
            output_zip = process_zip_file(uploaded_file, selected_type)
            
            # Read the output zip file for download
            with open(output_zip, 'rb') as f:
                st.download_button(
                    label=f'Download Extracted {selected_type}',
                    data=f,
                    file_name=f'extracted_{selected_type.lower().replace(" ", "_")}.zip',
                    mime='application/zip'
                )
            
            # Show success message
            st.success(f'{selected_type} have been extracted successfully!')
            
        except Exception as e:
            st.error(f'An error occurred: {str(e)}')
            
        finally:
            # Clean up temporary files
            if os.path.exists('temp_files'):
                shutil.rmtree('temp_files')
            if os.path.exists('output'):
                shutil.rmtree('output')

# Add some helpful information
st.markdown("""
---
### Supported File Types:
- Excel: .xlsx, .xls
- Word: .docx, .doc
- PowerPoint: .pptx, .ppt
""")