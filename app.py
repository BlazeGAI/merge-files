import streamlit as st
import zipfile
import os
import shutil
import re

# Dictionary of file types and their extensions
FILE_TYPES = {
    'Excel Files': ['.xlsx', '.xls'],
    'Word Files': ['.docx', '.doc'],
    'PowerPoint Files': ['.pptx', '.ppt']
}

def clean_student_name(folder_name):
    # Extract just the student name from the folder pattern
    match = re.match(r"([^_]+_[^_]+)", folder_name)
    if match:
        # Get the name part and replace space with underscore
        return match.group(1).replace(" ", "_")
    return folder_name

def clean_file_name(file_name):
    # Remove 'assignsubmission_file_' from the filename if it exists
    return file_name.replace('assignsubmission_file_', '')

def process_zip_file(uploaded_zip, selected_file_type):
    # Get the file extensions for the selected type
    target_extensions = FILE_TYPES[selected_file_type]
    
    # Create a temporary directory to store files
    if not os.path.exists('temp_files'):
        os.makedirs('temp_files')
    
    # Create a directory for the final zip file
    if not os.path.exists('output'):
        os.makedirs('output')
    
    # Dictionary to store folder paths and their files
    folder_files = {}
    
    # Read and extract the uploaded zip file
    with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
        # Get list of all files in zip
        file_list = zip_ref.namelist()
        
        # Extract only files with target extensions
        for file in file_list:
            if any(file.lower().endswith(ext) for ext in target_extensions):
                # Get the folder path and filename
                folder_path = os.path.dirname(file)
                file_name = os.path.basename(file)
                
                # Skip if it's just a directory
                if not file_name:
                    continue
                
                # Extract the file
                zip_ref.extract(file, 'temp_files')
                
                # Store the folder path and filename
                if folder_path not in folder_files:
                    folder_files[folder_path] = []
                folder_files[folder_path].append((file, file_name))
    
    # Create new zip file with renamed files
    output_zip_path = os.path.join('output', f'extracted_{selected_file_type.lower().replace(" ", "_")}.zip')
    with zipfile.ZipFile(output_zip_path, 'w') as zipf:
        for folder_path, files in folder_files.items():
            # Get student name from folder path
            folder_name = os.path.basename(folder_path)
            student_name = clean_student_name(folder_name)
            
            for full_path, original_name in files:
                # Clean the original filename
                cleaned_original_name = clean_file_name(original_name)
                
                # Create new filename: StudentName_OriginalFileName
                new_name = f"{student_name}_{cleaned_original_name}"
                
                # Add file to zip with new name
                file_path = os.path.join('temp_files', full_path)
                zipf.write(file_path, new_name)
    
    return output_zip_path

# Streamlit app interface
st.title('BlazeGAI File Merger')
st.write('Upload a zip file containing student submissions and select the type of files to extract')

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

### File Naming:
Files will be renamed using the pattern: StudentName_FileName.extension
Example: For a student "Abigail Miller", a file will be renamed to "Abigail_Miller_assignment1.xlsx"
""")
