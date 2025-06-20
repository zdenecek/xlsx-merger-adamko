import datetime
import streamlit as st
import pandas as pd
import json
import os
from io import BytesIO
from typing import List, Dict, Any, Optional, Tuple, Union
from streamlit_sortables import sort_items

from data_quality import validate_data, schema

def remove_duplicates(lst: List[Any]) -> List[Any]:
    """
    Remove duplicates from a list while preserving order.
    
    Args:
        lst: List that may contain duplicate elements
        
    Returns:
        List with duplicates removed, preserving original order
    """
    return list(dict.fromkeys(lst))

def save_configuration(config: Dict[str, Any], filename: str = 'config.json') -> None:
    """
    Save configuration dictionary to a JSON file.
    
    Args:
        config: Configuration dictionary to save
        filename: Name of the file to save configuration to
    """
    with open(filename, 'w') as f:
        json.dump(config, f)
    st.success(f'Configuration saved to {filename}')

def load_configuration(file: Any) -> Optional[Dict[str, Any]]:
    """
    Load configuration from an uploaded JSON file.
    
    Args:
        file: Uploaded file object from Streamlit file uploader
        
    Returns:
        Configuration dictionary if successful, None if failed
    """
    try:
        # Read the configuration from the uploaded JSON file
        config = json.load(file)
        st.success('Configuration loaded successfully!')
        return config
    except Exception as e:
        st.error(f'Error loading configuration: {e}')
        return None

def merge_files(files: List[Any], data_header_idx: int, data_col_idx: int, 
                selected_headers: List[str], order: str) -> pd.DataFrame:
    """
    Merge data from multiple Excel files based on specified columns.
    
    Args:
        files: List of uploaded file objects
        data_header_idx: Column index containing data headers
        data_col_idx: Column index containing data values
        selected_headers: List of headers to include in merged data
        order: Sorting order for files ("By filename" or "By last word in filename")
        
    Returns:
        Merged DataFrame with filename as first column and selected data
    """
    merged_data = []

    if order == "By last word in filename":
        files.sort(key=lambda file: file.name.split()[-1])
    if order == "By filename":
        files.sort(key=lambda file: file.name)

    for idx, file in enumerate(files):
        df = pd.read_excel(file, header=None)  # Read without headers
        # Handle potential mismatch in columns
        max_col_idx = max(data_header_idx, data_col_idx)
        if df.shape[1] <= max_col_idx:
            st.error(f'File "{file.name}" does not have enough columns.')
            return pd.DataFrame()  # Return empty DataFrame

        # Extract headers and data using column indexes
        headers = df.iloc[:, data_header_idx].astype(str).tolist()
        data_values = df.iloc[:, data_col_idx].fillna('').astype(str).tolist()
        data_dict = dict(zip(headers, data_values))
        row = []
       
        # first is filename
        row.append(file.name) 
        headers = ["Filename" ] + selected_headers
        
        for header in selected_headers:
            value = data_dict.get(str(header), None)
            row.append(value)
        merged_data.append(row)

    merged_df = pd.DataFrame(merged_data, columns=headers)
    return merged_df

def main() -> None:
    """
    Main Streamlit application function for Excel Data Merger.
    
    This function creates a web interface that allows users to:
    - Upload multiple Excel files
    - Configure column mappings for headers and data
    - Select and reorder data headers
    - Merge files into a consolidated dataset
    - Validate data quality
    - Download results
    """
    st.set_page_config(page_title='Adamkův Excel Merger', page_icon=':bar_chart:', layout='wide')
    st.title('Excel Data Merger')
    
    st.write("""
This Streamlit app allows you to transpose and merge data from multiple Excel files based on specified columns for headers and data. Here's how it works:

- **Upload Multiple Excel Files:** Select and upload multiple `.xlsx` files that you want to merge.
- **Specify Data Columns:** Define which columns in your Excel files contain the data headers and the actual data values by selecting their column indices.
- **Select and Reorder Headers:** Choose the specific data headers you want to include in the merged dataset and reorder them according to your preferences.
- **Merge Data:** Combine the selected data from all uploaded files into a single, consolidated dataset.
- **Edit Merged Data:** Preview and make any necessary edits to the merged data directly within the app before finalizing.
- **Validate Data Quality:** Run data quality checks against a predefined schema to ensure the merged data meets the required standards.
- **Download Results:** Easily download the merged and validated data as an Excel file for further use.
- **Save and Load Configurations:** Save your column selections and header configurations as a JSON file to reuse settings in future sessions.
""")

    # Initialize session state variables
    if 'data_header_idx' not in st.session_state:
        st.session_state['data_header_idx'] = None
    if 'data_col_idx' not in st.session_state:
        st.session_state['data_col_idx'] = None
    if 'selected_headers' not in st.session_state:
        st.session_state['selected_headers'] = []
    elif st.session_state['selected_headers']:
        # Ensure no duplicates in existing session state
        st.session_state['selected_headers'] = remove_duplicates(st.session_state['selected_headers'])
    if 'config_loaded' not in st.session_state:
        st.session_state['config_loaded'] = False

    st.header('Load Configuration')
    config_file = st.file_uploader('Choose a configuration file to load', type=['json'])
    if config_file is not None:
        config = load_configuration(config_file)
        if config:
            st.session_state['data_header_idx'] = config['data_header_idx']
            st.session_state['data_col_idx'] = config['data_col_idx']
            # Filter duplicates from loaded config while preserving order
            loaded_headers = config.get('selected_headers', [])
            st.session_state['selected_headers'] = remove_duplicates(loaded_headers) if loaded_headers else []
            st.session_state['order'] = config.get('order', "By filename")
            st.session_state['config_loaded'] = True



    st.header('Upload Files')
    # 1. File Upload
    uploaded_files = st.file_uploader(
        "Choose Excel files to merge",
        accept_multiple_files=True,
        type=['xlsx'],
    )

    if uploaded_files:
        st.header('Step 1: Specify Columns by Index')
        sample_file = uploaded_files[0]
        df_sample = pd.read_excel(sample_file, header=None)
        num_columns = df_sample.shape[1]

        column_examples = [
            f"Index {i}: {df_sample.iloc[0, i]}" for i in range(num_columns)
        ]

        # Select data_header_idx
        data_header_idx_options = list(range(num_columns))

        # Determine index for data_header_idx
        if st.session_state['data_header_idx'] is not None and st.session_state['data_header_idx'] in data_header_idx_options:
            data_header_idx_index = data_header_idx_options.index(st.session_state['data_header_idx'])
        else:
            data_header_idx_index = 0  # Default value

        data_header_idx = st.selectbox(
            'Select the column index that contains data headers',
            options=data_header_idx_options,
            format_func=lambda x: column_examples[x],
            index=data_header_idx_index,
        )
        st.session_state['data_header_idx'] = data_header_idx

        # Select data_col_idx
        # Determine index for data_col_idx
        if st.session_state['data_col_idx'] is not None and st.session_state['data_col_idx'] in data_header_idx_options:
            data_col_idx_index = data_header_idx_options.index(st.session_state['data_col_idx'])
        else:
            # Default to the next column index, if possible
            default_col_idx = data_header_idx + 1 if data_header_idx + 1 < num_columns else data_header_idx
            data_col_idx_index = data_header_idx_options.index(default_col_idx)

        data_col_idx = st.selectbox(
            'Select the column index that contains data',
            options=data_header_idx_options,
            format_func=lambda x: column_examples[x],
            index=min(data_header_idx + 1, len(data_header_idx_options) - 1),
        )
        st.session_state['data_col_idx'] = data_col_idx

        st.header('Step 2: Select Data Headers')
        all_headers = []
        for file in uploaded_files:
            df = pd.read_excel(file, header=None)
            # Ensure the file has enough columns
            if df.shape[1] <= max(data_header_idx, data_col_idx):
                st.error(f"File \"{file.name}\" does not have enough columns.")
                return
            headers = df.iloc[:, data_header_idx].dropna().astype(str).tolist()
            all_headers.extend(headers)

        # Remove duplicates while preserving order
        all_headers = remove_duplicates(all_headers)

        # Create a DataFrame for headers with a selection column
        headers_df = pd.DataFrame({'Header': all_headers})

        # Determine which headers are selected based on session state
        if st.session_state['selected_headers']:
            # Filter duplicates from session state while preserving order
            session_headers = remove_duplicates(st.session_state['selected_headers'])
            st.session_state['selected_headers'] = session_headers  # Update session state with filtered headers
            headers_df['Select'] = headers_df['Header'].isin(session_headers)
        else:
            headers_df['Select'] = True  # Default to all selected

        # Allow user to deselect headers
        edited_df = st.data_editor(
            headers_df,
            use_container_width=True,
            key='headers_editor',
            disabled=["Header"]  # Disable editing of the 'Header' column
        )



        # Update selected_headers in session state
        selected_headers = edited_df[edited_df['Select']]['Header'].tolist()
        # Filter duplicates while preserving order
        selected_headers = remove_duplicates(selected_headers)


        # Allow reordering of selected headers using sort_items
        st.header('Step 3: Reorder Data Headers')
        if selected_headers:
            st.write('Drag to reorder the selected headers:')
            # Use the order from session state if available, otherwise use the current order
            if st.session_state['selected_headers']:
                # Preserve the order from session state for headers that exist in both lists
                unsorted_headers = [h for h in st.session_state['selected_headers'] if h in selected_headers]
                # Add any new headers that weren't in session state
                unsorted_headers.extend([h for h in selected_headers if h not in unsorted_headers])
                # Filter duplicates while preserving order
                unsorted_headers = remove_duplicates(unsorted_headers)
            else:
                unsorted_headers = selected_headers.copy()
                
            sorted_headers = sort_items(
                items=unsorted_headers,
                direction='horizontal',
                key=f'header_reordering_{hash(tuple(selected_headers))}'
            )

            # Remove duplicates while preserving order
            sorted_headers = remove_duplicates(sorted_headers)
            # Update sorted_headers in session state
        else:
            st.warning('Please select at least one header.')
            sorted_headers = []  # Initialize as empty list when no headers selected
	

        st.header('Step 4: Order rows')
        # Get the order from config if available, otherwise use default
        default_order = st.session_state.get('order', "By filename")
        order = st.selectbox(
            label="Order by", 
            options=["By filename", "By last word in filename"],
            index=0 if default_order == "By filename" else 1
        )
        st.session_state['order'] = order

        # Save Configuration
        st.header('Save Configuration')
        # Create configuration data
        config = {
            'data_header_idx': data_header_idx,
            'data_col_idx': data_col_idx,
            'selected_headers': sorted_headers,
            'order': order
        }
        config_json = json.dumps(config, indent=4)
        st.download_button(
            label='Download Configuration',
            data=config_json,
            file_name='config.json',
            mime='application/json'
        )
        
        st.header('Merge Files')
        
        
        if st.button('Merge Files'):
            if not sorted_headers:
                st.error('No headers selected. Please select at least one header before merging.')
            else:
                merged_df = merge_files(
                    uploaded_files,
                    data_header_idx,
                    data_col_idx,
                    sorted_headers,
                    order
                )
                if merged_df.empty:
                    st.error('Merging failed. No data to validate.')
                else:
                    st.session_state['merged_df'] = merged_df
                    st.session_state['sorted_headers'] = sorted_headers

        # Display merged data if available
        if 'merged_df' in st.session_state and st.session_state['merged_df'] is not None:
            merged_df = st.session_state['merged_df']
            sorted_headers = st.session_state['sorted_headers']
            st.success('Files merged successfully!')
            edited_data = st.data_editor(merged_df, use_container_width=True)
            
            filename = f"reports_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
            out_stream = BytesIO()
            edited_data.to_excel(out_stream, index=False)
            out_stream.seek(0)
            st.download_button(
                label="Download Merged Excel File",
                data=out_stream,
                file_name=filename,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
            st.header("Data quality checks")
            validated_df, validation_errors = validate_data(edited_data, schema)
            if validated_df is not None:
                st.success('Data validation passed!')
                # Display validated data
                st.subheader('Validated Data Preview')
                st.dataframe(validated_df)
                # Download Validated Merged File
                towrite = BytesIO()
                validated_df.to_excel(towrite, index=False)
                towrite.seek(0)
                st.download_button(
                    label="Download Validated Merged Excel File",
                    data=towrite,
                    file_name='validated_merged_data.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            else:
                st.error('Data validation failed. Please review the errors below.')
                # Display validation errors
                st.dataframe(validation_errors)
                # Optionally allow the user to download the errors
                error_buffer = BytesIO()
                validation_errors.to_excel(error_buffer, index=False)
                error_buffer.seek(0)
                st.download_button(
                    label="Download Validation Errors",
                    data=error_buffer,
                    file_name='validation_errors.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )



if __name__ == '__main__':
    main()