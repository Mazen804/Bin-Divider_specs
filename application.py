import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Set page configuration
st.set_page_config(page_title="Bin Divider Specification Generator", page_icon=":package:", layout="wide")

# Title
st.title("Bin Divider Specification Generator")
st.write("Enter bin divider specifications to generate an Excel file matching the provided format.")

# Initialize session state for storing groups
if 'groups' not in st.session_state:
    st.session_state.groups = []

# Function to calculate derived fields
def calculate_fields(group_data, bin_data):
    # Calculate # of Aisles
    bin_data['# of Aisles'] = group_data['End Aisle'] - group_data['Start Aisle'] + 1
    # Calculate Qty Per Bay
    bin_data['Qty Per Bay'] = bin_data['# of Shelves per Bay'] * bin_data['Qty bins per Shelf']
    # Calculate Total Quantity
    bin_data['Total Quantity'] = bin_data['Qty Per Bay'] * group_data['# of Bays'] * bin_data['# of Aisles']
    # Calculate Bin Gross CBM
    bin_data['Bin Gross CBM'] = (bin_data['Depth (mm)'] * bin_data['Height (mm)'] * bin_data['Width (mm)']) / 1_000_000
    # Calculate Bin Net CBM
    bin_data['Bin Net CBM'] = bin_data['Bin Gross CBM'] * bin_data['UT']
    return bin_data

# Function to generate Excel file
def generate_excel(groups):
    df = pd.DataFrame()
    columns = [
        'Floor', 'Mod', 'Depth', 'Start Aisle', 'End Aisle', '# of Aisles', '# of Bays',
        'Total # of Shelves per Bay', 'Bay Design', 'Bin Box Type', 'Depth (mm)',
        'Height (mm)', 'Width (mm)', 'Lip (cm)', '# of Shelves per Bay',
        'Qty bins per Shelf', 'Qty Per Bay', 'Total Quantity', 'UT',
        'Bin Gross CBM', 'Bin Net CBM'
    ]
    
    for group in groups:
        group_data = group['group_data']
        for bin_data in group['bins']:
            row = {**group_data, **bin_data}
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    
    # Set column order
    df = df[columns]
    
    # Create Excel file in memory
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Bin Box"
    
    # Write DataFrame to Excel
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# Form for adding a new group
st.subheader("Add New Group")
with st.form(key='group_form'):
    floor = st.text_input("Floor", value="P-1")
    mod = st.text_input("Mod", value="H")
    depth = st.text_input("Depth", value="600mm")
    start_aisle = st.number_input("Start Aisle", min_value=1, value=200, step=1)
    end_aisle = st.number_input("End Aisle", min_value=1, value=200, step=1)
    num_bays = st.number_input("# of Bays", min_value=1, value=12, step=1)
    total_shelves_per_bay = st.number_input("Total # of Shelves per Bay", min_value=1, value=5, step=1)
    bay_design = st.text_input("Bay Design", value="60 Deep HRV")
    
    # Submit button for group
    submit_group = st.form_submit_button("Add Group")
    
    if submit_group:
        group_data = {
            'Floor': floor,
            'Mod': mod,
            'Depth': depth,
            'Start Aisle': start_aisle,
            'End Aisle': end_aisle,
            '# of Bays': num_bays,
            'Total # of Shelves per Bay': total_shelves_per_bay,
            'Bay Design': bay_design
        }
        st.session_state.groups.append({'group_data': group_data, 'bins': []})
        st.success("Group added! Now add bin box types below.")

# Form for adding bin box types to the latest group
if st.session_state.groups:
    st.subheader("Add Bin Box Type to Latest Group")
    with st.form(key='bin_form'):
        bin_box_type = st.text_input("Bin Box Type", value="60Deep w/o lip 600*440")
        depth_mm = st.number_input("Depth (mm)", min_value=0.0, value=600.0, step=0.1)
        height_mm = st.number_input("Height (mm)", min_value=0.0, value=440.0, step=0.1)
        width_mm = st.number_input("Width (mm)", min_value=0.0, value=464.4, step=0.1)
        lip_cm = st.number_input("Lip (cm)", min_value=0.0, value=0.0, step=0.1)
        shelves_per_bay = st.number_input("# of Shelves per Bay", min_value=1, value=4, step=1)
        qty_bins_per_shelf = st.number_input("Qty bins per Shelf", min_value=1, value=3, step=1)
        ut = st.number_input("UT", min_value=0.0, value=0.525, step=0.01)
        
        # Submit button for bin
        submit_bin = st.form_submit_button("Add Bin Box Type")
        
        if submit_bin:
            bin_data = {
                'Bin Box Type': bin_box_type,
                'Depth (mm)': depth_mm,
                'Height (mm)': height_mm,
                'Width (mm)': width_mm,
                'Lip (cm)': lip_cm,
                '# of Shelves per Bay': shelves_per_bay,
                'Qty bins per Shelf': qty_bins_per_shelf,
                'UT': ut
            }
            # Calculate derived fields
            bin_data = calculate_fields(st.session_state.groups[-1]['group_data'], bin_data)
            st.session_state.groups[-1]['bins'].append(bin_data)
            st.success("Bin box type added!")

# Display current groups and bins
if st.session_state.groups:
    st.subheader("Current Data")
    for i, group in enumerate(st.session_state.groups):
        st.write(f"**Group {i+1}:** {group['group_data']}")
        for j, bin_data in enumerate(group['bins']):
            st.write(f"  Bin {j+1}: {bin_data}")
    
    # Generate and download Excel file
    excel_data = generate_excel(st.session_state.groups)
    st.download_button(
        label="Download Excel File",
        data=excel_data,
        file_name="Bin_Divider_Specification.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Clear all data
if st.button("Clear All Data"):
    st.session_state.groups = []
    st.experimental_rerun()