import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Set page configuration
st.set_page_config(page_title="Bin Divider Specification Generator", page_icon=":package:", layout="wide")

# Title
st.title("Bin Divider Specification Generator")
st.write("Enter bin divider specifications group by group to generate an Excel file.")

# Initialize session state
if 'groups' not in st.session_state:
    st.session_state.groups = []
if 'current_group' not in st.session_state:
    st.session_state.current_group = None
if 'bin_count' not in st.session_state:
    st.session_state.bin_count = 0
if 'bins' not in st.session_state:
    st.session_state.bins = []

# Function to calculate derived fields
def calculate_fields(group_data, bin_data):
    bin_data['# of Aisles'] = group_data['End Aisle'] - group_data['Start Aisle'] + 1
    bin_data['Qty Per Bay'] = bin_data['# of Shelves per Bay'] * bin_data['Qty bins per Shelf']
    bin_data['Total Quantity'] = bin_data['Qty Per Bay'] * group_data['# of Bays'] * bin_data['# of Aisles']
    bin_data['Bin Gross CBM'] = (bin_data['Depth (mm)'] * bin_data['Height (mm)'] * bin_data['Width (mm)']) / 1_000_000
    bin_data['Bin Net CBM'] = bin_data['Bin Gross CBM'] * bin_data['UT']
    return bin_data

# Function to generate Excel file
def generate_excel(groups):
    columns = [
        'Floor', 'Mod', 'Depth', 'Start Aisle', 'End Aisle', '# of Aisles', '# of Bays',
        'Total # of Shelves per Bay', 'Bay Design', 'Bin Box Type', 'Depth (mm)',
        'Height (mm)', 'Width (mm)', 'Lip (cm)', '# of Shelves per Bay',
        'Qty bins per Shelf', 'Qty Per Bay', 'Total Quantity', 'UT',
        'Bin Gross CBM', 'Bin Net CBM'
    ]
    
    # Initialize empty DataFrame with correct columns
    df = pd.DataFrame(columns=columns)
    
    # Populate DataFrame
    for group in groups:
        group_data = group['group_data']
        for bin_data in group['bins']:
            row = {**group_data, **bin_data}
            # Ensure all columns are present
            for col in columns:
                if col not in row:
                    row[col] = None
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    
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

# Add new group
if st.session_state.current_group is None:
    st.subheader("Add New Group (e.g., 60Deep HRV with 12 Bays)")
    with st.form(key='group_form'):
        floor = st.text_input("Floor", value="P-1")
        mod = st.text_input("Mod", value="H")
        depth = st.text_input("Depth", value="600mm")
        start_aisle = st.number_input("Start Aisle", min_value=1, value=200, step=1)
        end_aisle = st.number_input("End Aisle", min_value=1, value=200, step=1)
        num_bays = st.number_input("# of Bays", min_value=1, value=12, step=1)
        total_shelves_per_bay = st.number_input("Total # of Shelves per Bay", min_value=1, value=5, step=1)
        bay_design = st.text_input("Bay Design", value="60 Deep HRV")
        
        submit_group = st.form_submit_button("Add Group")
        
        if submit_group:
            st.session_state.current_group = {
                'Floor': floor,
                'Mod': mod,
                'Depth': depth,
                'Start Aisle': start_aisle,
                'End Aisle': end_aisle,
                '# of Bays': num_bays,
                'Total # of Shelves per Bay': total_shelves_per_bay,
                'Bay Design': bay_design
            }
            st.session_state.bins = []
            st.session_state.bin_count = 0
            st.success("Group added! Select the number of bin box types below.")

# Add bin box types for the current group
if st.session_state.current_group and st.session_state.bin_count == 0:
    st.subheader("Select Number of Bin Box Types")
    bin_count = st.selectbox("How many bin box types for this group?", options=[1, 2, 3, 4, 5], index=0)
    if st.button("Confirm Bin Box Types"):
        st.session_state.bin_count = bin_count
        st.session_state.bins = [{} for _ in range(bin_count)]
        st.success(f"Enter details for {bin_count} bin box type(s).")

# Input bin box types
if st.session_state.bin_count > 0:
    st.subheader(f"Enter {st.session_state.bin_count} Bin Box Type(s)")
    with st.form(key='bin_form'):
        for i in range(st.session_state.bin_count):
            st.write(f"**Bin Box Type {i+1}**")
            bin_box_type = st.text_input(f"Bin Box Type {i+1}", value="60Deep w/o lip 600*440", key=f"bin_type_{i}")
            depth_mm = st.number_input(f"Depth (mm) {i+1}", min_value=0.0, value=600.0, step=0.1, key=f"depth_{i}")
            height_mm = st.number_input(f"Height (mm) {i+1}", min_value=0.0, value=440.0, step=0.1, key=f"height_{i}")
            width_mm = st.number_input(f"Width (mm) {i+1}", min_value=0.0, value=464.4, step=0.1, key=f"width_{i}")
            lip_cm = st.number_input(f"Lip (cm) {i+1}", min_value=0.0, value=0.0, step=0.1, key=f"lip_{i}")
            shelves_per_bay = st.number_input(f"# of Shelves per Bay {i+1}", min_value=1, value=4, step=1, key=f"shelves_{i}")
            qty_bins_per_shelf = st.number_input(f"Qty bins per Shelf {i+1}", min_value=1, value=3, step=1, key=f"qty_bins_{i}")
            ut = st.number_input(f"UT {i+1}", min_value=0.0, value=0.525, step=0.01, key=f"ut_{i}")
            
            st.session_state.bins[i] = {
                'Bin Box Type': bin_box_type,
                'Depth (mm)': depth_mm,
                'Height (mm)': height_mm,
                'Width (mm)': width_mm,
                'Lip (cm)': lip_cm,
                '# of Shelves per Bay': shelves_per_bay,
                'Qty bins per Shelf': qty_bins_per_shelf,
                'UT': ut
            }
        
        submit_bins = st.form_submit_button("Finalize Group")
        
        if submit_bins:
            for i in range(st.session_state.bin_count):
                st.session_state.bins[i] = calculate_fields(st.session_state.current_group, st.session_state.bins[i])
            st.session_state.groups.append({
                'group_data': st.session_state.current_group,
                'bins': st.session_state.bins
            })
            st.session_state.current_group = None
            st.session_state.bin_count = 0
            st.session_state.bins = []
            st.success("Group finalized! Add a new group or download the Excel file.")

# Display current groups and bins
if st.session_state.groups:
    st.subheader("Finalized Groups")
    for i, group in enumerate(st.session_state.groups):
        st.write(f"**Group {i+1}:** {group['group_data']}")
        for j, bin_data in enumerate(group['bins']):
            st.write(f"  Bin {j+1}: {bin_data}")
    
    # Generate and download Excel file
    excel_data = generate_excel(st.session_state.groups)
    st.download_button(
        label="Download Excel File",
        data=excel_data,
        file_name="RUH5_Bin_Divider_Specification.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Clear all data
if st.button("Clear All Data"):
    st.session_state.groups = []
    st.session_state.current_group = None
    st.session_state.bin_count = 0
    st.session_state.bins = []
    st.experimental_rerun()