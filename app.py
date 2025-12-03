import streamlit as st
import pandas as pd
import itertools
import io

def find_combination(target, available_bookings):
    """
    Finds a combination of bookings that sum up to the target number of people.
    Restricted to max 2 bookings.
    Allows tolerance of +/- 1, and then tries to fill as much as possible.
    """
    # Define targets in order of preference: 
    # 1. Exact
    # 2. Overfill by 1
    # 3. Underfill by 1, 2, 3... down to 1 (Maximize output)
    
    targets_to_try = [target, target + 1] + list(range(target - 1, 0, -1))
    
    # Filter candidates once based on the maximum possible target
    max_target = max(targets_to_try)
    candidates = [b for b in available_bookings if b['No. of Person'] <= max_target]
    
    for t in targets_to_try:
        # Try to find a single booking first
        for b in candidates:
            if b['No. of Person'] == t:
                return [b]
                
        # Try combinations of 2
        # We need to filter candidates again for the specific target 't' to avoid overshooting
        # (e.g. if t=8, we don't want to pick a candidate with 9 people)
        current_candidates = [b for b in candidates if b['No. of Person'] < t]
        
        for pair in itertools.combinations(current_candidates, 2):
            if sum(b['No. of Person'] for b in pair) == t:
                return list(pair)
            
    return None

def process_allocation(file1, file2):
    try:
        # Load Book1
        df1 = pd.read_excel(file1)
        # Ensure 'No. of Person' is numeric
        df1['No. of Person'] = pd.to_numeric(df1['No. of Person'], errors='coerce').fillna(0).astype(int)
        
        # Initialize tracking columns for Book1
        df1['Allocation Status'] = "Not Allocated"
        df1['Allotted Dhaja No'] = None
        
        # Load Book2 (All Sheets)
        # Returns a dictionary of DataFrames
        xls2 = pd.read_excel(file2, sheet_name=None, header=1)
        
    except Exception as e:
        st.error(f"Error loading files: {e}")
        return None, None

    # Add index to track rows
    df1['original_index'] = df1.index
    bookings = df1.to_dict('records')
    
    # Helper to mark as assigned in the list
    def mark_assigned(booking_idx, dhaja_no):
        for b in bookings:
            if b['original_index'] == booking_idx:
                b['Allocation Status'] = "Allocated"
                b['Allotted Dhaja No'] = dhaja_no
                break
    
    processed_sheets = {}
    
    # Iterate through each sheet in Book2
    progress_bar = st.progress(0)
    total_sheets = len(xls2)
    current_sheet_idx = 0
    
    for sheet_name, df2 in xls2.items():
        # Ensure 'test' is numeric
        if 'test' not in df2.columns:
             st.warning(f"Sheet '{sheet_name}' does not have a 'test' column. Skipping.")
             processed_sheets[sheet_name] = df2
             continue
             
        df2['test'] = pd.to_numeric(df2['test'], errors='coerce').fillna(0).astype(int)
        
        # Initialize new columns for Book2
        df2['Booking 1 Id'] = None
        df2['Booking 1 Persons'] = None
        df2['Booking 2 Id'] = None
        df2['Booking 2 Persons'] = None
        
        # Ensure columns that will receive string data are of object type to avoid FutureWarnings
        cols_to_convert = ['Unique Id', 'Group Admin Name', 'Age', 'WhatsApp No', 'BOOKING']
        for col in cols_to_convert:
            if col in df2.columns:
                df2[col] = df2[col].astype('object')
            else:
                df2[col] = None # Create if doesn't exist
                df2[col] = df2[col].astype('object')

        total_rows = len(df2)
        
        for idx, row in df2.iterrows():
            # Update progress (approximate)
            progress = (current_sheet_idx / total_sheets) + ((idx + 1) / total_rows / total_sheets)
            progress_bar.progress(min(progress, 1.0))
            
            target = row['test']
            if target <= 0:
                continue
                
            # Get unassigned bookings
            unassigned = [b for b in bookings if b['Allocation Status'] == "Not Allocated"]
            
            # Find a combination (Max 2)
            match = find_combination(target, unassigned)
            
            if match:
                # Prepare strings for Book2 columns
                unique_ids = []
                names = []
                ages = []
                whatsapps = []
                
                for i, booking in enumerate(match):
                    # Mark as assigned in our list
                    mark_assigned(booking['original_index'], row['New Dhaja No.'])
                    
                    # Collect details
                    unique_ids.append(str(booking['Unique Id']))
                    names.append(str(booking['Group Admin Name']))
                    ages.append(str(booking['Age']))
                    whatsapps.append(str(booking['WhatsApp No']))
                    
                    # Fill specific columns
                    col_prefix = f"Booking {i+1}"
                    df2.at[idx, f'{col_prefix} Id'] = booking['Unique Id']
                    df2.at[idx, f'{col_prefix} Persons'] = booking['No. of Person']
                
                # Update Book2 DataFrame
                # We use 'at' to modify the specific cell
                df2.at[idx, 'Unique Id'] = ", ".join(unique_ids)
                df2.at[idx, 'Group Admin Name'] = ", ".join(names)
                df2.at[idx, 'Age'] = ", ".join(ages)
                df2.at[idx, 'WhatsApp No'] = ", ".join(whatsapps)
                df2.at[idx, 'BOOKING'] = "Allocated"
        
        processed_sheets[sheet_name] = df2
        current_sheet_idx += 1
            
    # Update df1 from the modified bookings list
    status_map = {b['original_index']: b['Allocation Status'] for b in bookings}
    dhaja_map = {b['original_index']: b['Allotted Dhaja No'] for b in bookings}
    
    df1['Allocation Status'] = df1['original_index'].map(status_map)
    df1['Allotted Dhaja No'] = df1['original_index'].map(dhaja_map)
    
    # Drop the helper column
    df1.drop(columns=['original_index'], inplace=True)
            
    return df1, processed_sheets

st.set_page_config(page_title="Dhaja Allocation Tool", layout="wide")

st.title("Dhaja Allocation Tool")
st.markdown("""
This tool allocates bookings from **Book1** to allotments in **Book2**.
- Matches bookings to the 'test' column in Book2.
- Combines at most **2 bookings** to match the target number.
""")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Upload Book1 (Bookings)")
    uploaded_file1 = st.file_uploader("Choose Book1 Excel file", type="xlsx", key="file1")

with col2:
    st.subheader("Upload Book2 (Allotments)")
    uploaded_file2 = st.file_uploader("Choose Book2 Excel file", type="xlsx", key="file2")

if uploaded_file1 and uploaded_file2:
    if st.button("Run Allocation"):
        with st.spinner("Allocating..."):
            df1_result, processed_sheets = process_allocation(uploaded_file1, uploaded_file2)
            
            if df1_result is not None and processed_sheets is not None:
                st.success("Allocation Complete!")
                
                # Preview first sheet
                first_sheet_name = list(processed_sheets.keys())[0]
                st.subheader(f"Preview of Allotments (Sheet: {first_sheet_name})")
                st.dataframe(processed_sheets[first_sheet_name].head())
                
                st.subheader("Preview of Bookings (Book1)")
                st.dataframe(df1_result.head())
                
                # Convert to Excel for download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Write all processed Book2 sheets
                    for sheet_name, df_sheet in processed_sheets.items():
                        df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # Write Book1
                    df1_result.to_excel(writer, sheet_name='Bookings', index=False)
                    
                    # Add auto-filter to all sheets
                    workbook = writer.book
                    for sheet_name in writer.sheets:
                        worksheet = writer.sheets[sheet_name]
                        worksheet.auto_filter.ref = worksheet.dimensions
                
                st.download_button(
                    label="Download Final Excel Output",
                    data=output.getvalue(),
                    file_name="Allocation_Results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
