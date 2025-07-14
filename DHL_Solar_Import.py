import streamlit as st
import pandas as pd
from datetime import datetime
import io

def process_solar_data(solar_sheet_df, template_df, month, year, meter_mapping):
    try:
        # Convert date column to datetime (MM/DD/YYYY format)
        date_col = next(col for col in solar_sheet_df.columns if 'date' in col.lower().replace(' ', ''))
        solar_sheet_df[date_col] = pd.to_datetime(solar_sheet_df[date_col], format='%m/%d/%Y', errors='coerce')
        
        # Drop rows with invalid dates
        solar_sheet_df = solar_sheet_df.dropna(subset=[date_col])
        
        # Get the last day of previous month to calculate Day 1 delta
        first_day = datetime(year, month, 1)
        last_day_prev_month = first_day - pd.Timedelta(days=1)
        
        # Get data for selected month plus previous month's last day
        month_data = solar_sheet_df[
            ((solar_sheet_df[date_col].dt.month == month) & 
            (solar_sheet_df[date_col].dt.year == year)) |
            (solar_sheet_df[date_col] == last_day_prev_month)
        ].copy()
        
        # Sort by date
        month_data = month_data.sort_values(date_col)
        
        # Get all meter columns (assuming they follow the pattern "JAFZA X-Meter Y")
        meter_columns = [col for col in solar_sheet_df.columns if 'JAFZA' in col.upper() or 'Jafza' in col]
        
        # Calculate deltas for each meter
        deltas = {}
        for meter in meter_columns:
            # Calculate daily differences
            month_data[f'{meter}_delta'] = month_data[meter].diff()
            # Store deltas with date as key (only for the selected month)
            for date, delta in zip(month_data[date_col], month_data[f'{meter}_delta']):
                if date >= first_day and not pd.isna(delta):
                    deltas[(date, meter)] = delta
        
        # Process the template - convert time to datetime
        template_df['Time'] = pd.to_datetime(template_df['Time'], format='%d-%m-%Y %H:%M:%S')
        
        # Update template with deltas
        for idx, row in template_df.iterrows():
            date = row['Time'].date()
            ref_meter = row['Reference meter']
            
            # Find the corresponding meter in solar sheet
            solar_meter = None
            for k, v in meter_mapping.items():
                if v == ref_meter:
                    solar_meter = k
                    break
            
            if solar_meter:
                try:
                    delta = deltas[(pd.to_datetime(date), solar_meter)]
                    template_df.at[idx, 'Solar Energy Meter Reading'] = delta
                except KeyError:
                    continue
        
        return template_df
    
    except Exception as e:
        st.error(f"Error processing data: {str(e)}")
        return None

def main():
    st.title("Solar Energy Data Processor")
    
    # File uploaders
    st.header("Upload Files")
    solar_sheet = st.file_uploader("Upload Solar Sheet", type=['xlsx', 'xls'])
    dhl_template = st.file_uploader("Upload DHL Solar Import Template", type=['xlsx', 'xls'])
    
    # Month and year selection
    st.header("Select Month and Year")
    col1, col2 = st.columns(2)
    with col1:
        month = st.selectbox("Month", range(1, 13), format_func=lambda x: datetime(1900, x, 1).strftime('%B'))
    with col2:
        current_year = datetime.now().year
        year = st.selectbox("Year", range(current_year - 5, current_year + 5), index=5)
    
    # Meter mapping input
    st.header("Meter Name Mapping")
    st.info("Please enter the mapping between Solar Sheet meter names and DHL Template reference meters")
    
    # Default meter mapping - can be edited by user
    meter_mapping = {
        'JAFZA 1-Meter 1': 'Jafza 1-Meter 1',
        'JAFZA 1-Meter 2': 'Jafza 1-Meter 2',
        'JAFZA 2-Meter 1': 'Jafza 2-Meter 1',
        'JAFZA 2-Meter 2': 'Jafza 2-Meter 2',
        'JAFZA 3-Meter 1': 'Jafza 3-Meter 1',
        'JAFZA 3-Meter 2': 'Jafza 3-Meter 2',
        'JAFZA 4': 'Jafza 4-Meter 1',
        'DAFZA 2': 'DAFZA 2-Meter 1',
        'CGF': 'DWC-CGF-Meter 1',
        'AFR': 'DWC-AFR-Meter 2',
    }
    
    # Display editable meter mapping
    mapping_expander = st.expander("Edit Meter Mapping")
    with mapping_expander:
        num_mappings = st.number_input("Number of meter mappings", min_value=1, value=len(meter_mapping))
        
        updated_mapping = {}
        for i in range(num_mappings):
            col1, col2 = st.columns(2)
            with col1:
                solar_name = st.text_input(f"Solar Sheet Meter Name {i+1}", 
                                         value=list(meter_mapping.keys())[i] if i < len(meter_mapping) else "")
            with col2:
                template_name = st.text_input(f"Template Reference Meter {i+1}", 
                                            value=list(meter_mapping.values())[i] if i < len(meter_mapping) else "")
            if solar_name and template_name:
                updated_mapping[solar_name] = template_name
        
        if updated_mapping:
            meter_mapping = updated_mapping
    
    if st.button("Process Data"):
        if solar_sheet and dhl_template:
            try:
                # Read files with different header positions
                solar_df = pd.read_excel(solar_sheet, header=2)  # 3rd row (0-indexed as 2) as header
                template_df = pd.read_excel(dhl_template)  # 1st row as header
                
                # Process data
                result_df = process_solar_data(solar_df, template_df, month, year, meter_mapping)
                
                if result_df is not None:
                    # Prepare Excel file for download
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False)
                    output.seek(0)
                    
                    # Download button
                    st.download_button(
                        label="Download Processed Template",
                        data=output,
                        file_name=f"DHL_Solar_Import_Template_{month:02d}_{year}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("Processing completed successfully!")
                
            except Exception as e:
                st.error(f"An error occurred while reading files: {str(e)}")
        else:
            st.warning("Please upload both files before processing.")

if __name__ == "__main__":
    main()
