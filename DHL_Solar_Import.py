import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import calendar
import io

# Complete meter mapping dictionary based on your example
METER_MAPPING = {
    "JAFZA 1-Meter 1": "DAN1/TY/PHLS/ELEC/LV/01-MDB_Energy Meter",
    "JAFZA 1-Meter 2": "DAN1/TY/PHLS/ELEC/LV/02-MDB_Energy Meter",
    "JAFZA 3-Meter 1": "DAN3/ELEC/MDB/01-MDB_Energy Meter",
    "JAFZA 3-Meter 2": "DAN3/ELEC/MDB/02-MDB_Energy Meter",
    "JAFZA 4": "DAN4/MEP/ELEC/LV /01-LVP_Energy Meter",
    "DAFZA 2": "DAN/DAFZA/MEP/ELE/LV/01-LVP_Energy Meter",
    "AFR": "DAN/DWC-AFR/MEP/LVR/LVP-02-MDB_Energy Meter",
    "CGF": "DAN/DWC-CGF/MEP/LVR/LVP-01-MDB_Energy Meter",
    "JAFZA 2-Meter 2": "DAN2/MEP/HVAC/LV Panel/02 LVP_Energy Meter",
    "JAFZA 2-Meter 1": "DAN2/MEP/HVAC/LV Panel/01 LVP_Energy Meter"
}

def process_solar_data(input_file, month, year):
    # Read the Excel file with header starting at row 3 (0-indexed)
    # Skip the first column if it's empty/unwanted
    df = pd.read_excel(input_file, header=2, usecols=lambda x: 'Unnamed' not in x)
    
    # Convert date column to datetime
    df['Date'] = pd.to_datetime(df['Date'], dayfirst=True)
    
    # Filter data for the selected month/year and previous month
    selected_month_data = df[(df['Date'].dt.month == month) & (df['Date'].dt.year == year)].sort_values('Date')
    
    # Get previous month and year
    if month == 1:
        prev_month = 12
        prev_year = year - 1
    else:
        prev_month = month - 1
        prev_year = year
    
    # Get last day of previous month
    last_day_prev_month = calendar.monthrange(prev_year, prev_month)[1]
    prev_month_last_day = datetime(prev_year, prev_month, last_day_prev_month)
    
    # Get data for last day of previous month
    prev_month_data = df[df['Date'] == prev_month_last_day]
    
    # Create output dataframe
    output_data = []
    
    # Get meter names (all columns except Date)
    meter_names = [col for col in df.columns if col != 'Date']
    
    # Process each day in the selected month
    for i, day in enumerate(selected_month_data['Date']):
        # Get current day data
        current_day_data = selected_month_data[selected_month_data['Date'] == day]
        
        # For each meter, calculate the delta
        for meter in meter_names:
            try:
                current_val = float(current_day_data[meter].values[0])
                
                # For the first day of the month, use last day of previous month
                if day.day == 1:
                    prev_day_data = prev_month_data
                else:
                    # Get previous day data (within same month)
                    prev_day = day - timedelta(days=1)
                    prev_day_data = selected_month_data[selected_month_data['Date'] == prev_day]
                
                prev_val = float(prev_day_data[meter].values[0])
                delta = current_val - prev_val
                
                # Format date as "DD-MM-YYYY 23:59:00"
                formatted_date = day.strftime("%d-%m-%Y") + " 23:59:00"
                
                # Get the actual meter name from mapping
                actual_meter = METER_MAPPING.get(meter, meter)
                
                output_data.append({
                    'Time': formatted_date,
                    'Reference Meter': meter,
                    'Solar Energy': delta,
                    'Meter': actual_meter
                })
            except (IndexError, ValueError) as e:
                st.warning(f"Missing or invalid data for {meter} on {day.strftime('%d-%m-%Y')}")
                continue
    
    output_df = pd.DataFrame(output_data)
    return output_df

def main():
    st.title("Solar Data Import Generator")
    
    # File upload
    uploaded_file = st.file_uploader("Upload Solar Sheet Excel File", type=['xlsx', 'xls'])
    
    # Month and year input
    col1, col2 = st.columns(2)
    with col1:
        month = st.selectbox("Select Month", range(1, 13), format_func=lambda x: datetime(1900, x, 1).strftime('%B'))
    with col2:
        current_year = datetime.now().year
        year = st.selectbox("Select Year", range(current_year - 5, current_year + 5), index=5)
    
    if uploaded_file is not None:
        try:
            # Process the data
            result_df = process_solar_data(uploaded_file, month, year)
            
            # Show preview
            st.subheader("Preview of Solar Data Import Sheet")
            st.dataframe(result_df.head(20))  # Show more rows to see multiple days
            
            # Show meter mapping information
            st.subheader("Current Meter Mapping")
            st.write(pd.DataFrame.from_dict(METER_MAPPING, orient='index', columns=['Actual Meter']))
            
            # Create a BytesIO buffer for the Excel file
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                result_df.to_excel(writer, index=False)
            
            # Download button
            st.download_button(
                label="Download Solar Data Import Sheet",
                data=output.getvalue(),
                file_name=f"Solar_Data_Import_Sheet_{month}_{year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
