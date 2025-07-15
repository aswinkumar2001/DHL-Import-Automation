import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import calendar
import openpyxl

def process_solar_data(input_file, month, year):
    # Read the Excel file with header starting at row 3 (0-indexed)
    df = pd.read_excel(input_file, header=2)
    
    # Convert date column to datetime
    df['Date'] = pd.to_datetime(df['Date'], dayfirst=True)
    
    # Filter data for the selected month/year and previous month
    selected_month_data = df[(df['Date'].dt.month == month) & (df['Date'].dt.year == year)]
    
    # Get previous month and year
    if month == 1:
        prev_month = 12
        prev_year = year - 1
    else:
        prev_month = month - 1
        prev_year = year
    
    # Get last day of previous month
    last_day_prev_month = calendar.monthrange(prev_year, prev_month)[1]
    
    # Get data for last day of previous month
    prev_month_data = df[(df['Date'].dt.month == prev_month) & 
                         (df['Date'].dt.year == prev_year) &
                         (df['Date'].dt.day == last_day_prev_month)]
    
    # Create output dataframe
    output_data = []
    
    # Get meter names (all columns except Date)
    meter_names = [col for col in df.columns if col != 'Date']
    
    # Process each day in the selected month
    for day in selected_month_data['Date']:
        # Get current day data
        current_day_data = selected_month_data[selected_month_data['Date'] == day]
        
        # Get previous day data
        prev_day = day - timedelta(days=1)
        prev_day_data = df[df['Date'] == prev_day]
        
        # If previous day data doesn't exist (first day of month), use last day of previous month
        if prev_day_data.empty:
            prev_day_data = prev_month_data
        
        # Calculate delta for each meter
        for meter in meter_names:
            try:
                current_val = float(current_day_data[meter].values[0])
                prev_val = float(prev_day_data[meter].values[0])
                delta = current_val - prev_val
                
                # Format date as "DD-MM-YYYY 23:59:00"
                formatted_date = day.strftime("%d-%m-%Y") + " 23:59:00"
                
                output_data.append({
                    'Time': formatted_date,
                    'Reference Meter': f"{month}/{year}",
                    'Solar Energy': delta,
                    'Meter': ''
                })
            except (IndexError, ValueError):
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
            st.dataframe(result_df.head())
            
            # Download button
            output = result_df.to_excel(index=False)
            st.download_button(
                label="Download Solar Data Import Sheet",
                data=output,
                file_name=f"Solar_Data_Import_Sheet_{month}_{year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
