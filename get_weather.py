################################################################################
# Python3 script to grab my local weather and append it to an XLS file
# Initial program created by ChatGPT, then heavily edited and modified by Jds
# 2024-10-16
#
# Requires the following libraries:  requests, pandas, openpyxl
# Install using:  pip install requests pandas openpyxl
# Run using:  python3 get_weather.py
#
# If you are not in Japan, you'll want to go to open-meteo.com and select 
# a different source as this one uses a Japanese source for the weather data.
#
# You can set your own latitude and longitude (find using the URL from Google Maps)
# You can change the output file by editing output filename
#
# UNITS USED
#    temperature    = °C
#    feels_like     = °C
#    humidity       = %
#    pressure       = hPa
#    windspeed      = mps (converted from kph)
#    wind_dir       = °
#    cloud_cover    = %
#    precipitation  = mm
#
################################################################################

import requests
import pandas as pd
from datetime import datetime
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Currently using the OpenMeteo JMA API
# https://open-meteo.com/en/docs/jma-api
# The lat and lon are for Shizuoka Station, Shizuoka, Japan
LATITUDE = '34.971'
LONGITUDE = '138.378599'
URL = f'https://api.open-meteo.com/v1/forecast?latitude={LATITUDE}&longitude={LONGITUDE}&current=temperature_2m,relative_humidity_2m,apparent_temperature,is_day,precipitation,weather_code,cloud_cover,surface_pressure,wind_speed_10m,wind_direction_10m&timezone=Asia%2FTokyo&models=jma_seamless'


# Function to get the weather data
def get_weather_data():
    response = requests.get(URL)
    if response.status_code == 200:
        data = response.json()
        weather = {
            'date': datetime.now().strftime('%Y-%m-%d'),
            'time': datetime.now().strftime('%H:%M:%S'),
            'temp': data['current']['temperature_2m'],
            'feels_like': data['current']['apparent_temperature'],
            'humidity': data['current']['relative_humidity_2m'],
            'pressure': data['current']['surface_pressure'],
            'wind_speed': float(f"{data['current']['wind_speed_10m']/3.6:.2f}"),
            'wind_dir': data['current']['wind_direction_10m'],
            'cloud_cover': data['current']['cloud_cover'],
            'precipitation': data['current']['precipitation']
        }
        return weather
    else:
        print('Error fetching weather data')
        return None


# Function to adjust column width and center cell content
def adjust_column_width_and_center(file_name):
    workbook = load_workbook(file_name)
    worksheet = workbook.active

    # Iterate over all rows and columns to apply alignment and adjust column width
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column letter (e.g., A, B, C, ...)
        
        for cell in col:
            # Apply center alignment to each cell
            cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # Calculate maximum length of data in the column
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        
        # Adjust column width based on the longest value
        adjusted_width = max_length + 2  # Add some padding for readability
        worksheet.column_dimensions[column].width = adjusted_width

    workbook.save(file_name)


# Function to create or append to the Excel file
def save_weather_data(weather_data):
    file_name = 'shizuoka_wx_data.xlsx'

    # Check if file exists, if not create a new file
    if not os.path.exists(file_name):
        df = pd.DataFrame([weather_data])
        df.to_excel(file_name, index=False)
    else:
        # If file exists, append the new data
        df = pd.read_excel(file_name)
        df = df._append(weather_data, ignore_index=True)
        df.to_excel(file_name, index=False)

    # Adjust the column width and center the content
    adjust_column_width_and_center(file_name)


# Main function to get weather data and save it
def main():
    weather_data = get_weather_data()
    if weather_data:
        save_weather_data(weather_data)
        print(f"Weather data for {weather_data['date']} saved successfully.")
    else:
        print("Failed to retrieve weather data.")

if __name__ == '__main__':
    main()
