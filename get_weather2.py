################################################################################
# Get Weather 2 -- Version 2.0
# Jeffrey D. Shaffer and Google Gemini
# 2025-06-08
#
################################################################################
# Grabs the local weather data and air quality data, 
# calculates the AQI (based on PM 2.5, PM 10.0, CO, NO2, SO2, and Ozone data),
# and appends the weather data and AQI to an XLSX file.
#
################################################################################
# Requires the following libraries:  requests, pandas, openpyxl
# Install using:  pip install requests pandas openpyxl
# Run using:  python3 get_weather.py
#
# If you are not in Japan, you'll want to go to open-meteo.com and select 
# a different source as this one uses a Japanese source for the weather data.
#
# You can set your own latitude and longitude (find using the URL from Google Maps).
# You can change the output file by editing output filename.
#
# To record the wind direction in degrees, instead of compass directions, 
# comment out the second wind_dir (with the convert_wind_to_compass function)
# and uncomment the first wind_dir (with no convert function).
#
# UNITS USED
#    temperature    = °C
#    feels_like     = °C
#    humidity       = %
#    pressure       = hPa
#    windspeed      = mps (converted from kph)
#    wind_dir       = (compass directions by default, degrees by choice)
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
# The lat and lon are for Shizuoka, Japan
LATITUDE = '34.975'
LONGITUDE = '138.4088016'
WEATHER_API_URL = f'https://api.open-meteo.com/v1/forecast?latitude={LATITUDE}&longitude={LONGITUDE}&current=temperature_2m,relative_humidity_2m,apparent_temperature,is_day,precipitation,weather_code,cloud_cover,surface_pressure,wind_speed_10m,wind_direction_10m&timezone=Asia%2FTokyo&models=jma_seamless'

# --- Configuration for Air Quality Data ---
AQ_API_URL = f'https://air-quality-api.open-meteo.com/v1/air-quality?latitude={LATITUDE}&longitude={LONGITUDE}&current=pm2_5,carbon_monoxide,nitrogen_dioxide,sulphur_dioxide,ozone,pm10'

OUTPUT_FILENAME = 'shizuoka_wx_data.xlsx'

# Function to convert wind direction (degrees) to compass directions
def convert_wind_to_compass(wind_dir):
    if wind_dir >= 0 and wind_dir < 22.5:
        return "N"
    elif wind_dir >= 22.5 and wind_dir < 45:
        return "NNE"
    elif wind_dir >= 45 and wind_dir < 67.5:
        return "NE"
    elif wind_dir >= 67.5 and wind_dir < 90:
        return "ENE"
    elif wind_dir >= 90 and wind_dir < 112.5:
        return "E"
    elif wind_dir >= 112.5 and wind_dir < 135:
        return "ESE"
    elif wind_dir >= 135 and wind_dir < 157.5:
        return "SE"
    elif wind_dir >= 157.5 and wind_dir < 180:
        return "SSE"
    elif wind_dir >= 180 and wind_dir < 202.5:
        return "S"
    elif wind_dir >= 202.5 and wind_dir < 225:
        return "SSW"
    elif wind_dir >= 225 and wind_dir < 247.5:
        return "SW"
    elif wind_dir >= 247.5 and wind_dir < 270:
        return "WSW"
    elif wind_dir >= 270 and wind_dir < 292.5:
        return "W"
    elif wind_dir >= 292.5 and wind_dir < 315:
        return "WNW"
    elif wind_dir >= 315 and wind_dir < 337.5:
        return "NW"
    elif wind_dir >= 337.5 and wind_dir < 360:
        return "NNW"
    elif wind_dir == 360:
        return "N"
    else: # Handle potential unexpected values
        return "Unknown"


# --- AQI Calculation Function ---
def calculate_aqi_from_data(json_data):
    current = json_data.get("current", {})

    pm2_5_ugm3 = current.get("pm2_5")
    carbon_monoxide_ugm3 = current.get("carbon_monoxide")
    nitrogen_dioxide_ugm3 = current.get("nitrogen_dioxide")
    sulphur_dioxide_ugm3 = current.get("sulphur_dioxide")
    ozone_ugm3 = current.get("ozone")
    pm10_ugm3 = current.get("pm10")

    # Conversion Factors (approximate, at 25°C and 1 atm)
    # CO: 1 μg/m³ = 0.873 ppb (1 mg/m³ = 0.873 ppm)
    # NO2: 1 μg/m³ = 0.532 ppb
    # SO2: 1 μg/m³ = 0.375 ppb
    # Ozone: 1 μg/m³ = 0.5 ppb

    # US EPA AQI Breakpoints (Concentration Hi/Lo and AQI Hi/Lo)
    aqi_breakpoints = {
        "pm2_5": [
            (0.0, 12.0, 0, 50), (12.1, 35.4, 51, 100), (35.5, 55.4, 101, 150),
            (55.5, 150.4, 151, 200), (150.5, 250.4, 201, 300),
            (250.5, 350.4, 301, 400), (350.5, 500.4, 401, 500),
        ],
        "pm10": [
            (0, 54, 0, 50), (55, 154, 51, 100), (155, 254, 101, 150),
            (255, 354, 151, 200), (355, 424, 201, 300),
            (425, 504, 301, 400), (505, 604, 401, 500),
        ],
        "carbon_monoxide": [ # in ppm
            (0.0, 4.4, 0, 50), (4.5, 9.4, 51, 100), (9.5, 12.4, 101, 150),
            (12.5, 15.4, 151, 200), (15.5, 30.4, 201, 300),
            (30.5, 40.4, 301, 400), (40.5, 50.4, 401, 500),
        ],
        "nitrogen_dioxide": [ # in ppb (1-hour average, used for AQI only if 1-hour ozone is not available or very high)
            (0, 53, 0, 50), (54, 100, 51, 100), (101, 360, 101, 150),
            (361, 649, 151, 200), (650, 1249, 201, 300),
            (1250, 1649, 301, 400), (1650, 2049, 401, 500),
        ],
        "sulphur_dioxide": [ # in ppb (1-hour average)
            (0, 35, 0, 50), (36, 75, 51, 100), (76, 185, 101, 150),
            (186, 304, 151, 200), (305, 604, 201, 300),
            (605, 804, 301, 400), (805, 1004, 401, 500),
        ],
        "ozone": [ # in ppb (8-hour average for 0-100, 1-hour for higher)
            (0, 54, 0, 50), (55, 70, 51, 100), (71, 85, 101, 150),
            (86, 105, 151, 200), (106, 200, 201, 300),
        ],
    }

    def get_aqi_category(aqi_value):
        if aqi_value >= 301: return "Hazardous"
        elif aqi_value >= 201: return "Very Unhealthy"
        elif aqi_value >= 151: return "Unhealthy"
        elif aqi_value >= 101: return "Unhealthy for Sensitive Groups"
        elif aqi_value >= 51: return "Moderate"
        else: return "Good"

    def calculate_sub_aqi(pollutant_value, pollutant_type):
        if pollutant_value is None:
            return 0 # Handle as missing data

        breakpoints = aqi_breakpoints.get(pollutant_type)
        if not breakpoints:
            return 0 # Unknown pollutant type

        # Find the correct breakpoint range
        for C_Lo, C_Hi, I_Lo, I_Hi in breakpoints:
            if C_Lo <= pollutant_value <= C_Hi:
                if C_Hi == C_Lo:
                    return I_Lo
                aqi_sub = ((I_Hi - I_Lo) / (C_Hi - C_Lo)) * (pollutant_value - C_Lo) + I_Lo
                return round(aqi_sub)
            elif pollutant_value > C_Hi and (C_Hi == breakpoints[-1][1]):
                return 501 # Indicate value is in "Beyond AQI" or "Hazardous"
        return 0

    sub_aqis = []

    if pm2_5_ugm3 is not None:
        sub_aqis.append(calculate_sub_aqi(pm2_5_ugm3, "pm2_5"))
    if pm10_ugm3 is not None:
        sub_aqis.append(calculate_sub_aqi(pm10_ugm3, "pm10"))
    if carbon_monoxide_ugm3 is not None:
        co_ppm = carbon_monoxide_ugm3 * (0.873 / 1000)
        sub_aqis.append(calculate_sub_aqi(co_ppm, "carbon_monoxide"))
    if nitrogen_dioxide_ugm3 is not None:
        no2_ppb = nitrogen_dioxide_ugm3 * 0.532
        sub_aqis.append(calculate_sub_aqi(no2_ppb, "nitrogen_dioxide"))
    if sulphur_dioxide_ugm3 is not None:
        so2_ppb = sulphur_dioxide_ugm3 * 0.375
        sub_aqis.append(calculate_sub_aqi(so2_ppb, "sulphur_dioxide"))
    if ozone_ugm3 is not None:
        o3_ppb = ozone_ugm3 * 0.5
        sub_aqis.append(calculate_sub_aqi(o3_ppb, "ozone"))

    final_aqi = max(sub_aqis) if sub_aqis else 0
    category = get_aqi_category(final_aqi)

    return final_aqi, category


# Function to get both weather and air quality data
def get_combined_data():
    weather_data = None
    aq_data = None

    # Fetch weather data
    try:
        response_weather = requests.get(WEATHER_API_URL)
        response_weather.raise_for_status()
        weather_data = response_weather.json()
        print("Successfully fetched weather data.") # Debug print
    except requests.exceptions.RequestException as e:
        print(f"Error fetching weather data: {e}")

    # Fetch air quality data
    try:
        response_aq = requests.get(AQ_API_URL)
        response_aq.raise_for_status()
        aq_data = response_aq.json()
        print("Successfully fetched air quality data.") # Debug print
    except requests.exceptions.RequestException as e:
        print(f"Error fetching air quality data: {e}")

    if weather_data and aq_data:
        # Extract weather parameters
        current_weather = weather_data['current']
        temp_c = current_weather['temperature_2m']
        feels_like_c = current_weather['apparent_temperature']
        humidity_percent = current_weather['relative_humidity_2m']
        pressure_hpa = current_weather['surface_pressure']
        wind_speed_kph = current_weather['wind_speed_10m']
        wind_direction_deg = current_weather['wind_direction_10m']
        cloud_cover_percent = current_weather['cloud_cover']
        precipitation_mm = current_weather['precipitation']

        # Calculate AQI
        aqi_value, aqi_category = calculate_aqi_from_data(aq_data)

        # Combine all data into a single dictionary
        combined_record = {
            'date': datetime.now().strftime('%Y-%m-%d'),
            'time': datetime.now().strftime('%H:%M:%S'),
            'temp': temp_c,
            'feels_like': feels_like_c,
            'humidity': humidity_percent,
            'pressure': pressure_hpa,
            'wind_speed': float(f"{wind_speed_kph/3.6:.2f}"),                  # Convert kph to mps
#            'wind_dir': data['current']['wind_direction_deg'],                # use this one for degrees
            'wind_dir': convert_wind_to_compass(wind_direction_deg),           # use this one for compass 
            'cloud_cover': cloud_cover_percent,
            'precipitation': precipitation_mm,
            'aqi': aqi_value
        }
        return combined_record
    else:
        return None


# Function to adjust column width and center cell content
def adjust_column_width_and_center(file_name):
    workbook = load_workbook(file_name)
    worksheet = workbook.active

    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter
        
        for cell in col:
            cell.alignment = Alignment(horizontal='center', vertical='center')
            
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        
        adjusted_width = max_length + 2
        worksheet.column_dimensions[column].width = adjusted_width

    workbook.save(file_name)


# Function to create or append to the Excel file
def save_combined_data(combined_data):
    if not os.path.exists(OUTPUT_FILENAME):
        df = pd.DataFrame([combined_data])
        df.to_excel(OUTPUT_FILENAME, index=False)
    else:
        df = pd.read_excel(OUTPUT_FILENAME)
        df = df._append(combined_data, ignore_index=True)
        df.to_excel(OUTPUT_FILENAME, index=False)

    adjust_column_width_and_center(OUTPUT_FILENAME)


# Main function to get combined data and save it
def main():
    combined_data = get_combined_data()
    if combined_data:
        save_combined_data(combined_data)
        print(f"Combined weather and AQI data for {combined_data['date']} {combined_data['time']} saved successfully to {OUTPUT_FILENAME}.")
    else:
        print("Failed to retrieve combined weather and AQI data.")

if __name__ == '__main__':
    main()
