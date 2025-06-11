# GetWeather

A simple python3 script to grab local weather data and append it to an XLSX file.
I have this running with an hourly cron job. Very fun!
Initial program created by ChatGPT, then heavily edited and modified by myself.

Requires the following libraries:  `requests, pandas, openpyxl`

install using:  `pip install requests pandas openpyxl
run using:  python3 get_weather.py`

You can set your own latitude and longitude using Google Maps. Find your home, click it, then grab the lattitude and longitude shown in the URL:
(this part of the URL --> @34.9717465,138.378599)

You can change the output file by editing output filename near the bottom of the program.

_Version 2 now grabs air particulate data, calculates the AQI, and writes the particulate data and AQI to the XLSX file._

---

**UNITS USED**

*    temperature    = °C
*    feels_like     = °C
*    humidity       = %
*    pressure       = hPa
*    windspeed      = mps (converted from kph)
*    wind_dir       = °
*    cloud_cover    = %
*    precipitation  = mm
*    pm 2.5         = μg/m3 (micrograms per cubic meter)
*    pm 10          = μg/m3
*    co             = μg/m3
*    no2            = μg/m3
*    so2            = μg/m3
*    o3 (ozone)     = μg/m3
