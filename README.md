# GrabWeatherHourly

A simple python3 script to grab local weather data and append it to an XLSX file
I have this running with an hourly cron job. Very fun!
Initial program created by ChatGPT, then heavily edited and modified by myself.

Requires the following libraries:  requests, pandas, openpyxl
install using:  pip install requests pandas openpyxl
run using:  python3 get_weather.py

You can set your own latitude and longitude (find using the URL from Google Maps).

You can change the output file by editing output filename.

UNITS USED
   temperature    = °C
   feels_like     = °C
   humidity       = %
   pressure       = hPa
   windspeed      = mps (converted from kph)
   wind_dir       = °
   cloud_cover    = %
   precipitation  = mm
