import openmeteo_requests
import requests_cache
from retry_requests import retry
import requests
import pandas as pd
import openpyxl as xl 
import numpy as np
months = ['01','02','03','04','05','06','07','08','09','10','11','12']
def get_data(country_name, year):
    GEO_URL = "https://geocoding-api.open-meteo.com/v1/search"
    weather_url = "https://archive-api.open-meteo.com/v1/archive"
    temp_list = []
    rh_list = []
    
    # Parameters for the API request
    params = {
        "name": country_name,
        "count": 1,         # Request only the top result
        "language": "en",   # Optional: specify language
        "format": "json"    # Optional: specify format
    }
    
    try:
        response = requests.get(GEO_URL, params=params)
        response.raise_for_status() # Raise an exception for bad status codes (4xx or 5xx)
        
        data = response.json()
        
        # Check if any results were found
        if "results" in data and data["results"]:
            result = data["results"][0]
            latitude = result.get("latitude")
            longitude = result.get("longitude")

            for i in range(1,13):
                params = {
                        "latitude": latitude,
                        "longitude": longitude,
                        "hourly": ["temperature_2m", "relative_humidity_2m"],
                        "start_date": f'{year}-{months[i-1]}-01',
                        "end_date": f'{year}-{months[i-1]}-01'
                        }
                
                cache_session = requests_cache.CachedSession('.cache', expire_after = 3600)
                retry_session = retry(cache_session, retries = 5, backoff_factor = 0.2)
                openmeteo = openmeteo_requests.Client(session = retry_session)
                
                responses = openmeteo.weather_api(weather_url, params=params)
                response = responses[0]
                hourly = response.Hourly()
                temperature = hourly.Variables(0).ValuesAsNumpy()
                humidity = hourly.Variables(1).ValuesAsNumpy()
                temp_list.append(temperature)
                rh_list.append(humidity)
            
            temperature = np.concatenate(temp_list)
            humidity = np.concatenate(rh_list)
            return temperature, humidity

        else:
            print(f"No results found for '{country_name}'.")
            return None, None
            
    except requests.exceptions.RequestException as e:
        print(f"An error occurred during the API request: {e}")
        return None, None

df = pd.read_excel('data.xlsx', sheet_name='country')
country_name = df.values.flatten()
df = pd.read_excel('data.xlsx', sheet_name='env_condition')
names = df.iloc[:,0].values

years = [2011, 2013, 2014, 2015, 2017, 2018, 2019, 2021, 2022, 2023]
for year in years:
    temperatures = []
    humidities = []
    for country in country_name:
        temperature, humidity = get_data(country, year)   
        temp_cols = {f"temperature_{i+1}": v for i, v in enumerate(temperature)}
        hum_cols  = {f"humidity_{i+1}": v for i, v in enumerate(humidity)}
        name = names[np.where(country_name == country)][0]
        temperature_data = {
            "Country": name,
            **temp_cols
        }

        humidity_data = {
            "Country": name,
            **hum_cols
        }
        temperatures.append(temperature_data)
        humidities.append(humidity_data)

    temperature_sheet = "temperature_data_" + str(year)
    humidity_sheet = "humidity_data_" + str(year)
    writer = pd.ExcelWriter('weather_data.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace')
    data = pd.DataFrame(temperatures)
    data.to_excel(writer, sheet_name=temperature_sheet, index=False) 
    data = pd.DataFrame(humidities)
    data.to_excel(writer, sheet_name=humidity_sheet, index=False)   
    writer.close()