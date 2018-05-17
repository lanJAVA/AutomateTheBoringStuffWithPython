#! python3
# quickWeather.py - Prings the weather for a location from the command line.

import json, requests, sys

APPID = 'b8d078881d0359d4b2b072e7705eecee'
# Compute location from command line arguments.
if len(sys.argv) < 2:  # Python的命令行参数存放在sys.argv的列表中，第0个参数是Python脚本的文件名
    print('Usage: quickWeather.py location')
    sys.exit()
location = ' '.join(sys.argv[1:])

# Download the JSON data from OpenWeatherMap.org's API.
url = 'https://api.openweathermap.org/data/2.5/weather?q=%s&APPID=%s' % (location, APPID)
response = requests.get(url)
response.raise_for_status()

# Load JSON data into a Python variable.
weatherData = json.loads(response.text)
print(weatherData)
# Print weather descriptions.
print('Current weather in %s:' % location)
print(weatherData['name'], '-', weatherData['main'], '-', weatherData['weather'], '-', weatherData['coord'])
print()