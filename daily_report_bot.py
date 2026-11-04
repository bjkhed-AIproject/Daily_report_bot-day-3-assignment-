import pyautogui
import pyperclip
import requests
import subprocess
import time
import os

from datetime import datetime
from pathlib import Path


# --------------------------------------------------
# 1. BASIC SETTINGS
# --------------------------------------------------

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

city = "Chennai"

# Chennai location coordinates
latitude = 13.0827
longitude = 80.2707

project_folder = Path(__file__).parent
output_folder = project_folder / "output"
output_folder.mkdir(exist_ok=True)

now = datetime.now()
date_time = now.strftime("%Y-%m-%d %H:%M:%S")
today = now.strftime("%Y-%m-%d")

excel_path = output_folder / f"daily_report_{today}.xlsx"
screenshot_path = output_folder / f"daily_report_{today}.png"


# --------------------------------------------------
# 2. GET WEATHER DATA
# --------------------------------------------------

print("Getting current weather information...")

weather_url = (
    "https://api.open-meteo.com/v1/forecast"
    f"?latitude={latitude}"
    f"&longitude={longitude}"
    "&current=temperature_2m,relative_humidity_2m,"
    "apparent_temperature,precipitation,weather_code,"
    "wind_speed_10m"
    "&timezone=Asia%2FKolkata"
)

response = requests.get(weather_url, timeout=30)
response.raise_for_status()

weather_json = response.json()

if "current" not in weather_json:
    raise ValueError("Weather information was not returned by the service.")

current = weather_json["current"]

temperature = current["temperature_2m"]
humidity = current["relative_humidity_2m"]
feels_like = current["apparent_temperature"]
wind_speed = current["wind_speed_10m"]
rainfall = current["precipitation"]
weather_code = current["weather_code"]


# Convert weather code into readable condition
weather_conditions = {
    0: "Clear sky",
    1: "Mainly clear",
    2: "Partly cloudy",
    3: "Overcast",
    45: "Fog",
    48: "Depositing rime fog",
    51: "Light drizzle",
    53: "Moderate drizzle",
    55: "Dense drizzle",
    61: "Slight rain",
    63: "Moderate rain",
    65: "Heavy rain",
    71: "Slight snow",
    73: "Moderate snow",
    75: "Heavy snow",
    80: "Rain showers",
    81: "Moderate rain showers",
    82: "Heavy rain showers",
    95: "Thunderstorm"
}

condition = weather_conditions.get(
    weather_code,
    "Weather condition code " + str(weather_code)
)

weather_data = (
    f"{city}: {condition}; "
    f"Temperature: {temperature} °C; "
    f"Feels like: {feels_like} °C; "
    f"Humidity: {humidity}%; "
    f"Wind: {wind_speed} km/h; "
    f"Precipitation: {rainfall} mm"
)

print("Weather information received:")
print(weather_data)


# --------------------------------------------------
# 3. PREPARE A SHORT COMMENT
# --------------------------------------------------

if rainfall > 0:
    comment = "Rain recorded. Carry an umbrella."
elif temperature >= 35:
    comment = "Hot conditions. Stay hydrated."
elif "Clear" in condition:
    comment = "Clear weather conditions."
elif "Fog" in condition:
    comment = "Foggy conditions. Drive carefully."
else:
    comment = "Monitor weather conditions during the day."


# --------------------------------------------------
# 4. OPEN CHROME USING PYAUTOGUI
# --------------------------------------------------

print("Opening Google Chrome...")

chrome_locations = [
    Path(os.environ.get("PROGRAMFILES", "C:/Program Files"))
    / "Google/Chrome/Application/chrome.exe",

    Path(os.environ.get("PROGRAMFILES(X86)", "C:/Program Files (x86)"))
    / "Google/Chrome/Application/chrome.exe",

    Path(os.environ.get("LOCALAPPDATA", ""))
    / "Google/Chrome/Application/chrome.exe"
]

chrome = next((path for path in chrome_locations if path.exists()), None)

if chrome is None:
    raise FileNotFoundError("Google Chrome was not found.")

subprocess.Popen([str(chrome), "https://open-meteo.com/"])

time.sleep(8)

# Maximize Chrome
pyautogui.hotkey("win", "up")
time.sleep(2)

# Copy the important weather information to clipboard
pyperclip.copy(weather_data)

print("Weather information copied to clipboard.")


# --------------------------------------------------
# 5. OPEN MICROSOFT EXCEL
# --------------------------------------------------

print("Opening Microsoft Excel...")

subprocess.Popen(
    ["cmd", "/c", "start", "", "excel"],
    shell=False
)

time.sleep(12)

# Create a new workbook
pyautogui.hotkey("ctrl", "n")
time.sleep(5)


# --------------------------------------------------
# 6. ENTER DATA INTO EXCEL
# --------------------------------------------------

print("Entering weather information into Excel...")

excel_data = (
    "Date & Time\tCity\tWeather Condition\tTemperature (°C)"
    "\tFeels Like (°C)\tHumidity (%)\tWind Speed (km/h)"
    "\tPrecipitation (mm)\tComment\n"
    f"{date_time}\t{city}\t{condition}\t{temperature}"
    f"\t{feels_like}\t{humidity}\t{wind_speed}"
    f"\t{rainfall}\t{comment}"
)

# Copy Excel rows to clipboard
pyperclip.copy(excel_data)

# Click the first Excel cell and paste
pyautogui.click(300, 250)
pyautogui.hotkey("ctrl", "v")

time.sleep(4)

# Auto-fit columns
pyautogui.hotkey("ctrl", "a")
pyautogui.hotkey("alt", "h")
pyautogui.hotkey("o")
pyautogui.hotkey("i")

time.sleep(3)


# --------------------------------------------------
# 7. SAVE EXCEL FILE
# --------------------------------------------------

print("Saving Excel report...")

pyautogui.press("f12")
time.sleep(4)

# Enter the full file path
pyautogui.hotkey("ctrl", "a")
pyperclip.copy(str(excel_path))
pyautogui.hotkey("ctrl", "v")

pyautogui.press("enter")
time.sleep(6)

# Handle an overwrite confirmation if today's file exists
pyautogui.press("left")
pyautogui.press("enter")
time.sleep(3)

print("Excel report saved:", excel_path)


# --------------------------------------------------
# 8. TAKE SCREENSHOT
# --------------------------------------------------

print("Taking screenshot of the final report...")

time.sleep(3)

pyautogui.screenshot().save(str(screenshot_path))


# --------------------------------------------------
# 9. COMPLETION MESSAGE
# --------------------------------------------------

print("\n----------------------------------")
print("DAILY REPORT AUTOMATION COMPLETED")
print("----------------------------------")

print("City:", city)
print("Weather:", condition)
print("Temperature:", temperature, "°C")
print("Date & Time:", date_time)
print("Excel file:", excel_path)
print("Screenshot:", screenshot_path)