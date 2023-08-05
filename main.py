import requests
import json
import win32com.client as wincom
# Make sure to download pywin32 by using the command  'pip install pywin32' before executing this line
# it will also work if pywinauto library is enabled

while True:
    city = input("Enter the name of the city: ")
    if city == 'q':
        break
    deg = input("Your preferred unit of measurement(Please enter Celsius of Fahrenheit): ")
    speak = wincom.Dispatch("SAPI.SpVoice")

    url = f"http://api.weatherapi.com/v1/current.json?key=YOUR_API_KEY&q={city}"
    # Make sure to replace 'YOUR_API_KEY' with your own unique API key from the website https://www.weatherapi.com/.
    # For more details, consider reading the README File

    r = requests.get(url)
    # print(type(r.text))  # This line prints the type of r.text
    w_dic = json.loads(r.text)

    if deg == "Celsius":
        temp = w_dic["current"]["temp_c"]
    elif deg == 'Fahrenheit':
        temp = w_dic["current"]["temp_f"]

    text = f'The weather of {city} is {temp} degree {deg}'
    print(text)
    speak.Speak(text)
