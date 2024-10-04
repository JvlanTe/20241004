import requests
import os
from dotenv import load_dotenv

load_dotenv()


def search_route(start, destination):
    url = "https://mapfanapi-route.p.rapidapi.com/calcroute"

    querystring = {
        "start": start,
        "destination": destination,
        "priority": "0",
        "tollway": "0",
        "ferry": "0",
        "smartic": "0",
        "etc": "0",
        "tolltarget": "0",
        "cartype": "0",
        "vehicletype": "0",
        "danger": "0",
        "daytime": "0",
        "generalroad": "0",
        "tollroad": "0",
        "regulations": "0",
        "travel": "0",
        "uturnavoid": "0",
        "uturn": "0",
        "resulttype": "0",
        "fmt": "json",
    }

    headers = {
        "x-rapidapi-key": os.getenv("API"),
        "x-rapidapi-host": "mapfanapi-route.p.rapidapi.com",
    }

    response = requests.get(url, headers=headers, params=querystring)

    dic = response.json()

    # print(dic)

    distance = dic["summary"]["totalDistance"]

    summary = f"{distance}"

    return summary


lola1 = "141.1519179,39.7039458"
lola2 = "141.1327454,39.7019277"
print(search_route(lola1, lola2))
