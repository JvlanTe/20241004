import requests
import os
from dotenv import load_dotenv

load_dotenv()

ERROR_MESSAGE_NOT_SET_ADDRESS = "職員住所または勤務先が未設定"

ERROR_FILE_TITLE = ["職員番号", "氏名", "エラーメッセージ"]


def search_address(adress):
    url = os.environ["SEARCH_URL"]
    key = os.environ["API"]
    host = os.environ["SEARCH_HOST"]

    querystring = {
        "addr": adress,
        "gov": "0",
        "fmt": "json",
    }

    headers = {
        "x-rapidapi-key": f"{key}",
        "x-rapidapi-host": f"{host}",
    }

    response = requests.get(url, headers=headers, params=querystring)

    dic = response.json()

    search_remaining = response.headers["X-RateLimit-Requests-Remaining"]

    lon = dic["results"][0]["lon"]
    lat = dic["results"][0]["lat"]

    lon_lat = f"{lon},{lat}"

    return lon_lat, search_remaining


def search_route(start, destination):
    url = os.environ["ROUTE_URL"]
    key = os.environ["API"]
    host = os.environ["ROUTE_HOST"]

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
        "x-rapidapi-key": f"{key}",
        "x-rapidapi-host": f"{host}",
    }

    response = requests.get(url, headers=headers, params=querystring)

    dic = response.json()

    route_remaining = response.headers["X-RateLimit-Requests-Remaining"]

    # print(dic)

    distance = dic["summary"]["totalDistance"]

    return distance, route_remaining
