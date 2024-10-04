import openpyxl
import requests
import os
from dotenv import load_dotenv

load_dotenv()


def search_address(adress):
    url = "https://mapfanapi-search.p.rapidapi.com/addr"

    querystring = {
        "addr": adress,
        "gov": "0",
        "fmt": "json",
    }

    headers = {
        "x-rapidapi-key": os.getenv("API"),
        "x-rapidapi-host": "mapfanapi-search.p.rapidapi.com",
    }

    response = requests.get(url, headers=headers, params=querystring)

    dic = response.json()

    print(dic)

    lon = dic["results"][0]["lon"]
    lat = dic["results"][0]["lat"]

    lon_lat = f"{lon},{lat}"

    return lon_lat


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

    return distance


wb = openpyxl.load_workbook("C:/Users/pytho/OneDrive/Desktop/20241004/route_search.xlsx")
ws = wb["Sheet1"]

count = 0
row_list = []
for row in ws.iter_rows(min_row=2, max_col=5):
    # 範囲の最小列が２、範囲の最大列が５
    if row[0].value is None:
        continue
    row_list.append([cell.value for cell in row])

for row in row_list:
    staff_address = row[2]
    office_address = row[3]

    staff_lon_lat = search_address(staff_address)
    office_lon_lat = search_address(office_address)

    distance = search_route(staff_lon_lat, office_lon_lat)

    d = round(distance / 1000, 1)

    row[4] = d

for idx, row in enumerate(row_list):
    ws.cell(idx + 2, 5).value = row[4]

wb.save("route_search.xlsx")


# 誰から見てもわかりにくい変数設定になってしまった
