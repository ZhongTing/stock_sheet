#!/usr/bin/env python

import urllib.request
import json
import datetime
import openpyxl as excel


# noinspection SpellCheckingInspection
def fetch(stock):
    url_pattern = "http://invest.wessiorfinance.com/Stock_api/Notation_cal?Stock={}&Odate={" \
                  "}&Period=3.5&is_log=0&is_adjclose=0 "
    url = url_pattern.format(stock, datetime.datetime.now().date())
    request = urllib.request.Request(url)
    request.add_header("Referer", "http://invest.wessiorfinance.com/notation.html")
    request.add_header("User-Agent",
                       "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_3) AppleWebKit/602.4.8 (KHTML, like Gecko) "
                       "Version/10.0.3 Safari/602.4.8")
    response = urllib.request.urlopen(request).read().decode("ascii")
    return json.loads(response)


def parse(json_data):
    data = json_data[[key for key in json_data][-1]]
    tl = data["TL"]
    std = data["STD"]
    price = data["theClose"]
    return {
        "position": round((price - tl) / std, 2),
        "std": std,
        "price": price,
        "tl": tl
    }


def valid(worksheet, cell):
    if cell.value is None:
        return False
    value = worksheet["F{}".format(cell.row)].value
    if value is None:
        return True
    return value.date() != datetime.datetime.now().date()


def main(filename="target.xlsx"):
    print("start processing...")
    workbook = excel.load_workbook(filename)
    for worksheet in workbook:
        for cell in worksheet["A"]:
            if valid(worksheet, cell):
                print(cell.value)
                response = fetch(cell.value)
                data = parse(response)
                worksheet["B{}".format(cell.row)].value = data["price"]
                worksheet["C{}".format(cell.row)].value = data["position"]
                worksheet["D{}".format(cell.row)].value = data["tl"]
                worksheet["E{}".format(cell.row)].value = data["std"]
                worksheet["F{}".format(cell.row)].value = datetime.datetime.now().date()
                workbook.save(filename)
    print("process end")


main()
