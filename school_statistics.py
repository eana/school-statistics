#!/usr/bin/env python
import glob
import os
import re
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

headers = [
    "School",
    "Ämne",
    "Total",
    "Flickor",
    "Pojkar",
    "Total",
    "Flickor",
    "Pojkar",
    "Total",
    "Flickor",
    "Pojkar",
]

kommun = {
    'stockholm': '0180',
    'huddinge': '0126'
}


def get_all_schools(path, kommun_namn, kommun_cod):
    Path(path).mkdir(parents=True, exist_ok=True)

    df = pd.read_csv(f"./{path}.csv")

    for school in df.iloc:
        value, name = school
        url = f"https://siris.skolverket.se/reports/rwservlet?cmdkey=common&geo=1&report=gr6_betyg_amne&p_flik=G&p_verksamhetsar=2019&p_hmantyp=01&p_hmankod=&p_lankod=01&p_kommunkod={kommun_cod}&p_skolkod={value}"
        r = requests.get(url)

        pattern = 'rep_out.*.xls'
        urlpart = re.search(pattern, r.content.decode('ISO-8859-1')).group()
        url = f'https://siris.skolverket.se/{urlpart}'
        r = requests.get(url)
        open(f'./{path}/{name}.xls', 'wb').write(r.content)


def read_html(filename):
    with open(filename, "r", encoding="ISO-8859-1") as f:
        soup = BeautifulSoup(f, features="html.parser")
        table = soup.find_all('table')[0]
        rows = table.find_all('tr')
        sheetname = os.path.splitext(os.path.basename(filename))[0]
        headers = ['School'] + [x.string for x in rows[9].find_all('td')]
        if len(rows[10].find_all('td')) == 1:
            return

        rows = rows[10:33]

        global dfObj
        for row in rows:
            values = [sheetname] + [x.string for x in row.find_all('td')]
            dfObj = dfObj.append(dict(zip(headers, values)), ignore_index=True)
        global writer
        dfObj.to_excel(writer, sheet_name="Sheet", index=False)


for key in kommun:
    path = f"./schools-{key}"
    kommun_namn = key
    kommun_cod = kommun[key]

    dfObj = pd.DataFrame(columns=headers)
    writer = pd.ExcelWriter(f"{path}.xlsx", engine="xlsxwriter")

    print(f"Kommun: {kommun_namn}")
    print("Get all schools!")
    get_all_schools(path, kommun_namn, kommun_cod)

    print("Create the excel file!\n")
    paths = glob.glob(f"{path}/*.xls")
    paths.sort()

    for path in paths:
        # print(path)
        read_html(path)

    writer.save()
