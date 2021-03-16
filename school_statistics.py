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
    "Totalt",
    "Flickor",
    "Pojkar",
    "Betygspoäng - Totallt",
    "Andel (%) med A-E1 - Totallt",
    "Betygspoäng - Flickor",
    "Andel (%) med A-E2 - Flickor",
    "Betygspoäng3 - Pojkar",
    "Andel (%) med A-E3 - Pojkar",
]

subject = [
    "Bild",
    "Biologi",
    "Engelska",
    "Fysik",
    "Geografi",
    "Hem och konsumentkunskap",
    "Historia",
    "Idrott och hälsa",
    "Kemi",
    "Matematik",
    "Moderna språk, språkval",
    "Modersmål",
    "Musik",
    "Religionskunskap",
    "Samhällskunskap",
    "Slöjd",
    "Svenska",
    "Svenska som andraspråk",
    "Teknik"
]

kommun = {
    "stockholm": "0180",
    "huddinge": "0126"
}


def get_all_schools(path, kommun_cod, grade):
    Path(f"{path}-{grade}").mkdir(parents=True, exist_ok=True)

    df = pd.read_csv(f"{path}.csv")

    for school in df.iloc:
        value, name = school

        url = f"https://siris.skolverket.se/reports/rwservlet?cmdkey=common&geo=1&report=gr_betyg_amne&p_flik=G&p_ar=2019&p_lankod=&p_kommunkod={kommun_cod}&p_skolkod={value}&p_hmantyp=05&p_hmankod=5566641691&p_flik=H"

        r = requests.get(url)
        pattern = "rep_out.*.xls"
        urlpart = re.search(
            pattern, r.content.decode("ISO-8859-1")).group()
        url = f"https://siris.skolverket.se/{urlpart}"
        r = requests.get(url)
        open(f"./{path}-{grade}/{name}.xls", "wb").write(r.content)


def read_html(path, grade):
    paths = glob.glob(f"{path}-{grade}/*.xls")
    # paths = ["schools-stockholm-9/Kungliga Svenska Balettskolan.xls"]
    paths.sort()

    dfObj = pd.DataFrame(columns=headers)
    writer = pd.ExcelWriter(f"{path}-{grade}.xlsx", engine="xlsxwriter")

    for filename in paths:
        with open(filename, "r", encoding="ISO-8859-1") as f:
            soup = BeautifulSoup(f, features="html.parser")
            table = soup.find_all("table")[0]
            rows = table.find_all("tr")
            sheetname = os.path.splitext(os.path.basename(filename))[0]

            rows = rows[11:30]

            for row in rows:
                if str(row.find_all("td")[0].text) not in subject:
                    continue
                values = [sheetname] + \
                    [x.string for x in row.find_all("td")]
                dfObj = dfObj.append(
                    dict(zip(headers, values)), ignore_index=True)

    dfObj.to_excel(writer, sheet_name="Sheet", index=False)
    writer.save()


for key in kommun:
    # the grade can be 6 or 9
    grade = 9
    path = f"./schools-{key}"
    kommun_namn = key
    kommun_cod = kommun[key]

    print(f"Kommun: {kommun_namn}")
    print("Get all schools!")
    # get_all_schools(path, kommun_cod, grade)

    print("Create the excel file!\n")
    read_html(path, grade)
