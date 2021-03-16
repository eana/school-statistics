#!/usr/bin/env python
import glob
import os
import re
from pathlib import Path

import click
import pandas as pd
import requests
from bs4 import BeautifulSoup


@click.command()
@click.option(
    "--grade",
    required=True,
    type=int,
    help="The grade for which you want me to get the statistics",
)
def main(grade):
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

    kommun = {
        "stockholm": "0180",
        "huddinge": "0126"
    }

    def get_all_schools(path, kommun_namn, kommun_cod, grade):
        Path(f"{path}-{grade}").mkdir(parents=True, exist_ok=True)

        df = pd.read_csv(f"{path}.csv")

        for school in df.iloc:
            value, name = school

            if grade == 6:
                # 6th grade
                url = f"https://siris.skolverket.se/reports/rwservlet?cmdkey=common&geo=1&report=gr6_betyg_amne&p_flik=G&p_verksamhetsar=2019&p_hmantyp=01&p_hmankod=&p_lankod=01&p_kommunkod={kommun_cod}&p_skolkod={value}"
            else:
                # 9th grade
                url = f"https://siris.skolverket.se/reports/rwservlet?cmdkey=common&geo=1&report=gr_betyg_amne&p_flik=G&p_ar=2019&p_lankod=01&p_kommunkod={kommun_cod}&p_skolkod=&p_hmantyp=&p_hmankod=&p_flik=G"

            r = requests.get(url)
            pattern = "rep_out.*.xls"
            urlpart = re.search(
                pattern, r.content.decode("ISO-8859-1")).group()
            url = f"https://siris.skolverket.se/{urlpart}"
            r = requests.get(url)
            open(f"./{path}-{grade}/{name}.xls", "wb").write(r.content)

    def read_html(path):
        paths = glob.glob(f"{path}-{grade}/*.xls")
        paths.sort()

        dfObj = pd.DataFrame(columns=headers)
        writer = pd.ExcelWriter(f"{path}.xlsx", engine="xlsxwriter")

        for filename in paths:
            with open(filename, "r", encoding="ISO-8859-1") as f:
                soup = BeautifulSoup(f, features="html.parser")
                table = soup.find_all("table")[0]
                rows = table.find_all("tr")
                sheetname = os.path.splitext(os.path.basename(filename))[0]

                if len(rows[11].find_all("td")) == 1:
                    continue

                rows = rows[11:30]

                for row in rows:
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
        get_all_schools(path, kommun_namn, kommun_cod, grade)

        print("Create the excel file!\n")
        read_html(path)


if __name__ == "__main__":
    main()
