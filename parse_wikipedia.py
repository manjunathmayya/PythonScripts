'''
1. Extract information regarding Indian states from wikipedia,
using BeautifulSoup.
2. Write the extracted data into a excel file

'''

import requests
from bs4 import BeautifulSoup
import pandas as pd


page = requests.get("https://en.wikipedia.org/wiki/States_and_union_territories_of_India")
soup = BeautifulSoup(page.text,'lxml')


mytable = soup.find("table", {"class": "wikitable sortable plainrowheaders"})

df = pd.DataFrame()

states = []
population = []
area = []
vehicle_code = []
capital = []


states_from_table = mytable.findAll("th", {"scope": "row"})
for state in states_from_table:
    states.append(state.text.strip())


body = mytable.find("tbody")
rows = body.findAll("tr")


for row in rows:
    td = row.findAll("td")
    length = len(td)
    if length > 0:
        if length == 9:
            population.append(td[5].text.strip())
            area.append(td[6].text.strip())
        else:
            population.append(td[6].text.strip())
            area.append(td[7].text.strip())
            
        vehicle_code.append(td[1].text.strip())
        
        #capital
        capitals = td[3].findAll("a")
        cap_string = ''
        for cap in capitals:
            title = cap.get('title')
            if title is not None:
                cap_string = cap_string + cap.get('title') + '; '
        capital.append(cap_string)  


df['States'] = states
df['Population'] = population
df['Area'] = area
df['vehicle_code'] = vehicle_code
df['capital'] = capital

df.to_excel('states.xlsx')
