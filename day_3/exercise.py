import os
import json
import requests
import openpyxl
from openpyxl.workbook.workbook import Workbook

wb = openpyxl.load_workbook('day_3/output.xlsx')

sheet_one = wb['Sheet1']

sheet_one['A1'] = "Name"
sheet_one['B1'] = "Species"
sheet_one['C1'] = "Gender"
sheet_one['D1'] = "Location"

response = requests.get("https://rickandmortyapi.com/api/character")
clean_data = json.loads(response.text)
result = clean_data['results']

# print(working_sheet)


# using request package, we can make a API  call to retrieve JSON
#  and storing it into a variable here called 'response'
# verify the sttus as 200
# print(response.text)
# load as a python json object and store into a variable


# print(clean_text)

# print(result['name'])
# go through the results

# for each row in the excel spreadsheet
# get name, species, gender, and location name
counter = 2

for i in result:
  print("Name: " + i['name'])
  sheet_one['A' + str(counter)] = i['name']
  print("Species: " + i['species'])
  sheet_one['B' + str(counter)] = i['species']
  print("Gender: " + i['gender'])
  sheet_one['C' + str(counter)] = i['gender']
  print("Location: " + i['location']['name'])
  sheet_one['D' + str(counter)] = i['location']['name']
  counter += 1

# wb.save('day_3/output.xlsx')