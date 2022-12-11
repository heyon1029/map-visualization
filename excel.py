import openpyxl
import requests
import numpy as np
import pandas as pd
import seaborn 
import matplotlib.pyplot as plt

# load excel file
file_path = r'C:\Users\82108\Downloads\서울시 학원 교습소정보.xlsx'
wb = openpyxl.load_workbook(file_path)
sheet = wb['ac']
print(sheet['A5'].value)
addresses = []
 
# save values in addresses
for i in range(2,101):
    addresses.append(sheet['E' + str(i)].value)
    
print(addresses[1])

# check if addresses contain values 
for h in range(2,11):
    print(addresses[h])

api_key = 'XXXXXXXXXX'
#req_url = 'https://maps.googleapis.com/maps/api/geocode/json?address=' + addresses[0] + '&key=' + api_key
#res = requests.get(req_url)
#print(req_url)
#print(res)

#print(res.json()['results'][0]['geometry']['location']['lat'])

#lat = res.json()['results'][0]['geometry']['location']['lat']
#lng = res.json()['results'][0]['geometry']['location']['lng']

# make matrix
pt = np.zeros((len(addresses), 2))

for idx in range(len(addresses)):
    req_url = 'https://maps.googleapis.com/maps/api/geocode/json?address=' + addresses[idx] + '&key=' + api_key
    res = requests.get(req_url)
    if res.status_code == 200:
        #data =res.json()['results'][0]['geometry']['location']['lat']
        #print(data)
        pt[idx][0] = res.json()['results'][0]['geometry']['location']['lat']
        pt[idx][1] = res.json()['results'][0]['geometry']['location']['lng']

data = pd.DataFrame(pt, columns=['latitude', 'longitude'])

print(data)

# visualize data
seaborn.kdeplot(data=data, x="latitude", y="longitude", fill=True,)
plt.show()