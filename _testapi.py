#Need to install requests package for python
#easy_install requests
import requests
import os

user = os.getenv('sn_user')
pwd = os.getenv('sn_pwd_dev')

print user
print pwd

# Set the request parameters
url = 'https://redbrickhealthdev.service-now.com/api/now/table/sn_customerservice_case/37dc811f6fce1f00a1af77f16a3ee4f3?sysparm_display_value=true'

# Eg. User name="admin", Password="admin" for this code sample.


# Set proper headers
headers = {"Content-Type":"application/json","Accept":"application/json"}

# Do the HTTP request
response = requests.put(url, auth=(user, pwd), headers=headers ,data="{\"state\":\"New\"}")

print response
# Check for HTTP codes other than 200
if response.status_code != 200: 
    print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
    exit()

# Decode the JSON response into a dictionary and use the data
data = response.json()
print(data)

for item in data:
    #print item.get("assigned_to").get("display_value")
    print item 