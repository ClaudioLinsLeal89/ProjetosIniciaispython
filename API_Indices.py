##API Tendencias
import requests
import json

url = "https://api.machina.tendencias.com.br/query"
username = ""
password = ""

payload = json.dumps({
  "list": [
    "M91746",
    "M91773"
  ]
})
headers = {
  'Content-Type': 'application/json'
}

response = requests.request("POST", url, headers=headers, auth=(username, password), data=payload)

print(response.text)