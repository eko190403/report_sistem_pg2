import requests

url = 'http://127.0.0.1:5000/process_overtime'
files = {'overtime_file': open('Overtime 7 2026.XLSX', 'rb')}

response = requests.post(url, files=files)

if response.status_code == 200:
    with open('scratch/downloaded_overtime.xlsx', 'wb') as f:
        f.write(response.content)
    print("Success! File saved to scratch/downloaded_overtime.xlsx")
else:
    print("Error:", response.status_code, response.text)
