import requests
from bs4 import BeautifulSoup

word = "hint"
url = f"https://www.wordreference.com/enko/{word}"

response = requests.get(url)

if response.status_code == 200:
    html = response.text
    soup = BeautifulSoup(html, 'html.parser')
    title = soup.select_one('#enko\:15553 > td.ToWrd')
    print(title)
    print(title.get_text())
else : 
    print(response.status_code)

