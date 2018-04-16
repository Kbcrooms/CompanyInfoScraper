import requests
url = "https://google.com/search?q=pokemon"
html = requests.get(url)

print(html.content)
