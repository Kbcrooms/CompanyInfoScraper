import requests, re, html2text
from bs4 import BeautifulSoup

website = "www.sysplus.com/"
foundEmails = []
searchQuery = re.search('(?:www.)?([a-zA-Z0-9]+(?:-[a-zA-Z0-9]+)*(?:\.[a-zA-Z]+)+)',website,re.M|re.I)
searchEmail = "info%40"+searchQuery.group(1)
for i in range(0,5):

    searchPage = '&start='+ str(i*10)
    searchURL = "https://google.com/search?q=" + searchEmail + searchPage
    html = requests.get(searchURL)
    text_maker = html2text.HTML2Text()
    text_maker.ignore_emphasis = True

    htmlText = text_maker.handle(BeautifulSoup(html.content,"html5lib").prettify())
    foundEmails += re.findall(r'([\-\w\d.]+\s*\@\s*'+re.escape(searchQuery.group(1))+')',htmlText)
    #print("CONTENT START")
    #print(htmlText)
    print (htmlText)
    #print("CONTENT END")
    #print('[-\w\d.]+\@'+re.escape(searchQuery.group(1)))
for email in foundEmails:
    print("".join(email.split()))
#print(foundEmails)
