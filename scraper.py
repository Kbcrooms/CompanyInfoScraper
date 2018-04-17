import requests, re, html2text
from bs4 import BeautifulSoup
from openpyxl import load_workbook, Workbook
from collections import Counter
wb = load_workbook('CompanyProfilespja_MASTER.xlsx')
ws = wb[wb.sheetnames[2]]
startRow = 2800
endRow = 2810
chosenEmail = ""
chosenPN = ""
scrubbedName = "Khristan"
for row in range(2800,2801):
    prevScrubber = ws['B'+str(row)].value
    if  prevScrubber == scrubbedName or prevScrubber == "":
        website = ws['U'+str(row)].value
        #website = 'sysplus.com'
        foundEmails = []
        foundPhoneNumbers = []
        searchQuery = re.search('(?:www.)?([a-zA-Z0-9]+(?:-[a-zA-Z0-9]+)*(?:\.[a-zA-Z]+)+)',website,re.M|re.I)
        if searchQuery:
            searchEmail = "info%40"+searchQuery.group(1)
            for i in range(0,3):

                searchPage = '&start='+ str(i*10)
                searchURL = "https://google.com/search?q=" + searchEmail + searchPage
                html = requests.get(searchURL)
                text_maker = html2text.HTML2Text()
                text_maker.ignore_emphasis = True

                htmlText = text_maker.handle(BeautifulSoup(html.content,"html5lib").prettify())
                foundEmails += re.findall(r'([\-\w\d.]+\s*\@\s*'+re.escape(searchQuery.group(1))+')',htmlText)
                foundPhoneNumbers += re.findall(r'(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})',htmlText)
                #print("CONTENT START")
                #print(htmlText)
                #print (htmlText)
                #print("CONTENT END")
            for i in range(0,len(foundEmails)):
                foundEmails[i]="".join(foundEmails[i].lower().split())
            commonEmails = Counter(foundEmails).most_common(2)
            eLen = len(commonEmails)
            if(eLen>0):
                if(eLen>1 and commonEmails[1][1]>=3):
                    chosenEmail = commonEmails[1][0]
                else:
                    chosenEmail = commonEmails[0][0]
            print(chosenEmail)
            for i in range(0,len(foundPhoneNumbers)):
                foundPhoneNumbers[i]=re.sub("[^0-9]","", foundPhoneNumbers[i])
            commonPN = Counter(foundPhoneNumbers).most_common(len(foundPhoneNumbers))
            pnLen = len(commonPN)
            for pn in commonPN:
                if(len(pn[0])==10):
                    chosenPN = pn[0][:3] + "-" + pn[0][3:6] + "-" + pn[0][6:]
            print(chosenPN)
            ws['O'+str(row)] = chosenEmail
            ws['Q'+str(row)] = chosenPN
            ws['B'+str(row)] = scrubbedName
wb.save('CompanyProfilespja_MASTER.xlsx')
