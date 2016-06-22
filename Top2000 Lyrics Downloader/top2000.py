# encoding: utf-8

import xlrd as x
import mechanize as m
import sys
import urllib
import time
from bs4 import BeautifulSoup as BS
import sys
from string import maketrans, punctuation

#set start to provided index
if (len(sys.argv) == 1):
    start = 0
else:
    start = int(sys.argv[1])

  
datalist = []
artistlist = []

#open the excel sheet
sheet = x.open_workbook("TOP-2000-2015.xls").sheet_by_index(0)

#put data in list
for cell in sheet.col_values(1):
    datalist.append(cell)

for cell in sheet.col_values(2):
    artistlist.append(cell)
    

hdr = 'User-Agent'
exclude = set(punctuation)
exclude.remove('-')

for i in range(start, len(datalist)):
    item = datalist[i]
    if (isinstance(item, float)):
        item = str(int(item))
    if (isinstance(item, basestring)):
        item = item.lower().replace(" ", "-").replace(u'’', u"'")
    item = item.replace(u'\xa0', u' ').encode('utf-8')
    
    artist = artistlist[i]
    artist = artist.replace(u'\xa0', u' ')
    artist = artist.lower().replace(" ", "-").encode('utf-8')
    print i, item, artist
    
    artistList = []
    for char in artist:
        if char in exclude:
            char = '-'
        artistList.append(char)
    artist = ''.join(artistList)
    print artist

    if item != "titel":
        if (isinstance(item, basestring)):
            item = item.replace(" ", "+")
            
        url = "http://songteksten.nl/search?query=" + item + "&type=title"
        request = m.Request(url)
        request.add_header(hdr, "http://wwwsearch.sourceforge.net/mechanize/")
        response = m.urlopen(url)
        data = response.read()
        soup = BS(data, 'html.parser')
        links = soup.find_all('a')
        
        #
        for j in range(len(links)): 
            link = links[j].get('href')
            if "'" in item:
                item = item.replace("'", "-")
                print item
                
            if (item.decode('utf-8') in link and artist.decode('utf-8') in link):
                url2 = "http://songteksten.nl" + link
                request2 = m.Request(url2)
                request2.add_header(hdr, "http://wwwsearch.sourceforge.net/mechanize/")
                response2 = m.urlopen(request2)
                data2 = response2.read()
                soup2 = BS(data2, 'html.parser')
                text = soup2.find('p')
                text = str(text).split(">")[2::] #change soup object to string, split at > to split at tag ends, remove the style and span tags
                outtext = []
                f = open("{} - {} - {}.txt".format(i, artist, item), 'w')
                for string in text:
                    string = string.replace("<br/", "")
                    string = string.replace("\r\n", "")
                    string = string.replace("\n", "")
                    if string == "" or string == "</p":
                        continue
                    elif "</span" in string:
                        string = string.split("</span")[0] #split at </span, select the part of the string without that and join it back to a string
                    f.write(string + "\n")
                    outtext.append(string)
                
                f.close()
                #print outtext
                break
        #time.sleep(1)


    
