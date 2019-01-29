#!/usr/bin/python
# -*- coding: utf-8 -*-

# pip install bs4
# pip install xlsxwriter

try:
    from urllib.request import urlopen, Request  # Python 3
except ImportError:
    from urllib2 import urlopen, Request         # with python < 3 
from bs4 import BeautifulSoup as soup
from operator import itemgetter
import xlsxwriter as xlsx


try: input = raw_input
except NameError: pass

# cities from user input   
citiesinput = input("Cities you want to scrape separated by comma: ")
cities = citiesinput.replace(' ','').replace('ä','a').replace('ö','o').replace('å','o').split(',')

# scraping tripadvisor links from google
urls = [] 
for city in cities:
    url = '{}{}{}'.format(
        "https://www.google.com/search?q=tripadvisor+", city, "+hotellit&gws_rd=cr")
    q = Request(url)
    # Google blocks scraping without user-agent
    q.add_header('user-agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.77 Safari/537.36')
    try:
        client = urlopen(q)
    except:
        print("Invalid input, don't use special characters.")
        input("Press enter to quit")    
    googlehtml = client.read()
    client.close()
    page = soup(googlehtml, "html.parser")
    tripadvisorurl = page.find("h3", {"class": "LC20lb"}).parent['href']
    urls += [(tripadvisorurl)]


# Scraping cities from tripadvisor
for u in urls:
    url = u
    citynumber = url[34:].partition('-')[0]
    cityname = url[42:].partition('_')[0].partition('-')[0]
    if cityname == "":  
        input("City not found, press enter to continue")
        continue
    listofhotels = []
    pagenro = 0
    revisiontimes = 20      # 20 times revision to collect hotels missing due to dynamic pages of tripadvisor
    revision = 0            # Counting times revised
    revisionpages = 6       # default value of how many pages checked in one revision
    

    # Scraping pages of one city
    while True:

        # Open Connection
        client = urlopen(url)
        pagehtml = client.read()
        client.close()
        page = soup(pagehtml, "html.parser")
        pageurl = page.find("link", {"hreflang": "fi"})["href"]
        print(pageurl)
    
        # Move to revision if all the pages have been scraped
        if (pagenro > 0 and url[:44] != pageurl[:44]):
            revisionpages = min((pagenro / 30),20)
            pagenro = 0
            revision += 1
            print("Revision " + str(revision) + "/" + str(revisiontimes))

    
        # Fetch link to every hotel of the page + price of each hotel
        hotels = page.findAll("div", {"class": "listing_title"})
        prices = page.findAll("div", {"data-clickpart": "chevron_price"})
        reviews = page.findAll("a", {"class": "review_count"})
        hotelIndex = 0
        scrapednames = [i[0] for i in listofhotels]

        # Scraping data from hotel-pages
        for hotel in hotels:

            # Fetch the name of a hotel
            name = hotel.a.text
            if name in scrapednames:
                hotelIndex += 1
                continue


            # Price for a night if exists in tripadvisor
            try:
                price = prices[hotelIndex].text
            except:
                price = "Not Found"

            # Number of reviews of a hotes
            review = int(reviews[hotelIndex].text[:-10].encode('ascii', 'ignore').decode('ascii'))


            sublink = "https://www.tripadvisor.fi" + hotel.a["href"]

            # Open subConnection
            subClient = urlopen(sublink)
            subpagehtml = subClient.read()
            subClient.close()
            subpage = soup(subpagehtml, "html.parser")

            # Fetch the rating
            try:
                rating = str(subpage.find(
                    "span", {"class": "ui_star_rating"}))[33:35]
                if len(rating) == 0:
                    rating = -1
                else:
                    rating = float(rating) / 10
            except Exception: 
                rating = -1
                print("ratingexception occuren")
                pass

            # Fetch the number of rooms
            rooms = subpage.find("div", text="Huonemäärä")

            # Ensuring that the number of rooms was fetched
            try:
                for x in range(max(2, (8 - int(pagenro / 30)))):
                    if rooms:
                        rooms = str(rooms.findNext("div").text.encode('ascii', 'ignore').decode('ascii'))
                        break
                    else:
                        subClient = urlopen(sublink)
                        subpagehtml = subClient.read()
                        subClient.close()
                        subpage = soup(subpagehtml, "html.parser")
                        rooms = subpage.find("div", text="Huonemäärä")
                if (not rooms) or type(rooms) is not str:
                    rooms = -1
                else:
                    rooms = float(rooms)
            except Exception: 
                rooms = -1
                print("roomexception occured")
                pass

            # Add scraped data to list of hotels
            listofhotels += [(name,rating,rooms,price,review)]

            # Next hotel of the page
            hotelIndex += 1

            

            print(name)
          #  print("Price: " + price)
          #  print("Reviews: " + review)
          #  print("Stars: " + rating)
          #  print("Rooms: " + rooms + "\n")

        # Next page + reset the hotel index
        pagenro += 30
        hotelIndex = 0

        # If revision is ready break loop
        if revision > 0 and pagenro >= revisionpages * 30:
            if revision == revisiontimes:
                break
            else:
                revision += 1
                print("Revision " + str(revision) + "/" + str(revisiontimes))
                pagenro = 0

        # Url of next page of the city
        url = "{}{}{}{}".format(
            "https://www.tripadvisor.fi/Hotels-", citynumber, "-oa", pagenro)
    
    # City scraped, create, write and format the xlsx sheet
    filename = "{}{}".format(cityname, 'Hotels.xlsx')
    workbook = xlsx.Workbook(filename, {'strings_to_numbers': True})
    worksheet = workbook.add_worksheet()

    headformat = workbook.add_format({'bold': True, 'font_size': 11, 'font_name': 'Arial', 'align': 'center'})
    nameformat = workbook.add_format({'font_size': 11, 'font_name': 'Arial'})
    dataformat = workbook.add_format({'font_size': 10, 'font_name': 'Arial', 'align': 'center'})

    headers = ('Name', 'Stars', 'Rooms', 'Price', 'Reviews')
    worksheet.write_row(0,0, headers)

    # remove duplicates
    listofh = list(set(listofhotels))

    # sort the list by stars and rooms
    sortedlist = sorted(listofh, key=itemgetter(1,2), reverse=True)

    # filter unwanted hotels 
    a = 1
    for h in sortedlist:
        worksheet.write_row(a,0,h)
        if (h[2] < 50 and h[2] >= 0) or h[4] < 21:
            worksheet.set_row(a, options={'hidden': True})
        a += 1

    worksheet.autofilter(0,1,4,a)
    worksheet.filter_column(2, 'Rooms >= 50 or Rooms < 0')
    worksheet.filter_column(4, 'Reviews > 20')

    # format the file
    worksheet.set_column('A:A', 50, nameformat)
    worksheet.set_column('B:E', 10, dataformat)
    worksheet.set_row(0, None, headformat)
    workbook.close()

    # Script finished
    print("All hotels scraped, " + filename + " -file created")
    close = input("Press enter to exit")
