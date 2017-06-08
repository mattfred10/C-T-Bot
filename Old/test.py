import PyPDF2
import re
import glob
import csv
import datetime

#set output file name
outputfilename = datetime.datetime.now().strftime('%y-%m-%d_%H%M_SalesOrders.csv')
POContents = [['Company', 'PO Number', 'Due Date', 'Item Number', 'Quantity', 'Price per item', 'Total order']]
#iterate over files in the folder
#will eventually set this as the download folder
for file in glob.glob('./*.pdf'):
    #create filename for scaped text
    textfile = file.replace('.pdf', '').replace('.PDF', '').replace('.\\','') + '_scrapedtext.txt'
    #create empty string for PDF contents
    PDFContents = ''
    #open text file
    with open(textfile, 'w') as pdfout:
        pdfFileObj = open(file, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        #combine all pages and concat into single string
        for page in range(0, pdfReader.numPages):
            pageObj = pdfReader.getPage(page)
            pagetext = pageObj.extractText()
            # write scrapped pdf text for testing purposes
            pdfout.write(pagetext)
            PDFContents = PDFContents + pagetext


        #GE - refactor this whole section
        if 'GE Renewables' in pdfReader.getPage(0).extractText():
            POneeded = True  # variable checks for whether the PO has been found in GE orders. Long POs sometimes have po number on second page

            for page in range(0, pdfReader.numPages):
                pageObj = pdfReader.getPage(page)
                pagetext = pageObj.extractText()
                # write scrapped pdf text for testing purposes
                pdfout.write(pagetext)

                if POneeded:
                    POindex = pagetext.find('By:Purchase Order Number')
                    PO = pagetext[POindex + 22:POindex + 35]  # taking an extra digit just in case
                    # will probably want to throw an exception here in case PO format/length changes
                    PONumber = re.sub('[^0-9]', '', PO)
                    POneeded = False

                deliveryindex = pagetext.find('Delivery Schedule:')
                itemindex = pagetext.find('GE Item:')
                #
                #
                #Refactor this to use RE - need more samples
                #
                #
                while deliveryindex != -1 and itemindex != -1:
                    #date
                    date = pagetext[deliveryindex + 18:deliveryindex + 27]

                    #quantity
                    quantityend = pagetext[deliveryindex + 28:deliveryindex + 39].find('EACH')  # any other units add here
                    quantity = re.sub('[^0-9]', '', pagetext[deliveryindex + 27:deliveryindex + 33]) #why was this necessary? need to check. Does it include variable amounts of 'EACH'

                    #iteamnumber
                    itemnumber = pagetext[itemindex + 9:itemindex + 21]

                    #price
                    priceindex = pagetext.find('Hazard Code:') #need to move backward
                    prices = pagetext[priceindex - 25:priceindex - 10]
                    combinedprices = str(re.sub('[^0-9.]', '', prices))

                    #these values don't separate in the GE files use the number sold to find correct values
                    #throw an exception if they aren't found.
                    for n, c in enumerate(combinedprices):

                        firsthalf = combinedprices[0:0 + n]
                        if firsthalf == '.' or firsthalf == '':
                            first = 0.0
                        else:
                            first = float(firsthalf)

                        secondhalf = combinedprices[n:]
                        if first*float(quantity) == float(secondhalf):
                            priceper = first
                            totalprice = secondhalf
                            break

                    pagetext = pagetext[priceindex + 26:]
                    deliveryindex = pagetext.find('Delivery Schedule:')
                    itemindex = pagetext.find('GE Item:')  # need pretty precise search string here because there are many instances of Description.
                    POContents.append(['GEC01', PONumber, date, itemnumber, quantity, priceper, totalprice])

        #Vestas America (Vest01)
        elif 'Vestas Nacelles America' in pdfReader.getPage(0).extractText():
            #This may be unecessary. Only see 1 page documents so far.
            for page in range(0, pdfReader.numPages):
                pageObj = pdfReader.getPage(page)
                pagetext = pageObj.extractText()
                # write scrapped pdf text for testing purposes
                pdfout.write(pagetext)
                #use this to monitor the presence of multipage vestas orders - can use the POneeded bool to skip PO if it becomes an issue
                if pdfReader.numPages > 1:
                    print(pdfReader.numPages)
                #Get PO Number
                PO = re.search(r'P(K|1)[0-9]{5}', pagetext)  #Revision number and page number run right into this so it needs to be tight. If Vestas expands to
                PONumber = PO[0] #could just stick this in the output, but I want the write line to be consistent
                item = re.search(r'(    1200)[0-9]{6,}', pagetext)
                itemnumber = item[0].replace('    1200','')
                #1240 appears to be a formatting code that the scraper finds. Added the leading spaces to make sure item numbers or the line don't get caught
                #Everything esle is searched for and indexed from here.
                totalline = pagetext.find('     1240')
                dategroup = pagetext[totalline+9:totalline + 17].strip().split('.')
                date = '-'.join(dategroup)
                quantityindex = pagetext.find('EA     ')

                quantity = pagetext[totalline+18:quantityindex].strip().replace(',','.')
                prices = pagetext[quantityindex+2:quantityindex+50].strip().replace(',','.').split()
                priceper = prices[0]
                totalprice = prices[1]

                if quantityindex == -1:
                    print('Error! Unit quantities were not EA!')
                if float(quantity)*float(prices[0]) != float(prices[1]):
                    print("Error! Prices and quantities do not match!")
                    print(float(quantity) * float(prices[0]), prices[1])

            POContents.append(['Vest01', PONumber, date, itemnumber, quantity, priceper, totalprice])

        #Vest02 and Vest04 use the same format
        elif 'Vestas - American Wind Technology' in pdfReader.getPage(0).extractText():
            # This may be unnecessary. Only see 1 page documents so far.
            for page in range(0, pdfReader.numPages):
                pageObj = pdfReader.getPage(page)
                pagetext = pageObj.extractText()
                # write scrapped pdf text for testing purposes
                pdfout.write(pagetext)
                # use this to monitor the presence of multipage vestas orders - can use the POneeded bool to skip PO if it becomes an issue
                if pdfReader.numPages > 1:
                    print(pdfReader.numPages)
            # Get PO Number
            PO = re.search(r'Purchase order ([0-9]{1,})', pagetext)
            PONumber = PO[1]

            #multiple items per order
            #counts occurrences of 'Delivery date' as it appears once per item at the end of the item
            alldates = re.findall(r'Delivery date: ([0-9]{1,2}) ([A-z]{3}) ([0-9]{4})', pagetext)
            #regex pattern 10,20,30, etc item line + (part number) + spaces + (quantity) + EA + spaces + (price per) + (total item price)
            itemline = re.findall(r'[0-9]{1}0([0-9]{1,})[ ]{1,}([0-9]{1,}) EA[ ]{1,}([0-9.,]{1,})[ ]{1,}([0-9.,]{1,})', pagetext)
            POTotal = re.findall(r'Net value[ ]{1,}([0-9.,]{1,})', pagetext)[0].replace(',', '')

            #Need to iterate over these item lists and assign values. Not sure how many are in each.
            #One delivery date per item is used as a proxy for total number of items
            #If items span pages (haven't seen it happen yet), this procedure will need to be adapted (maybe just concat all pages at beginning?)
            #First, we keep track of our total value by adding the line items and check it against the reported PO value at the end
            itemtotal = 0
            for i, date in enumerate(alldates):
                itemnumber = itemline[i][0]
                quantity = itemline[i][1].replace(',', '')
                priceper = itemline[i][2].replace(',', '')
                totalprice = itemline[i][3].replace(',', '')
                POContents.append(['Vest02', PONumber, date, itemnumber, quantity, priceper, totalprice])
                if float(priceper) * int(quantity) > float(totalprice) + 1 or float(priceper) * int(quantity) < float(totalprice) - 1: #Set this as +/- 1 to deal with floating point precision errors
                    print('Per: ' + priceper + ' Quantity: ' + quantity + ' Total: ' + totalprice + ' Calcd: ' + str(float(priceper) * int(quantity)))
                    print('Error! Prices and quantities do not match!')
                itemtotal += float(totalprice)
            if float(itemtotal) != float(POTotal):
                print('Calcd: ' + totalprice + ' PO: ' + POTotal)
                print('Error! Total price not sum of individual items!')

        #Vest04 - very similar to vest02 but there are some spacing issues that are different
        elif 'Vestas Do Brasil Energia' in pdfReader.getPage(0).extractText():
            # This may be unecessary. Only see 1 page documents so far.

            # Get PO Number
            PO = re.search(r'Purchase order ([0-9]{1,})', pagetext)
            PONumber = PO[1]

            #multiple items per order
            alldates = re.findall(r'Delivery date: ([0-9]{1,2}) ([A-z]{3}) ([0-9]{4})', pagetext)
            # regex pattern = not space to avoid matching 201X from the date + 10,20,30, etc line item number + (part number/quantity together) + ' EA' + (price per)(total item price)
            alldata = re.findall(r'[^ ][1-9]0([0-9,]{1,}) EA([0-9.,]{1,})', pagetext)
            POTotal = re.findall(r'Net value[ ]{1,}([0-9.,]{1,})', pagetext)[0].replace(',', '')
            #Separate the combined terms
            #Individual item prices have 2 decimal places and should always occur first in the combined string
            priceper = re.match(r'[0-9]{1,}.[0-9]{2}', alldata[0][1])[0]
            totalprice = float(alldata[0][1].replace(priceper, '').replace(',',''))
            quantity = totalprice/float(priceper.replace(',',''))
            #need to insert commas so that we make sure we match the quantity and not another part of the item string
            commaquantity = format(int(quantity),',d')
            itemnumber = alldata[0][0].replace(commaquantity,'')


            itemtotal = 0
            for i, date in enumerate(alldates):
                priceper = re.match(r'[0-9]{1,}.[0-9]{2}', alldata[i][1])[0]
                totalprice = float(alldata[i][1].replace(priceper, '').replace(',', ''))
                quantity = totalprice / float(priceper.replace(',', ''))
                # need to insert commas so that we make sure we match the quantity and not another part of the item string
                commaquantity = format(int(quantity), ',d')
                itemnumber = alldata[i][0].replace(commaquantity, '')
                #need a better solution here. will somethin


                POContents.append(['Vest04', PONumber, date, itemnumber, quantity, priceper, totalprice])
                if float(priceper) * int(quantity) > float(totalprice) + 1 or float(priceper) * int(quantity) < float(totalprice) - 1: #Set this as +/- 1 to deal with floating point precision errors
                    print('Per: ' + priceper + ' Quantity: ' + quantity + ' Total: ' + totalprice + ' Calcd: ' + str(float(priceper) * int(quantity)))
                    print('Error! Prices and quantities do not match!')
                itemtotal += float(totalprice)
            if float(itemtotal) != float(POTotal):
                print('Calcd: ' + totalprice + ' PO: ' + POTotal)
                print('Error! Total price not sum of individual items!')

        elif 'Frontier Technologies Brewton' in pdfReader.getPage(0).extractText():
            for page in range(0, pdfReader.numPages):
                pageObj = pdfReader.getPage(page)
                pagetext = pageObj.extractText()
                # write scrapped pdf text for testing purposes
                pdfout.write(pagetext)
                # use this to monitor the presence of multipage vestas orders - can use the POneeded bool to skip PO if it becomes an issue

                #PO and date are combined
                #Need to verify that date is of format DD-Mon-YY
                PO = re.search(r'America([0-9]{1,})[0-9]{2}-', pagetext)[1]
                #Regex = ea + $(priceper) + $(total) + Due: (date) + (itemnumber) ea$5.12 $3,584.00 Due:04-Aug-17105W1931P016700H
                itemline = re.findall(r'ea\$([0-9,.]{1,}) \$([0-9,.]{1,}) Due:([0-9]{2})-([A-z]{3})-([0-9]{2})([A-Z0-9]{1,})', pagetext)
                print(itemline)

                if 'Total:' in pagetext:
                    break
                #exit loop if entire order fits on first page
                if pdfReader.numPages > 1:
                    print(pdfReader.numPages)
                    print('here')

        else:
            for page in range(0, pdfReader.numPages):
                pageObj = pdfReader.getPage(page)
                pagetext = pageObj.extractText()
                # write scrapped pdf text for testing purposes
                pdfout.write(pagetext)
            print("PO not recognized!")

#including hours and minutes so that this program can be run twice in one day


    #print(POContents)
    #with open(outputfilename, 'w', newline='') as outfile:
    #    itemwriter = csv.writer(outfile, delimiter=",")
    #    for item in POContents:
    #        itemwriter.writerow(item)
