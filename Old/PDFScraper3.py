import PyPDF2
import re
import glob
import csv
import datetime
import os

#Consider creating a dictionary of prices as an additional error check, that is, pricedict[(company, partnumber)] = priceper
class POScraper():


    def scrapeOne(self, filename):


    def outputPDFText(self):
        #create filename for scaped text
        textfile = originalPDF.replace('.pdf', '').replace('.PDF', '').replace('.\\','') + '_scrapedtext.txt'
        scrapeText()


    def scrapeText(self):
        #create empty string for PDF contents
        self.PDFContents = ''

        #open PDF
        pdfFileObj = open(self.originalPDF, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        #combine all pages and concat into single string
        for page in range(0, pdfReader.numPages):
            pageObj = pdfReader.getPage(page)
            pagetext = pageObj.extractText()
            self.PDFContents = self.PDFContents + pagetext

        # Close the file. Everything is done using the scraped PDFreader - This needs to be closed so we can move the file at the end
        pdfFileObj.close()


    def scrapeFolder(self, directory='./'):

    #POContents is the output. Create the header here.
    POContents = [['Company', 'PO Number', 'Due Date', 'Item Number', 'Quantity', 'Price Per Item', 'Order Total']]

    # create path for PDF (filename created after PO number acquired)
    datepath = datetime.date.today().strftime('%Y/%m/%d')  # date format for directory structure
    PDFDate = datetime.date.today().strftime('%Y-%m-%d')  # date format for file name

    #iterate over files in the folder
    #will eventually set this as the download folder
    for originalPDF in glob.glob(directory + '*.pdf'):
        #create filename for scaped text
        textfile = originalPDF.replace('.pdf', '').replace('.PDF', '').replace('.\\','') + '_scrapedtext.txt'

        #create empty string for PDF contents
        PDFContents = ''

        #open PDF
        pdfFileObj = open(originalPDF, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        #combine all pages and concat into single string
        for page in range(0, pdfReader.numPages):
            pageObj = pdfReader.getPage(page)
            pagetext = pageObj.extractText()
            PDFContents = PDFContents + pagetext

        # Close the file. Everything is done using the scraped PDFreader - This needs to be closed so we can move the file at the end
        pdfFileObj.close()

        # write scrapped pdf text for testing purposes
        # open text file and write the PDFContents
        with open(textfile, 'w') as pdfout:
            pdfout.write(PDFContents)

        #GE
        if 'GE Renewables' in PDFContents:
            company = 'GEC01'

            PONumber = re.search(r'Order Number[\s]+([0-9]+)', PDFContents)[1]
            POTotal = re.search(r'Total Amt:([0-9.,]+)', PDFContents)[1]

            datequantity = re.findall(r'Delivery Schedule:([0-9]+)-([A-Z]{3})-([0-9]{2})([0-9,]+)[ ]+EACH', PDFContents)
            partnumbers = re.findall(r'GE Item: ([A-Z0-9]+)[\s]+Rev: ([0-9]+)', PDFContents) #excluding revision for now


            #here we are moving backwards from 'Hazard Code' because the text in front is unreliable (text behind may be too if they optionally include 'promis shipment date' or 'need by shipment date')
            #we capture everything including the day of month in case they change from 2 digits to 1 digit
            prices = re.findall(r'([0-9.]+)-[A-Z]{3}-[0-9 \n]+Hazard', PDFContents)

            sumofitems = 0
            #using delivery dates as a proxy for number of items in order
            for i,dq in enumerate(datequantity):
                # first we will try to strip the day of month using the date grabbed elsewhere
                #check if it's one digit or two
                if prices[i][-1] == dq[0]:
                    #take all but the last number
                    combinedprice = prices[i][0:-1]
                else:
                    #take all but the last two numbers
                    combinedprice = prices[i][0:-2]

                date = (dq[0], dq[1], dq[2])
                quantity = dq[3]

                #The price per and item total don't separate in the GE files. Use the quantity sold to find correct values
                for n, c in enumerate(combinedprice):
                    firsthalf = combinedprice[0:0 + n]
                    if firsthalf == '.' or firsthalf == '':
                        first = 0.0 #used below
                    else:
                        first = float(firsthalf)

                    secondhalf = combinedprice[n:]
                    if first*float(quantity) == float(secondhalf):
                        priceper = first
                        itemtotal = secondhalf
                        break

                sumofitems += float(itemtotal)
                partnumber = partnumbers[i][0]

                ####output####
                POContents.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])

            if float(sumofitems) != float(POTotal):
                print('Error! Total price not sum of individual items!')
                print(originalPDF)
                print(company + ' Calcd: ' + sumofitems + ' PO: ' + POTotal)

        #Vestas America (Vest01) - Handles sjol01
        elif 'Vestas Nacelles America' in PDFContents:
            #currently only handles 1 item per PO. Should be easy to fix if it comes up. Need to see an example first.
            if 'SJOELUND US INC.' in PDFContents:
                #do something here about the different ways to handle different companies
                #may need to move this to the end when data are written
                #
                company = 'sjol01'
            else:
                company = 'VEST01'

            #Get PO Number
            PO = re.search(r'P(K|1)[0-9]{5}', PDFContents)  #Revision number and page number run right into this so it needs to be tight.
            PONumber = PO[0] #could just stick this in the output, but I want the write line to be consistent
            #1200, 1240 appear to be formatting information
            #if there is more than one item, it's going to need to change
            item = re.search(r'    1200([0-9]{6,})', PDFContents)
            partnumber = item[1]

            dategroup = re.search(r'1240([0-9 ]{2})\.([0-9]+)\.([0-9]+)', PDFContents)
            date = (dategroup[1], dategroup[2], dategroup[3])

            pricesandquantity = re.search(r'([0-9,]+)EA[ ]+([0-9,]+)[ ]+([0-9,]+)', PDFContents)

            quantity = pricesandquantity[1].replace(',', '.')
            priceper = pricesandquantity[2].replace(',','.')
            POTotal = pricesandquantity[3].replace(',','.')

            #if there is more than one line item, will need to find a way to collect PO total.
            if float(quantity)*float(priceper) < float(POTotal) - 1 or float(quantity)*float(priceper) > float(POTotal) + 1:
                print("Error! Prices and quantities do not match!")
                print(originalPDF)
                print(company, float(quantity) * float(priceper), POTotal)

            ####output####
            POContents.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])

        #Vest02 and Vest04 use the same format
        elif 'Vestas - American Wind Technology' in PDFContents:
            company = 'VEST02'

            # Get PO Number

            PO = re.search(r'Purchase order ([0-9]+)', PDFContents)
            PONumber = PO[1]

            #multiple items per order
            #counts occurrences of 'Delivery date' as it appears once per item at the end of the item
            alldates = re.findall(r'Delivery date: ([0-9]{1,2}) ([A-z]{3}) ([0-9]{4})', PDFContents)
            #regex pattern 10,20,30, etc item line + (part number) + spaces + (quantity) + EA + spaces + (price per) + (total item price)
            itemline = re.findall(r'[0-9]{1}0([0-9]+)[ ]+([0-9]+) EA[ ]+([0-9.,]+)[ ]+([0-9.,]+)', PDFContents)
            POTotal = re.findall(r'Net value[ ]+([0-9.,]+)', PDFContents)[0].replace(',', '')

            #Need to iterate over these item lists and assign values. Not sure how many are in each.
            #One delivery date per item is used as a proxy for total number of items
            #If items span pages (haven't seen it happen yet), this procedure will need to be adapted (maybe just concat all pages at beginning?)
            #First, we keep track of our total value by adding the line items and check it against the reported PO value at the end
            sumofitems = 0
            for i, date in enumerate(alldates):
                partnumber = itemline[i][0]
                quantity = itemline[i][1].replace(',', '')
                priceper = itemline[i][2].replace(',', '')
                itemtotal = itemline[i][3].replace(',', '')


                #####output#####
                POContents.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])

                if float(priceper) * int(quantity) > float(itemtotal) + 1 or float(priceper) * int(quantity) < float(itemtotal) - 1: #Set this as +/- 1 to deal with floating point precision errors
                    print('Error! Prices and quantities do not match!')
                    print(originalPDF)
                    print(company + ' Price per: ' + priceper + ' Quantity: ' + quantity + ' Total: ' + itemtotal + ' Calcd: ' + str(float(priceper) * int(quantity)))

                sumofitems += float(itemtotal)
            if float(sumofitems) != float(POTotal):
                print('Error! Total price not sum of individual items!')
                print(originalPDF)
                print(company + ' Calcd: ' + sumofitems + ' PO: ' + POTotal)

        #Vest04 - very similar to vest02 but there are some spacing issues that are different
        elif 'Vestas Do Brasil Energia' in PDFContents:
            company = 'vest04'

            # Get PO Number
            PO = re.search(r'Purchase order ([0-9]+)', PDFContents)
            PONumber = PO[1]

            #multiple items per order
            alldates = re.findall(r'Delivery date: ([0-9]{1,2}) ([A-z]{3}) ([0-9]{4})', PDFContents)
            # regex pattern = not space to avoid matching 201X from the date + 10,20,30, etc line item number + (part number/quantity together) + ' EA' + (price per)(total item price)
            alldata = re.findall(r'[^ ][1-9]0([0-9,]+) EA([0-9.,]+)', PDFContents)

            POTotal = re.search(r'Net value[ ]+([0-9.,]+)', PDFContents)[1].replace(',', '')


            sumofitems = 0
            for i, date in enumerate(alldates):
                # Separate the combined terms
                priceper = re.match(r'[0-9]+.[0-9]{2}', alldata[i][1])[0]
                itemtotal = float(alldata[i][1].replace(priceper, '').replace(',', ''))
                #don't have a great way to separate the quantity from the part number (comes as '290107241,000' with quantity as 1,000 or '153452600' with quantity as 600)
                #need to calculate it from the item total and the priceper (both of which we have at high confidence
                #Should double check with price dictionary
                #this is dangerous though because of rounding errors.
                quantity = float(itemtotal) / float(priceper.replace(',', ''))
                # need to insert commas so that we make sure we match the quantity and not another part of the item string
                commaquantity = format(int(quantity), ',d')
                partnumber = alldata[i][0].replace(commaquantity, '')

                ####output####
                POContents.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])

                if float(priceper) * int(quantity) > float(itemtotal) + 1 or float(priceper) * int(quantity) < float(itemtotal) - 1: #Set this as +/- 1 to deal with floating point precision errors
                    print('Error! Prices and quantities do not match!')
                    print(originalPDF)
                    print(company + ' Price per: ' + str(priceper) + ' Quantity: ' + str(quantity) + ' Total: ' + str(POTotal) + ' Calcd: ' + str(float(priceper) * int(quantity)))

                sumofitems += float(itemtotal)
            if float(sumofitems) != float(POTotal):
                print('Error! Total price not sum of individual items!')
                print(originalPDF)
                print(company + ' Calcd: ' + str(sumofitems) + ' PO: ' + str(POTotal))

        #FRON01
        elif 'Frontier Technologies Brewton' in PDFContents:
                company = 'FRON01'

                #PO and date are combined
                #Need to verify that date is of format DD-Mon-YY
                PONumber = re.search(r'America([0-9]+)[0-9]{2}-', PDFContents)[1]
                POTotal = re.search(r'Total:\$([0-9.,]+)', PDFContents)[1].replace(',','')
                #need to sum individual items
                sumofitems = 0

                #Regex = ea + $(priceper) + $(total) + Due: (date) + (partnumber) ea$5.12 $3,584.00 Due:04-Aug-17105W1931P016700H - quantity comes between last letter (=revision) and P### (part of PO)
                itemline = re.findall(r'ea\$([0-9,.]+) \$([0-9,.]+) Due:([0-9]{2})-([A-z]{3})-([0-9]{2})([A-Z0-9]+)(P[0-9]{3})([0-9,]+)([A-Z])Rev', PDFContents)
                for item in itemline:
                    priceper = item[0]

                    itemtotal = item[1].replace(',', '')
                    date = (item[2], item[3], item[4])
                    partnumber = item[5] + item[6]
                    quantity = item[7].replace(',', '')
                    #
                    #
                    #Does revision need to be included in part number?
                    #
                    rev = item[8]

                    #check if individual line items sum
                    if int(quantity) * float(priceper) != float(itemtotal):
                        print('Per: ' + priceper + ' Quantity: ' + quantity + ' Total: ' + POTotal + ' Calcd: ' + str(float(priceper) * int(quantity)))
                        print('Error! Prices and quantities do not match!')
                    sumofitems += float(itemtotal)

                    ####output####
                    POContents.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])
                if float(sumofitems) != float(POTotal):
                    print('Calcd: ' + sumofitems + ' PO: ' + POTotal)
                    print('Error! Total price not sum of individual items!')

        else:
            print("PO not recognized!")
            print(originalPDF)

        #check if a folder has been made for a day - if not, create it
        #if not os.path.exists(datepath):
        #    os.makedirs(datepath)
        #move original file to new location
        #os.rename(originalPDF, datepath + '/' + PDFDate + '_' + company + '_' + PONumber + '.pdf')

    # set output file name
    # including hours and minutes so that this program can be run twice in one day
    outputfilename = datetime.datetime.now().strftime('%y-%m-%d_%H%M_SalesOrders.csv')

    with open(outputfilename, 'w', newline='') as outfile:
        itemwriter = csv.writer(outfile, delimiter=",")
        for item in POContents:
            itemwriter.writerow(item)
    #print(POContents)
    print('SOs successfully extracted!')

if __name__ == "__main__":
    scrapePO()
