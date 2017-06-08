


#item numbers and prices are checked
#POs are checked to be self consistent
#Dates need to be checked somehow. I will write a function to convert the varying formats to how the DB stores it and
#do some sanity checking at that time (e.g., valid date, delievry date not in past)

def scrapePO(outputPDFText=False, movePDF=False, origindirectory='.\\',):
    """I think this makes a lot more sense as a script than a class. I've considered refactoring it to a class, but I 
    don't think it will operate as clearly. The routines are all very similar but different enough """

    #store a list of POs with bad entries and their errors.
    errors = []

    #Going to store a list of files that are output for attachment to email - bad dictionary entries, error log, and good SOs
    logs = []

    #generate dictionary of prices to check that POs are correct
    pricedictionary = {}
    # catch any unlisted company/part combinations in this list and print it at the end
    # if they are correct, they can be added to the price dictionary.
    # ensure that the prices are stored with 2 decimal places (even for e.g., $0.30 - the trailing zero is required)
    nodictentry = []
    with open(pricedictionarypath) as dictfile:
        dictentries = csv.reader(dictfile)
        for line in dictentries:
            #use tuple of company (i.e., VEST01, etc) and item
            #companies have different prices
            pricedictionary[(line[0], line[1])] = line[2]

    #POContents is the output. Create the header here.
    POContents = [['Company', 'PO Number', 'Due Date', 'Item Number', 'Quantity', 'Price Per Item', 'Order Total']]

    # create path for PDF (filename created after PO number acquired)
      # date format for directory structure
    PDFDate = datetime.date.today().strftime('%Y-%m-%d')  # date format for file name

    if not os.path.exists('.\\'+datepath):
        os.makedirs('.\\'+datepath)

    #iterate over files in the folder
    #will eventually set this as the download folder
    for originalPDF in glob.glob(origindirectory + '*.pdf'):
        processed = False
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
        if outputPDFText:
            with open(textfile, 'w') as pdfout:
                pdfout.write(PDFContents)

        #Can't effectively eliminate engineering drawing from file names
        #Check for them first because they likely have some of the same inforamtion.
        #Might produce false positives
        if 'Draw. format' in PDFContents: #this is for Siemans - no idea if others will appaer
            company = 'unknown'
            errors.append([company, 'unknown', originalPDF, "Appears to be an enginnering drawing. Please double check the document."])
        #GE
        elif 'GE Renewables' in PDFContents:
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

            try:
                if float(pricedictionary[(company, partnumber)]) != float(priceper):
                    errors.append([company, PONumber, originalPDF, "Incorrect item price.", "got: " + str(priceper), "expected: " + pricedictionary[(company, partnumber)]])
                elif float(sumofitems) != float(POTotal):
                    errors.append([company, PONumber, originalPDF, "Incorrect total price or number of items.", 'Calcd: ' + str(sumofitems), 'PO: ' + POTotal])
                else:
                    ####output####
                    POContents.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])
                    processed=True
            except:
                nodictentry.append([company, partnumber, priceper])
                errors.append([company, PONumber, originalPDF, "No price dictionary entry - see 'NoDictionaryEntry' file"])

        #Vestas America (Vest01) - Handles some sjol01 - sjol01 also handled below (after all Vestas)
        elif 'Vestas Nacelles America' in PDFContents:
            #currently only handles 1 item per PO. Should be easy to fix if it comes up. Need to see an example first.
            if 'SJOELUND US INC.' in PDFContents:
                #do something here about the different ways to handle different companies
                #may need to move this to the end when data are written
                #I think the only difference is inputting price in the SO, which will be handled during input.
                company = 'sjol01'
            else:
                company = 'VEST01'

            #Get PO Number
            PO = re.search(r'P(K|1)[0-9]{5}', PDFContents)  #Revision number and page number run right into this so it needs to be tight (i.e., 5 numbers)
            PONumber = PO[0]
            #1200, 1240 appear to be formatting information
            #if there is more than one item, it's going to need to change
            item = re.search(r'    1200([0-9]{6,})', PDFContents)
            partnumber = item[1]

            #dates separated by .
            datepattern = """
            1240            #string that appears to be pdf formatting information and consistently leads this line
            ([0-9 ]{2})     #Day - strictly 2 digits because it runs into 1240. Other dates appear to always be 2 digits but are given flexibility in case format changes.
            \.              #escaped '.' which is used to divide the days, months and years
            ([0-9]+)        #Month
            \.              #escaped '.' which is used to divide the days, months and years
            ([0-9]+)        #Year
            """
            dategroup = re.search(datepattern, PDFContents, re.VERBOSE)
            #dategroup = re.search(r'1240([0-9 ]{2})\.([0-9]+)\.([0-9]+)', PDFContents)

            date = (dategroup[1], dategroup[2], dategroup[3])

            pqpattern = """
            ([0-9,]+)       #quantities allowing for thousands (,)
            EA[ ]+          #EA is the units and the number of spaces varies by the length of the quantity
            ([0-9,]+)       #price per item - european notation with not thousands separator
            [ ]+            #variable number of spaces after unit price
            ([0-9,]+)       #order total (quanity * unit price)
            """
            pricesandquantity = re.search(pqpattern, PDFContents, re.VERBOSE)
            #pricesandquantity = re.search(r'([0-9,]+)EA[ ]+([0-9,]+)[ ]+([0-9,]+)', PDFContents)

            quantity = float(pricesandquantity[1].replace(',', '.'))
            priceper = float(pricesandquantity[2].replace(',','.'))
            POTotal = float(pricesandquantity[3].replace(',','.'))

            try:
                if float(pricedictionary[(company, partnumber)]) != round(float(priceper),2):
                    errors.append([company, PONumber, originalPDF, "Incorrect item price.", "Got: " + str(priceper), "Expected: " + pricedictionary[(company, partnumber)]])
                elif quantity * priceper < POTotal - 1 or quantity * priceper > POTotal + 1:
                    errors.append([company, PONumber, originalPDF, "Incorrect total price or number of items.", 'Calcd: ' + float(quantity) * float(priceper), 'PO: ' + POTotal])
                else:
                    ####output####
                    POContents.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])
                    processed = True
            except:
                nodictentry.append([company, partnumber, priceper])
                errors.append([company, PONumber, "No price dictionary entry - see 'NoDictionaryEntry' file"])

        #Vest02 and Vest04 use very similar formats
        elif 'Vestas - American Wind Technology' in PDFContents:
            company = 'VEST02'
            OtherError = False

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

                try:
                    if float(pricedictionary[(company, partnumber)]) != float(priceper):
                        errors.append([company, PONumber, originalPDF, "Incorrect item price.", "Got: " + str(priceper), "Expected: " + pricedictionary[(company, partnumber)]])
                        OtherError = True
                    elif float(priceper) * int(quantity) > float(itemtotal) + 1 or float(priceper) * int(quantity) < float(itemtotal) - 1: #Set this as +/- 1 to deal with floating point precision errors
                        errors.append([company, PONumber, originalPDF, "Incorrect quantity or price for line item.", 'Calcd: ' + float(quantity) * float(priceper), 'PO: ' + POTotal])
                        OtherError = True
                except:
                    nodictentry.append([company, partnumber, priceper])
                    errors.append([company, PONumber, "No price dictionary entry - see 'NoDictionaryEntry' file"])
                    OtherError = True

                sumofitems += float(itemtotal)

            #This check only comes at end of PO. The above errors can occur on individual line items, so we need to wait for this check.
            if float(sumofitems) != float(POTotal):
                errors.append([company, PONumber, originalPDF, "Total price not sum of individual items.", 'Calcd: ' + sumofitems, 'PO: ' + POTotal])
            elif not OtherError: #actual error stored above
                #####output#####
                POContents.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])
                processed = True

        #Vest04 - very similar to vest02 but there are some spacing issues that are different
        elif 'Vestas Do Brasil Energia' in PDFContents:
            company = 'vest04'
            OtherError: False

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

                # check if individual line items sum correctly
                sumofitems += float(itemtotal)

                try:
                    if float(pricedictionary[(company, partnumber)]) != float(priceper):
                        errors.append([company, PONumber, originalPDF, "Incorrect item price.", "Got: " + str(priceper), "Expected: " + pricedictionary[(company, partnumber)]])
                        OtherError = True
                    elif float(priceper) * int(quantity) > float(itemtotal) + 1 or float(priceper) * int(quantity) < float(itemtotal) - 1:  # Set this as +/- 1 to deal with floating point precision errors
                        errors.append([company, PONumber, originalPDF, "Incorrect quantity or price for line item.",'Calcd: ' + float(quantity) * float(priceper), 'PO: ' + POTotal])
                        OtherError = True
                except:
                    nodictentry.append([company, partnumber, priceper])
                    errors.append([company, PONumber, "No price dictionary entry - see 'NoDictionaryEntry' file"])
                    OtherError = True

            # This check only comes at end of PO (end of for loop). The above errors can occur on individual line items, so we need to wait for this check.
            if float(sumofitems) != float(POTotal):
                errors.append([company, PONumber, originalPDF, "Total price not sum of individual items.",
                               'Calcd: ' + sumofitems, 'PO: ' + POTotal])
            elif not OtherError:  # actual error stored above
                #####output#####
                POContents.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])
                processed = True

        #sjol01 - direct sales - some other Sjoeland POs use Vestas POs and are handled above
        elif 'SJØLUND' in PDFContents:
            company = 'sjol01'

            #requisition number not used - line items each get an order number - PO becomes item1/item2 - see below
            # POpatt = """
            # REQUISITION         #requisition number identifier
            # [\s]+               #White space including new line
            # ([0-9]+)            #requisition number
            # """
            #
            # PONumber = re.search(POpatt, PDFContents, re.VERBOSE)

            itemlinepatt = """
            ([0-9]+)            #quantity and part of item description - capture group [0]
            [ ]+                #white space
            ([A-z ]+)           #part description [1]
            [\s]                #white space including new line
            ([0-9]+)            #order [2] - may need to trim leading digit? (length = 0)
            [\s]                #white space including new line
            ([0-9,.]+)          #price in european notation [3]
            [\s]                #white space including new line  
            ([0-9]+)            #Day [4]
            [/]                 #Date divider
            ([0-9]+)            #Month [5]
            -                   #Divider
            ([0-9]+)            #Year [6]
            """
            itemline = re.findall(itemlinepatt, PDFContents, re.VERBOSE)


            POnumlist = []
            for item in itemline:
                POnumlist.append(item[2][1:]) #trim leading digit (length 0)
                date = (item[4], item[5], item[6])
                itemtotal = item[3].replace('.', '').replace(',','.')

            PONumber = '/'.join(POnumlist) #join PO numbers to create customer reference

        #FRON01
        elif 'Frontier Technologies Brewton' in PDFContents:
                company = 'FRON01'
                OtherError = False

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

                    #check if individual line items sum correctly
                    sumofitems += float(itemtotal)

                    try:
                        if float(pricedictionary[(company, partnumber)]) != float(priceper):
                            errors.append([company, PONumber, originalPDF, "Incorrect item price.", "Got: " + str(priceper), "Expected: " + pricedictionary[(company, partnumber)]])
                            OtherError = True
                        elif int(quantity) * float(priceper) != float(itemtotal):
                            errors.append([company, PONumber, originalPDF, "Incorrect quantity or price for line item.", 'Calcd: ' + float(quantity) * float(priceper), 'PO: ' + POTotal])
                            OtherError = True
                    except:
                        nodictentry.append([company, partnumber, priceper])
                        errors.append([company, PONumber, "no price dictionary entry"])
                        OtherError = True

                # This check only comes at end of PO (end of for loop). The above errors can occur on individual line items, so we need to wait for this check.
                if float(sumofitems) != float(POTotal):
                    errors.append([company, PONumber, originalPDF, "Total price not sum of individual items.", 'Calcd: ' + sumofitems, 'PO: ' + POTotal])
                elif not OtherError:
                    ####output####
                    POContents.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])
                    processed = True

        else:
            company = 'unknown'
            errors.append([company, 'unknown', originalPDF, "PO not recognized"])

        #check if a folder has been made for a day - if not, create it
        if movePDF and processed:
            if not os.path.exists(datepath+processeddirectory):
                os.makedirs(datepath+processeddirectory)
            #move original file to new location
            os.rename(originalPDF, datepath + processeddirectory + PDFDate + '_' + company + '_' + str(PONumber) + '.pdf')
        elif not processed:
            logs.append(originalPDF)

    if len(nodictentry) > 0:
        # set output file name
        # including hours and minutes so that this program can be run twice in one day
        baddictfilename = datetime.datetime.now().strftime(datepath+'%y-%m-%d_%H%M_NoDictionaryEntry.csv')
        logs.append(baddictfilename)
        with open(baddictfilename, 'w', newline='') as out:
            itemwr = csv.writer(out, delimiter=",")
            for item in nodictentry:
                itemwr.writerow(item)


    # set output file name
    # including hours and minutes so that this program can be run twice in one day
    outputfilename = datetime.datetime.now().strftime(datepath+'%y-%m-%d_%H%M_SalesOrders.csv')
    logs.append(outputfilename)
    with open(outputfilename, 'w', newline='') as outfile:
        itemwriter = csv.writer(outfile, delimiter=",")
        for item in POContents:
            itemwriter.writerow(item)
    #print(POContents)

    if not len(errors):
        print('SOs successfully extracted!')
    else:
        print('There were some errors. Check the error log.')
        errorfilename = datetime.datetime.now().strftime(datepath+'%y-%m-%d_%H%M_ErrorLog.csv')
        logs.append(errorfilename)
        with open(errorfilename, 'w', newline='') as errfile:
            errwriter = csv.writer(errfile, delimiter=",")
            for item in errors:
                errwriter.writerow(item)

    return logs

if __name__ == "__main__":
    scrapePO(origindirectory='./pdf/')
