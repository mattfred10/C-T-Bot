from email_services import *

import PyPDF2
import re
import glob
import csv
import datetime
import os


class SOBOT:
    """A mail checking, sending and PDF scraping bot designed to input C&T customer purchase orders into EFACS as sales 
    orders. It should be deployed as a cron job and require no inputs. All activity will be shared via email. There are 
    many routines built in for trouble shooting and modification.    
    """

    def __init__(self, optionsfile='.\\SOBotSettings\\SOBotSettings.txt'):
        """Load options"""

        # Get date and subsequent path creation
        self.today = datetime.date.today()
        self.datestring = self.today.strftime("%Y-%m-%d")
        self.datepath = self.today.strftime('%Y\\%m\\%d\\')
        self.processedpath = self.datepath+'ProcessedSOs\\'
        self.unprocessedpath = self.datepath+'UnprocessedSOs\\'
        if not os.path.exists(self.processedpath):
            os.makedirs(self.processedpath)
        if not os.path.exists(self.unprocessedpath):
            os.makedirs(self.unprocessedpath)

        # self.POContents is the output. Create the header here.
        self.POContents = [['Company', 'PO Number', 'Due Date', 'Item Number', 'Quantity', 'Price Per Item', 'Order Total']]
        # Store list of output files for attachment to email - bad dict entries, error log, bad pdfs, and good SOs
        self.logs = []

        with open(optionsfile, 'r') as options:
            opt = ''.join(options.readlines())

            self.LEADTIME = int(re.search(r'LEADTIME[ =]+([\S]+)', opt)[1])

            self.TO = re.search(r'TO[ =]+([\S]+)', opt)[1]
            self.SUBJECT = re.search(r'SUBJECT[ =]+([A-z0-9 .,]+)', opt)[1] + self.datestring
            self.BODY = re.search(r'BODY[ =]+([A-z0-9 .,%]+)', opt)[1] % self.datestring

            self.PRICEDICTIONARY = re.search(r'PRICEDICTIONARY[ =]+([\S]+)', opt)[1]
            self.PODICTIONARY = re.search(r'PODICTIONARY[ =]+([\S]+)', opt)[1]
            self.BOTACCOUNT = re.search(r'BOTACCOUNT[ =]+([\S]+)', opt)[1]
            self.BOTPASSWORD = re.search(r'BOTPASSWORD[ =]+([\S]+)', opt)[1]
            self.FETCHMAILSERVER = re.search(r'FETCHMAILSERVER[ =]+([\S]+)', opt)[1]
            self.SENDPORT = re.search(r'SENDPORT[ =]+([\S]+)', opt)[1]
            self.SENDMAILSERVER = re.search(r'SENDMAILSERVER[ =]+([\S]+)', opt)[1]

        # debug options
        self.printstatus = False
        self.movepdf = True
        self.PDFtoText = False

        self.leaveunread = False # True = will leave messages unread False - mark as read
        self.checkPOdictionary = True
        self.checkdate = True

    def debug(self, movepdf=True, PDFtoText=False, leaveunread=False, checkPOdictionary=True, checkdate=True, originfolder='', destfolder='', outputpath=''):
        """Override some options for debug purposes. Default options are intended to be completely autonomous. Only interaction
        occurs via email. Calling this method without any options will turn on status updates but leave other behavior
        alone.
        
        N.B.: leaveunread=True occasionally causes an exception during mail checking.
        """
        if originfolder:
            self.unprocessedpath = originfolder
        if destfolder:
            self.processedpath = destfolder
        if outputpath:
            self.datepath = outputpath

        self.printstatus = True
        self.movepdf = movepdf
        self.PDFtoText = PDFtoText
        self.leaveunread = leaveunread  # True = will leave messages unread False - mark as read
        self.checkPOdictionary = checkPOdictionary
        self.checkdate = checkdate

        if self.printstatus:
            print('Debug mode')

    def BOTfetch(self):
        """Download attachments fetching functions. 
        
        The FetchMail class is specialized for this bot:
        1) Only downloads pdfs
        2) Filters and does not download invoices or 'Terms for Goods Services' files
        3) Appends -i (where i is simply iterated until it doesn't match) to files with the same name (common in this use).
        
        It could be generalized further if necessary.
        """

        fetcher = FetchMail(self.FETCHMAILSERVER, self.BOTACCOUNT, self.BOTPASSWORD, self.leaveunread, download_folder=self.unprocessedpath)
        if self.printstatus:
            print("Succesfully connected to imap server.")
        emails = fetcher.fetch_unread_messages()
        if self.printstatus:
            print("Succesfully fetched emails.")
        for entry in emails:
            fetcher.save_attachment(entry)
        if self.printstatus and emails:
            print("Succesfully saved attachents.")
        fetcher.close_connection()
        if self.printstatus:
            print("Succesfully received mail.")

    def BOTsend(self):
        """Send log files and unprocessed PDFs to the recipients in the settings file"""
        sender = SendMail(self.SENDMAILSERVER, self.SENDPORT, self.BOTACCOUNT, self.BOTPASSWORD)
        if self.printstatus:
            print("Succesfully connected to smtp server.")
        sender.composemsg(self.TO, self.SUBJECT, self.BODY, self.logs)
        if self.printstatus:
            print("Succesfully composed message.")
        sender.open_connection()
        if self.printstatus:
            print("Succesfully opened connection.")
        sender.send()
        if self.printstatus:
            print("Succesfully sent mail.")
        sender.close_connection()
        if self.printstatus:
            print("Succesfully closed connection.")

    def checkPOdict(self, PONumber, company, podictionary, checkPOdictionary=True):
        if checkPOdictionary:
            try:
                if not podictionary[PONumber] == company:
                    return False
            except:
                podictionary[PONumber] = company
                return True
        else:
            return True

    def checkDate(self, datetuple):
        """Scraper generates date tuples. This method checks that they make sense and converts words (e.g., jul, JUL, 
        Jul, July, JULY) to the same format (DD-MM-YY)"""

        errstatus = ''  #empty string to store errors. if it exists, throw error in scraper

        # Using a list here and going to convert using indicies. dictinoary might be better?
        monthlist = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']

        # monthdict = {'jan' : 31,
        #             'feb' : 28,
        #             'mar' : 31,
        #             'apr' : 30,
        #             'may' : 31,
        #             'jun' : 30,
        #             'jul' : 31,
        #             'aug' : 31,
        #             'sep' : 30,
        #             'oct' : 31,
        #             'nov' : 30,
        #             'dec' : 31}

        monthdict = {1: 31,
                     2: 28,
                     3: 31,
                     4: 30,
                     5: 31,
                     6: 30,
                     7: 31,
                     8: 31,
                     9: 30,
                     10: 31,
                     11: 30,
                     12: 31}

        day = int(datetuple[0])
        month = str(datetuple[1])
        year = int(datetuple[2])

        if year < 2000:
            year = year + 2000

        # Determine if the date is an abbreviation and convert it to a number
        if len(month) > 2:
            if month[0:3].lower() in monthlist:
                month = monthlist.index(month[0:3].lower()) + 1
            else:
                errstatus = 'Error in month name.'

        # Make sure our months are ints
        month = int(month)

        if not 1 <= month <= 12:
            errstatus = 'Month out of range.'
        elif monthdict[month] < day:  # only check this if the month is valid, otherwise there will be a key error
            errstatus = 'Day out of range.'
        elif datetime.date(year, month, day) < self.today:  # month and day both need to be valid or datetime throws an exception
            errstatus = 'Past due.'
        elif (datetime.date(year, month, day) - self.today) > datetime.timedelta(days=self.LEADTIME):  # month and day both need to be valid or datetime throws an exception
            errstatus = 'Past due.'

        return day, month, year, errstatus  # returns tuple of ints, can change this return or reverse process to produce month names



    def BOTscrape(self):
        """Opens PDF files in the download directory, determines their origin, and finds the important information. 
        
        Outputs between 1 and 3 files depending on whether there are errors and the type of error:
        1) an SO output file that contains the relevant information. Eventually should be added directly to EFACS
        2) an error log showing the errors and the files they originated from 
        3) a list of price dictionary entries that were not matched - this is only for key errors not for price 
        discrepancies, which is a different error.
        
        The routine then returns a list of file paths for the log files and the pdfs with errors in them. BOTsend()
        will take this list as an argument and attach the files to the report email.
        
        Currently will only do entire directory, but could be changed if need be.
        
        Be aware that a lot of this code is very similar, but the original data are different, so the error checking
        and other routines would be difficult to reuse.
        """
    
        # store a list of POs with bad entries and their errors.
        errors = []

        # generate dictionary of prices to check that POs are correct
        pricedictionary = {}
        # catch any unlisted company/part combinations in this list and print it at the end
        # if they are correct, they can be added to the price dictionary.
        # ensure that the prices are stored with 2 decimal places (even for e.g., $0.30 - the trailing zero is required)
        nodictentry = []
        with open(self.PRICEDICTIONARY) as dictfile:
            dictentries = csv.reader(dictfile)
            for line in dictentries:
                # use tuple of company (i.e., VEST01, etc) and item
                # companies have different prices
                pricedictionary[(line[0], line[1])] = line[2]

        # Similarly, create a dictionary of previous POs. Don't want duplicate entries
        podictionary = {}
        try:
            with open(self.PODICTIONARY) as podictfile:
                podictentries = csv.reader(podictfile)
                for line in podictentries:
                    # use tuple of company (i.e., VEST01, etc) and item
                    # companies have different prices
                    podictionary[line[0]] = line[1]
        except:
            pass  # should really only occur when dictionary file is empty

        # iterate over files in the folder
        for originalPDF in glob.glob(self.unprocessedpath + '*.pdf'):
            if self.printstatus:
                print(originalPDF)

            processed = False

            # create empty string for PDF contents
            PDFContents = ''
            # Some POs have multiple items that need to be iteratively processed - store temporarily to ensure that there are no errors before writing to SO output
            tempitems = []
    
            # open PDF
            pdfFileObj = open(originalPDF, 'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
            # combine all pages and concat into single string
            for page in range(0, pdfReader.numPages):
                pageObj = pdfReader.getPage(page)
                pagetext = pageObj.extractText()
                PDFContents = PDFContents + pagetext
    
            # Close the file. Use string from scraped PDFreader - needs to be closed so we can move the file at the end
            pdfFileObj.close()

            # write scrapped pdf text for testing purposes
            # open text file and write the PDFContents
            if self.PDFtoText:
                # create filename for scraped text
                textfile = self.processedpath + originalPDF.split('\\')[-1].split('.')[0] + '_scraped.txt'
                with open(textfile, 'w') as pdfout:
                    pdfout.write(PDFContents)
    
            # Can't effectively eliminate engineering drawing from file names
            # Check for them first because they likely have some of the same information.
            # Might produce false positives
            if 'Draw. format' in PDFContents: # this is for Siemans - no idea if others will appaer
                company = 'Siemans'
                errors.append([company, 'unknown', originalPDF, "Appears to be an enginnering drawing. Please double check the document."])
                # self.logs.append(originalPDF)
            elif 'estes-express' in PDFContents:
                company = 'Estes'
                errors.append([company, 'unknown', originalPDF, "Appears to be an invoice. Please double check the document."])
                # self.logs.append(originalPDF)
            # GE
            elif 'GE Renewables' in PDFContents:
                company = 'GEC01'
    
                PONumber = re.search(r'Order Number[\s]+([0-9]+)', PDFContents)[1]
                POTotal = re.search(r'Total Amt:([0-9.,]+)', PDFContents)[1]
    
                datequantity = re.findall(r'Delivery Schedule:([0-9]+)-([A-Z]{3})-([0-9]{2})([0-9,]+)[ ]+EACH', PDFContents)
                partnumbers = re.findall(r'GE Item: ([A-Z0-9]+)[\s]+Rev: ([0-9]+)', PDFContents) #excluding revision for now

                # here we are moving backwards from 'Hazard Code' because the text in front is unreliable (text behind may be too if they optionally include 'promis shipment date' or 'need by shipment date')
                # we capture everything including the day of month in case they change from 2 digits to 1 digit
                prices = re.findall(r'([0-9.]+)-[A-Z]{3}-[0-9 \n]+Hazard', PDFContents)
    
                sumofitems = 0
                # using delivery dates as a proxy for number of items in order
                for i,dq in enumerate(datequantity):
                    # first we will try to strip the day of month using the date grabbed elsewhere
                    # check if it's one digit or two
                    if prices[i][-1] == dq[0]:
                        # take all but the last number
                        combinedprice = prices[i][0:-1]
                    else:
                        # take all but the last two numbers
                        combinedprice = prices[i][0:-2]
    
                    date = self.checkDate((dq[0], dq[1], dq[2]))
                    quantity = dq[3]
    
                    # The price per and item total don't separate in the GE files. Use the quantity sold to find correct values
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
                    date = (date[0], date[1], date[2])
                    tempitems.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])
    
                try:
                    if float(pricedictionary[(company, partnumber)]) != float(priceper):
                        errors.append([company, PONumber, originalPDF, "Incorrect item price.", "got: " + str(priceper), "expected: " + pricedictionary[(company, partnumber)]])
                        # self.logs.append(originalPDF)
                    elif float(sumofitems) != float(POTotal):
                        errors.append([company, PONumber, originalPDF, "Incorrect total price or number of items.", 'Calcd: ' + str(sumofitems), 'PO: ' + POTotal])
                        # self.logs.append(originalPDF)
                    elif date[3] and self.checkdate:
                        errors.append([company, PONumber, originalPDF, "Problem with date.", date])
                    else:
                        ####output####
                        if self.checkPOdict(PONumber, company, podictionary, checkPOdictionary=self.checkPOdictionary):
                            date = (date[0], date[1], date[2]) # Check date above outputs tuple with 4 entries - remake as 3
                            self.POContents.extend(tempitems)
                            processed = True
                        else:
                            errors.append([company, PONumber, originalPDF,
                                           "File appears to be a duplicate of an already processed PO."])
                except:
                    nodictentry.append([company, partnumber, priceper])
                    errors.append([company, PONumber, originalPDF, "No price dictionary entry - see 'NoDictionaryEntry' file."])
                    # self.logs.append(originalPDF)
    
            # Vestas America (Vest01) - Handles some sjol01 - sjol01 also handled below (after all Vestas)
            elif 'Vestas Nacelles America' in PDFContents:
                # currently only handles 1 item per PO. Should be easy to fix if it comes up. Need to see an example first.
                if 'SJOELUND US INC.' in PDFContents:
                    # do something here about the different ways to handle different companies
                    # may need to move this to the end when data are written
                    # I think the only difference is inputting price in the SO, which will be handled during input.
                    company = 'sjol01'
                else:
                    company = 'VEST01'
    
                # Get PO Number
                PO = re.search(r'P(K|1)[0-9]{5}', PDFContents)  # Revision number and page number run right into this so it needs to be tight (i.e., 5 numbers)
                PONumber = PO[0]
                # 1200, 1240 appear to be formatting information
                # if there is more than one item, this is going to need to change
                item = re.search(r'    1200([0-9]{6,})', PDFContents)
                partnumber = item[1]
    
                # dates separated by '.'
                datepattern = """
                1240            #string that appears to be pdf formatting information and consistently leads this line
                ([0-9 ]{2})     #Day - strictly 2 digits because it runs into 1240. Other dates appear to always be 2 digits but are given flexibility in case format changes.
                \.              #escaped '.' which is used to divide the days, months and years
                ([0-9]+)        #Month
                \.              #escaped '.' which is used to divide the days, months and years
                ([0-9]+)        #Year
                """
                dategroup = re.search(datepattern, PDFContents, re.VERBOSE)
                # dategroup = re.search(r'1240([0-9 ]{2})\.([0-9]+)\.([0-9]+)', PDFContents)
    
                date = self.checkDate((dategroup[1], dategroup[2], dategroup[3]))
    
                pqpattern = """
                ([0-9,]+)       #quantities allowing for thousands (,)
                EA[ ]+          #EA is the units and the number of spaces varies by the length of the quantity
                ([0-9,]+)       #price per item - european notation with not thousands separator
                [ ]+            #variable number of spaces after unit price
                ([0-9,]+)       #order total (quanity * unit price)
                """
                pricesandquantity = re.search(pqpattern, PDFContents, re.VERBOSE)
                # pricesandquantity = re.search(r'([0-9,]+)EA[ ]+([0-9,]+)[ ]+([0-9,]+)', PDFContents)
    
                quantity = float(pricesandquantity[1].replace(',', '.'))
                priceper = float(pricesandquantity[2].replace(',','.'))
                POTotal = float(pricesandquantity[3].replace(',','.'))
    
                try:
                    if float(pricedictionary[(company, partnumber)]) != round(float(priceper),2):
                        errors.append([company, PONumber, originalPDF, "Incorrect item price.", "Got: " + str(priceper), "Expected: " + pricedictionary[(company, partnumber)]])
                        # self.logs.append(originalPDF)
                    elif quantity * priceper < POTotal - 1 or quantity * priceper > POTotal + 1:
                        errors.append([company, PONumber, originalPDF, "Incorrect total price or number of items.", 'Calcd: ' + float(quantity) * float(priceper), 'PO: ' + POTotal])
                        # self.logs.append(originalPDF)
                    elif date[3] and self.checkdate:
                        errors.append([company, PONumber, originalPDF, "Problem with date.", date])
                    else:
                        ####output####
                        if self.checkPOdict(PONumber, company, podictionary, checkPOdictionary=self.checkPOdictionary):
                            date = (date[0], date[1], date[2]) #Check date above outputs tuple with 4 entries - remake as 3
                            self.POContents.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])
                            processed=True
                        else:
                            errors.append([company, PONumber, originalPDF,
                                           "File appears to be a duplicate of an already processed PO."])
                except:
                    nodictentry.append([company, partnumber, priceper])
                    errors.append([company, PONumber, originalPDF, "No price dictionary entry - see 'NoDictionaryEntry' file."])
                    # self.logs.append(originalPDF)
    
            # Vest02 and Vest04 use very similar formats
            elif 'Vestas - American Wind Technology' in PDFContents:
                company = 'VEST02'

                otherError = False


                # Get PO Number
                PO = re.search(r'Purchase order ([0-9]+)', PDFContents)
                PONumber = PO[1]
    
                # multiple items per order
                # counts occurrences of 'Delivery date' as it appears once per item at the end of the item
                alldates = re.findall(r'Delivery date: ([0-9]{1,2}) ([A-z]{3}) ([0-9]{4})', PDFContents)
                # regex pattern 10,20,30, etc item line + (part number) + spaces + (quantity) + EA + spaces + (price per) + (total item price)

                itemlinepatt ="""
                [0-9]{1}0       #line items begin with 10, 20, 30, etc.
                ([0-9]+)        #part number - capture group [0]
                [ ]+            #variable spaces
                ([0-9]+)        #quantitiy - capture group [1]
                [ ]EA[ ]+       #single space + EA + variable white space
                ([0-9.,]+)      #unit price - capture group [2]
                [ ]+            #variable spaces
                ([0-9.,]+)      #item total - capture group [3]
                """
                itemline = re.findall(itemlinepatt, PDFContents, re.VERBOSE)
                # itemline = re.findall(r'[0-9]{1}0([0-9]+)[ ]+([0-9]+) EA[ ]+([0-9.,]+)[ ]+([0-9.,]+)', PDFContents)

                POTotal = re.findall(r'Net value[ ]+([0-9.,]+)', PDFContents)[0].replace(',', '')
    
                #Need to iterate over these item lists and assign values. Not sure how many are in each.
                #One delivery date per item is used as a proxy for total number of items
                #If items span pages (haven't seen it happen yet), this procedure will need to be adapted (maybe just concat all pages at beginning?)
                #First, we keep track of our total value by adding the line items and check it against the reported PO value at the end
                sumofitems = 0
    
                for i, date in enumerate(alldates):
                    # regex only checks line items 10-90. Probably won't be 10 items but throw error just in case
                    if i > 9:
                        otherError=True
                        errors.append([company, PONumber, originalPDF, "More than 9 line items. File not processed correctly."])
                        break

                    date = self.checkDate(date)

                    partnumber = itemline[i][0]
                    quantity = itemline[i][1].replace(',', '')
                    priceper = itemline[i][2].replace(',', '')
                    itemtotal = itemline[i][3].replace(',', '')

                    try:
                        if float(pricedictionary[(company, partnumber)]) != float(priceper):
                            errors.append([company, PONumber, originalPDF, "Incorrect item price.", "Got: " + str(priceper), "Expected: " + pricedictionary[(company, partnumber)]])
                            otherError = True
                            break
                        elif float(priceper) * int(quantity) > float(itemtotal) + 1 or float(priceper) * int(quantity) < float(itemtotal) - 1: #Set this as +/- 1 to deal with floating point precision errors
                            errors.append([company, PONumber, originalPDF, "Incorrect quantity or price for line item.", 'Calcd: ' + float(quantity) * float(priceper), 'PO: ' + POTotal])
                            otherError = True
                            break
                        elif date[3] and self.checkdate:
                            errors.append([company, PONumber, originalPDF, "Problem with date.", date])
                            break
                        else:
                            date = (date[0], date[1], date[2])  # Check date above outputs tuple with 4 entries - remake as 3
                            tempitems.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])
                    except:
                        nodictentry.append([company, partnumber, priceper])
                        errors.append([company, PONumber, originalPDF, "No price dictionary entry - see 'NoDictionaryEntry' file."])
                        otherError = True
                        break
    
                    sumofitems += float(itemtotal)
    
                # Check comes at end of PO. Above errors occur on line items, this check is for total PO price.
                if float(sumofitems) != float(POTotal):
                    errors.append([company, PONumber, originalPDF, "Total price not sum of individual items. This error may appear due to a different error in the price checking.", 'Calcd: ' + str(sumofitems), 'PO: ' + POTotal])
                elif not otherError: # actual error stored above
                    #####output#####
                    if self.checkPOdict(PONumber, company, podictionary, checkPOdictionary=self.checkPOdictionary):
                        self.POContents.extend(tempitems)
                        processed = True
                    else:
                        errors.append([company, PONumber, originalPDF, "File appears to be a duplicate of an already processed PO."])
    
            # Vest04 - very similar to vest02 but there are some spacing issues that are different
            elif 'Vestas Do Brasil Energia' in PDFContents:
                company = 'vest04'
                otherError: False
    
                # Get PO Number
                PO = re.search(r'Purchase order ([0-9]+)', PDFContents)
                PONumber = PO[1]
    
                #multiple items per order
                alldates = re.findall(r'Delivery date: ([0-9]{1,2}) ([A-z]{3}) ([0-9]{4})', PDFContents)

                # regex pattern = not space to avoid matching 201X from the date + 10,20,30, etc line item number + (part number/quantity together) + ' EA' + (price per)(total item price)
                alldatapatt = """
                [^ ][1-9]0          #date, line number and part number and quantity collide forming, e.g., 2017|10|4003452|45, want to match the line item (10,20) but not 20 in 2017, which has a leading space
                ([0-9,]+)           #part number and quantity - capture group [0]
                [ ]EA               #units
                ([0-9.,]+)          #unit price and total price - capture group [1]
                """
                alldata = re.findall(alldatapatt, PDFContents, re.VERBOSE)
                # alldata = re.findall(r'[^ ][1-9]0([0-9,]+) EA([0-9.,]+)', PDFContents)
    
                POTotal = re.search(r'Net value[ ]+([0-9.,]+)', PDFContents)[1].replace(',', '')
    
                sumofitems = 0
                for i, date in enumerate(alldates):
                    # regex only checks line items 10-90. Probably won't be 10 items but throw error just in case
                    if i > 9:
                        otherError=True
                        errors.append([company, PONumber, originalPDF, "More than 9 line items. File not processed correctly."])
                        # self.logs.append(originalPDF)
                        break

                    date = self.checkDate(date)

                    # Separate the combined terms
                    priceper = re.match(r'[0-9]+.[0-9]{2}', alldata[i][1])[0]
                    itemtotal = float(alldata[i][1].replace(priceper, '').replace(',', ''))
                    # not a great way to separate the quantity from the part number
                    # comes as '290107241,000' with quantity as 1,000 or '153452600' with quantity as 600
                    # need to calculate from the item total and the priceper (both of which we have at high confidence)
                    # Should double check with price dictionary
                    # this is dangerous though because of rounding errors.
                    quantity = float(itemtotal) / float(priceper.replace(',', ''))
                    # insert commas so that we match the quantity and not another part of the item string
                    commaquantity = format(int(quantity), ',d')
                    partnumber = alldata[i][0].replace(commaquantity, '')
    
                    # check if individual line items sum correctly
                    sumofitems += float(itemtotal)

                    try:
                        if float(pricedictionary[(company, partnumber)]) != float(priceper):
                            errors.append([company, PONumber, originalPDF, "Incorrect item price.", "Got: " + str(priceper), "Expected: " + pricedictionary[(company, partnumber)]])
                            # self.logs.append(originalPDF)
                            otherError = True
                            break
                        elif float(priceper) * int(quantity) > float(itemtotal) + 1 or float(priceper) * int(quantity) < float(itemtotal) - 1:  # Set this as +/- 1 to deal with floating point precision errors
                            errors.append([company, PONumber, originalPDF, "Incorrect quantity or price for line item.",'Calcd: ' + float(quantity) * float(priceper), 'PO: ' + POTotal])
                            # self.logs.append(originalPDF)
                            otherError = True
                            break
                        elif date[3] and self.checkdate:
                            errors.append([company, PONumber, originalPDF, "Problem with date.", date])
                            break
                        else:
                            date = (date[0], date[1], date[2])  # Check date above outputs tuple with 4 entries - remake as 3
                            tempitems.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])
                    except:
                        nodictentry.append([company, partnumber, priceper])
                        errors.append([company, PONumber, originalPDF, "No price dictionary entry - see 'NoDictionaryEntry' file."])
                        otherError = True
                        break
    
                # This check only comes at end of PO (end of for loop). The above errors can occur on individual line items, so we need to wait for this check.
                if float(sumofitems) != float(POTotal):
                    errors.append([company, PONumber, originalPDF, "Total price not sum of individual items.",
                                   'Calcd: ' + str(sumofitems), 'PO: ' + POTotal])
                elif not otherError:  # actual error stored above
                    #####output#####
                    if self.checkPOdict(PONumber, company, podictionary, checkPOdictionary=self.checkPOdictionary):
                        self.POContents.extend(tempitems)
                        processed = True
                    else:
                        errors.append([company, PONumber, originalPDF,
                                       "File appears to be a duplicate of an already processed PO."])
    
            # sjol01 - direct sales - some other Sjoeland POs use Vestas POs and are handled above
            # elif 'SJØLUND' in PDFContents:
            #     company = 'sjol01'
            #
            #     #requisition number not used - line items each get an order number - PO becomes item1/item2 - see below
            #     # POpatt = """
            #     # REQUISITION         #requisition number identifier
            #     # [\s]+               #White space including new line
            #     # ([0-9]+)            #requisition number
            #     # """
            #     #
            #     # PONumber = re.search(POpatt, PDFContents, re.VERBOSE)
            #
            #     itemlinepatt = """
            #     ([0-9]+)            #quantity and part of item description - capture group [0]
            #     [ ]+                #white space
            #     ([A-z ]+)           #part description [1]
            #     [\s]                #white space including new line
            #     ([0-9]+)            #order [2] - may need to trim leading digit? (length = 0)
            #     [\s]                #white space including new line
            #     ([0-9,.]+)          #price in european notation [3]
            #     [\s]                #white space including new line
            #     ([0-9]+)            #Day [4]
            #     [/]                 #Date divider
            #     ([0-9]+)            #Month [5]
            #     -                   #Divider
            #     ([0-9]+)            #Year [6]
            #     """
            #     itemline = re.findall(itemlinepatt, PDFContents, re.VERBOSE)
            #
            #
            #     POnumlist = []
            #     for item in itemline:
            #         POnumlist.append(item[2][1:]) #trim leading digit (length 0)
            #         date = (item[4], item[5], item[6])
            #         itemtotal = item[3].replace('.', '').replace(',','.')
            #
            #     PONumber = '/'.join(POnumlist) #join PO numbers to create customer reference
    
            # FRON01
            elif 'Frontier Technologies Brewton' in PDFContents:
                    company = 'FRON01'
                    otherError = False
    
                    # PO and date are combined
                    PONumber = re.search(r'America([0-9]+)[0-9]{2}-', PDFContents)[1]
                    POTotal = re.search(r'Total:\$([0-9.,]+)', PDFContents)[1].replace(',','')
                    # need to sum individual items
                    sumofitems = 0
    
                    #Regex = ea + $(priceper) + $(total) + Due: (date) + (partnumber) ea$5.12 $3,584.00 Due:04-Aug-17105W1931P016700H - quantity comes between last letter (=revision) and P### (part of PO)
                    itemlinepatt = """
                    ea\$                                    #end of units + literal $
                    ([0-9,.]+)                              #unit price in US notation - capture group [0]
                    [ ]\$                                   #space plus $ literal
                    ([0-9,.]+)                              #item total - capture group [1]
                    [ ]Due:                                 #text indicated due date
                    ([0-9]{2})-([A-z]{3})-([0-9]{2})        #date CGs [2-4] day, mon, year
                    ([A-Z0-9]+)                             #part number
                    (P[0-9]{3})                             #end of part number - use to guarantee quantity separation
                    ([0-9,]+)                               #quantity
                    ([A-Z])Rev                              #revision
                    """
                    itemline = re.findall(itemlinepatt, PDFContents, re.VERBOSE)
                    # itemline = re.findall(r'ea\$([0-9,.]+) \$([0-9,.]+) Due:([A-Z0-9]+)(P[0-9]{3})([0-9,]+)([A-Z])Rev', PDFContents)
    
                    for item in itemline:
                        priceper = item[0]
    
                        itemtotal = item[1].replace(',', '')
                        date = self.checkDate((item[2], item[3], item[4]))
                        partnumber = item[5] + item[6]
                        quantity = item[7].replace(',', '')
                        rev = item[8] # Does revision need to be included in part number?
    
                        # check if individual line items sum correctly
                        sumofitems += float(itemtotal)
    
                        try:
                            if float(pricedictionary[(company, partnumber)]) != float(priceper):
                                errors.append([company, PONumber, originalPDF, "Incorrect item price.", "Got: " + str(priceper), "Expected: " + pricedictionary[(company, partnumber)]])
                                otherError = True
                                break
                            elif int(quantity) * float(priceper) != float(itemtotal):
                                errors.append([company, PONumber, originalPDF, "Incorrect quantity or price for line item.", 'Calcd: ' + float(quantity) * float(priceper), 'PO: ' + POTotal])
                                otherError = True
                                break
                            elif date[3] and self.checkdate:
                                errors.append([company, PONumber, originalPDF, "Problem with date.", date])
                                break
                            else:
                                date = (date[0], date[1],
                                        date[2])  # Check date above outputs tuple with 4 entries - remake as 3
                                tempitems.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])
                        except:
                            nodictentry.append([company, partnumber, priceper])
                            errors.append([company, PONumber, originalPDF, "No price dictionary entry - see 'NoDictionaryEntry' file."])
                            otherError = True
                            break
    
                    # This check only comes at end of PO (end of for loop). The above errors can occur on individual line items, so we need to wait for this check.
                    if float(sumofitems) != float(POTotal):
                        errors.append([company, PONumber, originalPDF, "Total price not sum of individual items.", 'Calcd: ' + str(sumofitems), 'PO: ' + POTotal])
                    elif not otherError:
                        ####output####
                        if self.checkPOdict(PONumber, company, podictionary, checkPOdictionary=self.checkPOdictionary):
                            self.POContents.extend(tempitems)
                            processed=True
                        else:
                            errors.append([company, PONumber, originalPDF,
                                           "File appears to be a duplicate of an already processed PO."])

            # Unidentified PO
            else:
                company = 'unknown'
                errors.append([company, 'unknown', originalPDF, "PO not recognized"])
                #self.logs.append(originalPDF)
    
            # check if a folder has been made for a day - if not, create it
            if self.movepdf and processed:
                try:
                    os.rename(originalPDF, self.processedpath + self.datestring + '_' + company + '_' + str(PONumber) + '.pdf')
                except:
                    # In general, this message should not appear. Duplicates should be blocked by the PONumber dictionary.
                    errors.append([company, PONumber, originalPDF, "This file is a duplicate of an already processed file. It was not moved. There should be another associated error in the error log."])
            elif not processed:
                self.logs.append(originalPDF)


        if len(nodictentry) > 0:
            # set output file name
            # including hours and minutes so that this program can be run twice in one day
            baddictfilename = datetime.datetime.now().strftime(self.datepath+'%y-%m-%d_%H%M_NoDictionaryEntry.csv')
            self.logs.append(baddictfilename)
            with open(baddictfilename, 'w', newline='') as out:
                itemwr = csv.writer(out, delimiter=",")
                for item in nodictentry:
                    itemwr.writerow(item)

        # set output file name
        # including hours and minutes so that this program can be run twice in one day
        outputfilename = datetime.datetime.now().strftime(self.datepath+'%y-%m-%d_%H%M_SalesOrders.csv')
        self.logs.append(outputfilename)
        with open(outputfilename, 'w', newline='') as outfile:
            itemwriter = csv.writer(outfile, delimiter=",")
            for item in self.POContents:
                itemwriter.writerow(item)

        numerrors = len(errors)
        if not numerrors:
            if self.printstatus:
                print('SOs successfully extracted!')
        else:
            if self.printstatus:
                print('There were %s errors. Check the error log.' % numerrors)
            errorfilename = datetime.datetime.now().strftime(self.datepath+'%y-%m-%d_%H%M_ErrorLog.csv')
            self.logs.append(errorfilename)
            with open(errorfilename, 'w', newline='') as errfile:
                errwriter = csv.writer(errfile, delimiter=",")
                for item in errors:
                    errwriter.writerow(item)

        self.logs = self.logs

        #save the dictionary
        with open(self.PODICTIONARY, 'w') as foundPOs:
            for key, value in podictionary.items():
                foundPOs.write('%s,%s\n' % (key, value))

if __name__ == "__main__":
    bot = SOBOT()
    bot.debug(checkPOdictionary=False, checkdate=False)  # leaveunread=True checkPOdictionary=False
    bot.BOTfetch()
    bot.BOTscrape()
    bot.BOTsend()