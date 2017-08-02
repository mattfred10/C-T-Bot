from CTemail.email_services import *
from CTcsv.csvfunctions import *
from month_dictionaries import *

import PyPDF2
import re
import glob
import datetime
import os
import xlrd
import xlwt
import itertools
import sys, traceback

class SOBOT:
    """A mail checking, sending and PDF scraping bot designed to input C&T customer purchase orders into EFACS as sales 
    orders. It should be deployed as a cron job and require no inputs. All activity will be shared via email. There are 
    many routines built in for trouble shooting and modification.    
    """


    def __init__(self, optionsfile='.\\SOBotSettings\\SOBotSettings.txt'):
        """Load options"""

        # Get date and subsequent path creation
        # Using windows file paths throughout. Glob kept giving errors with mixed filepaths.
        self.today = datetime.date.today()
        self.datestring = self.today.strftime("%Y-%m-%d")
        self.datepath = self.today.strftime('%Y\\%m\\%d\\')
        self.processedpath = self.datepath + 'ProcessedSOs\\'
        self.unprocessedpath = self.datepath + 'UnprocessedSOs\\'
        if not os.path.exists(self.processedpath):
            os.makedirs(self.processedpath)
        if not os.path.exists(self.unprocessedpath):
            os.makedirs(self.unprocessedpath)

        with open(optionsfile, 'r') as options:
            opt = ''.join(options.readlines())

            self.MAXLEADTIME = int(re.search(r'MAXLEADTIME[ ]*=[ ]*([\S]+)', opt)[1])  # including [ ]*=[ ]* so that users have some foregiveness in the settings file
            # Minimum leadtime uses a dictionary
            #self.MINLEADTIME = int(re.search(r'MINLEADTIME[ =]+([\S]+)', opt)[1])

            self.TO = re.search(r'TO[ ]*=[ ]*([\S]+)', opt)[1]
            self.SUBJECT = re.search(r'SUBJECT[ ]*=[ ]*([A-z0-9 .,]+)', opt)[1] + self.datestring
            self.BODY = re.search(r'BODY[ ]*=[ ]*([A-z0-9 .,%]+)', opt)[1] % self.datestring

            self.BOTACCOUNT = re.search(r'BOTACCOUNT[ ]*=[ ]*([\S]+)', opt)[1]
            self.BOTPASSWORD = re.search(r'BOTPASSWORD[ ]*=[ ]*([\S]+)', opt)[1]
            self.FETCHMAILSERVER = re.search(r'FETCHMAILSERVER[ ]*=[ ]*([\S]+)', opt)[1]
            self.SENDPORT = re.search(r'SENDPORT[ ]*=[ ]*([\S]+)', opt)[1]
            self.SENDMAILSERVER = re.search(r'SENDMAILSERVER[ ]*=[ ]*([\S]+)', opt)[1]

            LEADTIMEDICTIONARYPATH = re.search(r'LEADTIMEDICTIONARYPATH[ ]*=[ ]*([\S]+)', opt)[1]  # chosen by user
            PRICEDICTIONARYPATH = re.search(r'PRICEDICTIONARYPATH[ ]*=[ ]*([\S]+)', opt)[1]  # needs to be checked so not auto written
            self.PODICTIONARYPATH = re.search(r'PODICTIONARYPATH[ ]*=[ ]*([\S]+)', opt)[1]  # used to write file at end.
            QUANTITYDICTPATH = re.search(r'QUANTITYDICTPATH[ ]*=[ ]*([\S]+)', opt)[1]

            NUMVESTASTURBINES = re.search(r'NUMVESTTURBINES[ ]*=[ ]*([\S]+)', opt)[1]
            VESTASPARTFORECASTPATH = re.search(r'VESTASPARTFORECASTPATH[ ]*=[ ]*([\S]+)', opt)[1]

            MFPARTSPATH = re.search(r'MFPARTSPATH[ ]*=[ ]*([\S]+)', opt)[1]

            self.projectedstock = re.search(r'PROJECTEDSTOCK[ ]*=[ ]*([\S]+)', opt)[1]
            self.stockprojectionpath =  re.search(r'STOCKPROJECTIONPATH[ ]*=[ ]*([\S]+)', opt)[1]

        # Create output and logging lists
        # self.POContents is the so output. Create the header here.
        self.POContents = [['Company', 'PO Number', 'Due Date', 'Item Number', 'Quantity', 'Price Per Item', 'Order Total']]
        # GRN is the invoice output
        self.GRN = [['Supplier', 'Date', 'Invoice Number', 'PO Number', 'Item Number', 'Quantity']]
        # Store a list of POs with bad entries and their self.errors.
        self.errors = []
        # Specific list of unaccounted for company/part/price pairs
        self.nopricedictentry = []
        # Store unaccounted for quantity dictionary entries
        self.noquantitydictentry = []
        # Store list of output files for attachment to email - bad dict entries, error log, bad pdfs, and good SOs
        self.logs = []
        # Store list of open orders for checking with StockPredictor
        self.VESTASJOopenorders = []
        self.HYopenorders = []
        self.GEopenorders = []

        # Create dictionaries and lists used throughout
        # Month/date count dictionsries (found this easier than the built in Python calendar funciton
        self.abr2days = abr2days()
        self.abr2num = abr2num()
        self.num2days = num2days()
        self.num2abr = num2abr()
        # Dictionary of minimum lead times for POs key: company, value: time in days
        self.leadtimedictionary = readCSVtodictionary(LEADTIMEDICTIONARYPATH)
        # Ensure that PO item prices match correct values
        self.pricedictionary = readCSVto2tupledictionary(PRICEDICTIONARYPATH)
        # Create dictionary of part number : box quantities
        self.quantitydict = readCSVtodictionary(QUANTITYDICTPATH)
        # Import list of manufactured parts and the component parts
        self.mfparts = readCSVtolist(MFPARTSPATH)

        # Similarly, create a list of previous POs. Don't want duplicate entries in EFACS
        # This seems better than just checking if a tuple exists for the pair
        # Can't use PONum:Company in case companies have the same PONumber, so need to use tuple
        # This should be replaced with SQL in the future so we get POs from the db
        self.polist = []
        with open(self.PODICTIONARYPATH) as podictfile:
            podictentries = csv.reader(podictfile)
            for item in podictentries:
                self.polist.append((item[0], item[1]))

        # Create dictionary of vestas part forecasts
        # Need to multiply by variable number of turbines per week
        self.vestasforecast = {}
        with open(VESTASPARTFORECASTPATH) as forecast:
            forecastentries = csv.reader(forecast)
            for item in forecastentries:
                self.vestasforecast[item[0]] = int(item[1]) * int(NUMVESTASTURBINES)  # Forecasted number of parts

        # Create Excel workbook for output
        self.book = xlwt.Workbook()  # Create a workbook

        # debug options
        self.printstatus = False
        self.movepdf = True
        self.PDFtoText = False

        self.leaveunread = False # True = will leave messages unread False - mark as read
        self.POdictionarycheck = True
        self.datecheck = True

    def debug(self, movepdf=True, PDFtoText=False, leaveunread=False, POdictionarycheck=True, datecheck=True, originfolder='', destfolder='', outputpath=''):
        """Override some options for debug purposes. Default options are intended to be completely autonomous. Only interaction
        occurs via email. Calling this method without any options will turn on status updates but leave other behavior
        alone.

        N.B.: leaveunread=True occasionally causes an exception during mail checking. I haven't been able to track down the cause.
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
        self.POdictionarycheck = POdictionarycheck
        self.datecheck = datecheck

        if self.printstatus:
            print('Debug mode')

    def dateTupleToDatetime(self, datetuple):
        """
        Hopefully, I can replace this soon when I figure out how to store date for EFACS entry.
        """
        # Covers two cases found here
        datetuple = str(datetuple)
        if '-' in datetuple:
            date = datetuple.split('-')
        else:
            date = datetuple.replace('(','').replace(')','').split(',')

        return datetime.date(int(date[0]),int(date[1]),int(date[2]))

    def get_sheet_by_name(self, name):
        """
        Get sheet from xlwt object by name
        """
        try:
            for i in itertools.count():
                s = self.book.get_sheet(i)
                if s.name == name:
                    return s
        except IndexError:
            raise # We catch this later in case their are case issues (vesta123 vs VESTAS123)

    def fetchMail(self):
        """Download attachments fetching functions.

        The FetchMail class is specialized for this bot:
        1) Only downloads pdf and xls
        2) Filters and does not download invoices or 'Terms for Goods Services' files or some scanned files
        3) Appends -i (where i is simply iterated until it doesn't match) to files with the same name (common in this use).

        It could be generalized further if necessary.
        """

        fetcher = FetchMail(self.FETCHMAILSERVER, self.BOTACCOUNT, self.BOTPASSWORD, self.leaveunread, download_folder=self.unprocessedpath)
        if self.printstatus:
            print("Successfully connected to imap server.")
        emails = fetcher.fetch_unread_messages()
        numattachments = len(emails)
        if self.printstatus:
            print("Successfully fetched %s emails." % numattachments)  # probably want to log this message
        if numattachments:
            skipped = []
            for entry in emails:
                skipped.extend(fetcher.save_attachment(entry))
            if self.printstatus:
                print(skipped)
            if self.printstatus and emails:
                print("Successfully saved attachments.")
        fetcher.close_connection()
        if self.printstatus:
            print("Successfully received mail.")

        return numattachments

    def sendMail(self):
        """Send log files and unprocessed PDFs to the recipients in the settings file.
        """

        # We split '/services' off the bot account and just send as accounts1@ for now.
        # This method should be safe if it's every changed to a real email address
        sender = SendMail(self.SENDMAILSERVER, self.SENDPORT, self.BOTACCOUNT.split('/')[0], self.BOTPASSWORD)
        if self.printstatus:
            print("Successfully connected to smtp server.")
        sender.composemsg(self.TO, self.SUBJECT, self.BODY, self.logs)
        if self.printstatus:
            print("Successfully composed message.")
        sender.open_connection()
        if self.printstatus:
            print("Successfully opened connection.")
        try:
            sender.send()
            if self.printstatus:
                print("Successfully sent mail.")
            sender.close_connection()
            if self.printstatus:
                print("Successfully closed connection.")
        except:
            print('Send Mail Timeout')
            traceback.print_exc(file=sys.stdout)
            errorReporter = SendMail(self.SENDMAILSERVER, self.SENDPORT, self.BOTACCOUNT.split('/')[0], self.BOTPASSWORD)
            errorReporter.composemsg(self.TO, self.SUBJECT, "Email server timed out, likely because there were many abnormal files to upload. Process should have completed up to report generation. Please check the server for the files.")
            errorReporter.open_connection()
            errorReporter.send()
            errorReporter.close_connection()

    # Consider a set instead of a list (hash so faster)
    def checkPOdictionary(self, PONumber, company):
        """Checks the PO dictionary and adds any missing entries"""
        if self.POdictionarycheck:
            if (PONumber,company) in self.polist:
                return False
            else:
                self.polist.append((PONumber, company))
                return True
        else:
            return True

    def checkDate(self, datetuple, company):
        """Scraper generates date tuples. This method checks that they make sense and converts words (e.g., jul, JUL,
        Jul, July, JULY) to the same format (DD-MM-YY). Dates should be collected in the correct order by the
        scraper."""

        errstatus = ''  # empty string to store self.errors. if it exists, throw error in scraper

        day = int(datetuple[0])
        month = str(datetuple[1])
        year = int(datetuple[2])

        if year < 100:
            year = year + 2000

        # Determine if the date is an abbreviation and convert it to a number
        try:
            if len(month) > 2:
                month = self.abr2num[month[0:3].lower()]
        except:
            errstatus = 'Error in month name.'

        # Make sure our months are ints
        month = int(month)

        if not 1 <= month <= 12:
            errstatus = 'Month out of range.'
        elif self.num2days[month] < day and month != 2:  # only check this if the month is valid, otherwise there will be a key error
            errstatus = 'Day out of range.'
        elif month == 2 and year % 4 != 0 and self.num2days[month] < day:  # exception for leap year
            errstatus = 'Day out of range.'
        elif month == 2 and self.num2days[month] + 1 < day:
            errstatus = 'Day out of range.'
        elif datetime.date(year, month, day) < self.today + datetime.timedelta(days=int(
                self.leadtimedictionary[
                    company])):  # month and day both need to be valid or datetime throws an exception
            errstatus = 'Due date is earlier than allowed range.'
        elif (datetime.date(year, month, day) - self.today) > datetime.timedelta(
                days=self.MAXLEADTIME):  # month and day both need to be valid or datetime throws an exception
            errstatus = 'Due date more than ' + self.MAXLEADTIME + ' days in the future.'

        return day, month, year, errstatus  # consider refactoring to return 3 tuple if good and 4 tuple if bad - rather than check if date[3] check len(date)
        # There are many places below where we do (date[0], date[1], date[2]) would be a lot cleaner

    def checkQuantities(self, partnumber, quantity):
        try:
            if quantity % float(self.quantitydict[partnumber]) != 0:
                return "The box quantity is incorrect for part number " + partnumber
        except KeyError:
            self.noquantitydictentry.append([partnumber, quantity])
            return None # Just keeping track for now

        return None

    def checkPriceDictionary(self, company, partnumber, priceper):
        try:
            if float(self.pricedictionary[(company, partnumber)]) != float(priceper):
                return 'Price does not match price dictionary entry., Got: ' + str(priceper) + ', Expected: ' + self.pricedictionary[(company, partnumber)]
        except KeyError:
            self.nopricedictentry.append([company, partnumber, priceper])
            return 'No entry for company/part number. See attached NoPriceDictionaryEntry file.'

        return None

    def parseExcel(self):
        """
        Hyster Yale POs come in excel format. Collect data from these excel files.
        Files contain open orders as well.
        """

        for badxl in glob.glob(self.unprocessedpath + '*.xls'):  # Check if unknown file
            print("ERROR: Unexpected excel 2003 file: " + badxl)
            self.errors.append(['unknown', 'Many', badxl, "Unexpected excel file type."])
            self.logs.append(badxl)

        for excelFile in glob.glob(self.unprocessedpath + '*.xlsx'):  # Should be xlsx
            if self.printstatus:
                print(excelFile)
            processed = False

            s = xlrd.open_workbook(excelFile).sheet_by_index(0)  # Only need first sheet
            data = []

            for row in range(s.nrows):
                values = []
                for col in range(s.ncols):
                    values.append(str(s.cell(row, col).value))
                data.append(values)

            # This code should be sufficient to prevent index error below
            if not len(data) or len(data[0]) < 2:  # Throws error on empty excel file and moves on to next file.
                self.errors.append(['unknown', 'Many', excelFile, "Empty excel file."])
                continue

            if 'Report Generated' in data[0][0]:
                company = 'HYST01'
                for row in data[1:]:
                    PONumber = row[1][:10]
                    tempdate = xlrd.xldate_as_tuple(float(row[8]), 0)
                    date = self.checkDate((tempdate[2], tempdate[1], tempdate[0]), company)

                    partnumber = row[2]
                    quantity = float(row[9])

                    qtychk = self.checkQuantities(partnumber, quantity)
                    #
                    # Excel and python both like to trim leading 0s from the part numbers (casting as int)
                    # Casting as string when reading the csv and writing to the dict didn't help either
                    # Manually adding check here to improve usability
                    #
                    # Some serious EAFP programming here, but I have no idea how these lists will be altered in the future.
                    # 
                    # Consider moving this into checkprices funnction
                    #
                    try:
                        priceper =  float(self.pricedictionary[(company, partnumber)])
                    except KeyError:
                        try:
                            partnumber = '0' + partnumber
                            priceper = float(self.pricedictionary[(company, partnumber)])
                        except KeyError:
                            try:
                                partnumber = partnumber.lstrip('0')
                                priceper = float(self.pricedictionary[(company, partnumber)])
                            except KeyError:
                                self.errors.append([company, PONumber, excelFile, "The indicated PO has no price for part number: " + partnumber])
                    POTotal = priceper * quantity
                    
                    # Date error gets thrown for both open and new orders
                    if date[3]:
                        self.errors.append([company, PONumber, excelFile, "The indicated PO has a date issue.", date])
                    #However, we only check quantity on new orders
                    elif self.checkPOdictionary(row[1], company):  # Need to use entire PO number because line items take POnum001, POnum002, etc. and don't want false matches
                        if qtychk:
                            self.errors.append([company, PONumber, excelFile, qtychk])
                        else:
                            processed = True
                            self.POContents.append([company, PONumber, (date[2], date[1], date[0]), partnumber, quantity, priceper, POTotal])
                    else:
                        self.HYopenorders.append([partnumber, PONumber, (date[2],date[1],date[0]), quantity])

                # check if a folder has been made for a day - if not, create it
                if self.movepdf and processed:
                    try:
                        os.rename(excelFile, self.processedpath + self.datestring + '_' + company + '_OO.xlsx')
                    except:
                        # In general, this message should not appear. Duplicates should be blocked by PONumber dictionary
                        # Put here to avoid crashes during debugging (don't check PO dictionary flag)
                        self.errors.append([company, 'Many', excelFile, "This file is a duplicate of an already processed file. It was not moved."])
                elif not processed:
                    self.logs.append(excelFile)

            #SJOL
            elif 'Primary Vendor' in data[0][0]:
                company = 'SJOL'
                self.errors.append([company, 'Many', excelFile, "This appears to be a Sjoelund order."])

            #SJOL-VEST
            elif 'Vendor' in data[0][0]:
                company = 'SJOL-VEST'
                for row in data[1:]:
                    # Vestas has a habit of leaving information out of their OO reports
                    # Check for any missing information and throw an error.
                    if not row[3] or not row[2] or not row[8] or not row[7]:
                        self.errors.append([company, row[2], excelFile, "Missing information in SJOL-VEST OO report."])
                        # Don't add this information to the OO data or else you will get an error in the stock projection.
                        continue
                    # Also include random white space in these cells
                    # Throws a value error string to float tabs/spaces
                    try:
                        tempdate = xlrd.xldate_as_tuple(float(row[8]), 0)
                    except ValueError:
                        self.errors.append([company, row[2], excelFile, "Missing information/incorrect data in SJOL-VEST OO report."])
                        continue

                    # [part, po, date, quantity]
                    self.VESTASJOopenorders.append([row[3], row[2], (tempdate[0],tempdate[1],tempdate[2]), row[7]])
                try:
                    os.rename(excelFile, self.processedpath + self.datestring + '_' + company + '_OO.xlsx')
                except FileExistsError:
                    self.errors.append([company, 'Many', excelFile, "This appears to be a duplicate open order file."])

            #GE and Vestas(alt)
            elif 'Order' in data[0][0]:
                if len(data[0]) > 15:
                    company = 'GE'
                    for row in data[1:]:
                        try:  # GE empties this cell sometimes. It's shown as various length whitespace and empty. Just catch valueerror
                            tempdate = xlrd.xldate_as_tuple(float(row[14]), 0)
                            self.GEopenorders.append([row[9], row[0], (tempdate[0], tempdate[1], tempdate[2]), row[4]])
                        except ValueError:
                            self.errors.append(
                                [company, row[0], excelFile, "Past due or date error."])
                    try:
                        os.rename(excelFile, self.processedpath + self.datestring + '_' + company + '_OO.xlsx')
                    except:
                        self.errors.append([company, 'Many', excelFile, "This appears to be a duplicate open order file."])
                else:
                    company = 'Vestas'
                    self.errors.append([company, 'Many', excelFile, "Vestas changed OO format."])

            #C&T GRN
            elif 'Ningbo' in data[0][1]:
                company = 'coop01'
                invonum = data[6][5].replace('INV. NO:','')
                date = self.today - datetime.timedelta(days=1)  # date from inovice unreliable - usually arrives in the evening
                for row in data[12:]:  # Unknown number of rows. Iterate and match regex
                    if re.match(r'P/[0-9]+', row[0]):
                        # [['Supplier', 'Date', 'Invoice Number', 'PO Number', 'Item Number', 'Quantity']]
                        self.GRN.append([company, date, invonum, row[0], row[2], row[4]])
                try:
                    os.rename(excelFile, self.processedpath + self.datestring + '_' + company + '_GRN.xlsx')
                except:
                    self.errors.append([company, 'Many', excelFile, "This file appears to be a duplicate GRN invoice."])
            else:
                self.errors.append(['Unknown', 'unknown', excelFile, "Unknown excel file."])

    def scrapePDF(self):
        """Opens PDF files in the download directory, determines their origin, and finds the important information.

        Outputs between 1 and 3 files depending on whether there are self.errors and the type of error:
        1) an SO output file that contains the relevant information. Eventually should be added directly to EFACS
        2) an error log showing the self.errors and the files they originated from
        3) a list of price dictionary entries that were not matched - this is only for key self.errors not for price
        discrepancies, which is a different error.

        The routine then returns a list of file paths for the log files and the pdfs with self.errors in them. BOTsend()
        will take this list as an argument and attach the files to the report email.

        Currently will only do entire directory, but could be changed if need be.

        Be aware that a lot of this code is very similar, but the original data are different, so the error checking
        and other routines would be difficult to reuse.
        """

        # iterate over files in the folder
        for originalPDF in glob.glob(self.unprocessedpath + '*.pdf'):
            if self.printstatus:
                print(originalPDF)

            processed = False

            # create empty string for PDF contents
            PDFContents = ''
            # Some POs have multiple items that need to be iteratively processed - store temporarily to ensure that there are no errors before writing to output
            tempitems = []
            # open PDF
            pdfFileObj = open(originalPDF, 'rb')

            # Strict=False prevents some automatic error correction from executing (changing indices of xref table).
            # If the correction fails, the program halts. This off by one error is not catastrophic, so it should be fine to pass.
            try:  # There are other errors that could occur. Just going to catch them all and continue to next file.
                pdfReader = PyPDF2.PdfFileReader(pdfFileObj, strict=False)
                # Combine all pages and concat into single string
                for page in range(pdfReader.numPages):
                    PDFContents = PDFContents + pdfReader.getPage(page).extractText()
            except:
                self.errors.append(['unknown', 'File Issue', originalPDF, "Something went wrong attempting to read the PDF."])
                continue

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
            if 'Draw. format' in PDFContents: # this is for Siemens - no idea if others will appear
                company = 'Siemens'
                self.errors.append([company, 'Engineering drawing', originalPDF, "Appears to be an engineering drawing. Please double check the document."])
            # Estes invoice
            elif 'estes-express' in PDFContents:
                company = 'Estes'
                self.errors.append([company, 'Invoice', originalPDF, "Appears to be an invoice. Please double check the document."])
            # C&T invoice
            elif re.search(r'SI-[0-9]{6}To', PDFContents):
                company = 'C&T'
                self.errors.append([company, 'Invoice', originalPDF, "Appears to be an invoice. Please double check the document."])
            # C&T order acknowledgement
            elif re.search(r'Order Acknowledgement', PDFContents) and re.search(r'S-[0-9]{6}Invoice to', PDFContents):  # Probably don't need both
                company = 'C&T'
                self.errors.append([company, 'Order Acknowledgement', originalPDF, "Appears to be an invoice. Please double check the document."])
            # Multiple delivery dates
            elif re.search(r'Blankets', PDFContents):  # Hopefully, C&T doesn't start selling blankets.
                company = 'NEEDS ATTENTION'
                self.errors.append([company, 'NEEDS ATTENTION', originalPDF, "MULTIPLE DUE DATES ON PO."])
            # GE
            elif 'GE Renewables' in PDFContents:
                company = 'GEC01'
                otherError = False

                try: # Error handling for unmatched regex (NoneType is not subscriptable...)

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

                        date = self.checkDate((dq[0], dq[1], dq[2]), company)
                        date = (date[0], date[1], date[2])

                        partnumber = partnumbers[i][0]
                        quantity = dq[3]
                        qtychk = self.checkQuantities(partnumber, quantity)

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

                        pricechk = self.checkPriceDictionary(company, partnumber, priceper)

                        #Don't break the loop here. Want to keep checking for sum of items (i.e., total price).
                        if pricechk:
                            self.errors.append([company, PONumber, originalPDF, pricechk])
                            otherError = True
                        elif qtychk:
                            self.errors.append([company, PONumber, originalPDF, qtychk])
                            otherError = True
                        elif date[3] and self.datecheck:
                            self.errors.append([company, PONumber, originalPDF, "Problem with date.", date])
                            otherError = True
                        else:
                            date = (date[0], date[1], date[2])  # Check date above outputs tuple with 4 entries - remake as 3
                            tempitems.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])

                except TypeError:
                    self.errors.append([company, 'unknown', originalPDF,
                                        "PO values could not be found. The file was likely miscategorized or the format has changed."])
                    otherError = True

                if float(sumofitems) != float(POTotal):
                    self.errors.append([company, PONumber, originalPDF, "Incorrect total price or number of items.", 'Calcd: ' + str(sumofitems), 'PO: ' + POTotal])
                elif self.checkPOdictionary(PONumber, company) and not otherError:
                    ####output####
                    self.POContents.extend(tempitems)
                    processed = True
                elif not otherError:
                    self.errors.append([company, PONumber, originalPDF,
                                   "File appears to be a duplicate of an already processed PO."])
                #elif otherError - already appended the error in loop
            # Vest01 and Vest05 - Handles some sjol01
            elif 'Vestas Nacelles America' in PDFContents or 'Vestas Blades America Inc' in PDFContents:
                # currently only handles 1 item per PO. Should be easy to fix if it comes up. Need to see an example first.
                if 'SJOELUND US INC.' in PDFContents:
                    company = 'sjol01'
                elif 'Vestas Blades America Inc.' in PDFContents:
                    company = 'VEST05'
                else:
                    company = 'VEST01'

                try: # Error handling for unmatched regex (NoneType is not subscriptable...)

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

                    if PONumber[1] == 'K':
                        date = self.checkDate((dategroup[1], dategroup[2], dategroup[3]), company + 'K')
                    else:
                        date = self.checkDate((dategroup[1], dategroup[2], dategroup[3]), company)

                    pqpattern = """
                    ([0-9,]+)       #quantities allowing for thousands (,)
                    EA[ ]+          #EA is the units and the number of spaces varies by the length of the quantity
                    ([0-9,]+)       #price per item - european notation with no thousands separator
                    [ ]+            #variable number of spaces after unit price
                    ([0-9,]+)       #order total (quanity * unit price)
                    """
                    pricesandquantity = re.search(pqpattern, PDFContents, re.VERBOSE)

                    quantity = float(pricesandquantity[1].replace(',', '.'))
                    priceper = float(pricesandquantity[2].replace(',','.'))
                    POTotal = float(pricesandquantity[3].replace(',','.'))

                    pricechk = self.checkPriceDictionary(company, partnumber, priceper)
                    qtychk = self.checkQuantities(partnumber,quantity)

                    if pricechk:
                        self.errors.append([company, PONumber, originalPDF, pricechk])
                    elif qtychk:
                        self.errors.append([company, PONumber, originalPDF, qtychk])
                    elif quantity * priceper < POTotal - 1 or quantity * priceper > POTotal + 1:  # Added error margin for floating point math.
                        self.errors.append([company, PONumber, originalPDF, "Incorrect total price or number of items.", 'Calcd: ' + float(quantity) * float(priceper), 'PO: ' + POTotal])
                    elif date[3] and self.datecheck:
                        self.errors.append([company, PONumber, originalPDF, "Problem with date.", date])
                    elif self.checkPOdictionary(PONumber, company):
                        #####output#####
                        date = (date[0], date[1], date[2]) # Check date above outputs tuple with 4 entries - remake as 3
                        self.POContents.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])
                        processed = True
                    else:
                        self.errors.append([company, PONumber, originalPDF,"File appears to be a duplicate of an already processed PO."])

                except TypeError:
                    self.errors.append([company, 'unknown', originalPDF, "PO values could not be found. The file was likely miscategorized or the format has changed."])
            # Vest02
            elif 'Vestas - American Wind Technology' in PDFContents:
                company = 'VEST02'

                otherError = False

                try: # Error handling for unmatched regex (NoneType is not subscriptable...)

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
                    ([0-9,.]+)      #quantitiy - capture group [1]
                    [ ]EA[ ]+       #single space + EA + variable white space
                    ([0-9.,]+)      #unit price - capture group [2]
                    [ ]+            #variable spaces
                    ([0-9.,]+)      #item total - capture group [3]
                    """
                    itemline = re.findall(itemlinepatt, PDFContents, re.VERBOSE)

                    POTotal = re.findall(r'Net value[ ]+([0-9.,]+)', PDFContents)[0].replace(',', '')

                    #Need to iterate over these item lists and assign values. Not sure how many are in each.
                    #One delivery date per item is used as a proxy for total number of items
                    #First, we keep track of our total value by adding the line items and check it against the reported PO value at the end
                    sumofitems = 0

                    for i, date in enumerate(alldates):
                        # regex only checks line items 10-90. Probably won't ever be 10 items but throw error just in case
                        if i > 9:
                            otherError=True
                            self.errors.append([company, PONumber, originalPDF, "More than 9 line items. File not processed correctly."])
                            break

                        date = self.checkDate(date, company)

                        partnumber = itemline[i][0]
                        quantity = itemline[i][1].replace(',', '')
                        priceper = itemline[i][2].replace(',', '')
                        itemtotal = itemline[i][3].replace(',', '')

                        pricechk = self.checkPriceDictionary(company, partnumber,priceper)
                        qtychk = self.checkQuantities(partnumber, quantity)

                        if pricechk:
                            self.errors.append([company, PONumber, originalPDF, pricechk])
                            otherError = True
                        elif qtychk:
                            self.errors.append([company, PONumber, originalPDF, qtychk])
                            otherError = True
                        elif float(priceper) * int(quantity) > float(itemtotal) + 1 or float(priceper) * int(quantity) < float(itemtotal) - 1: #Set this as +/- 1 to deal with floating point precision self.errors
                            self.errors.append([company, PONumber, originalPDF, "Incorrect quantity or price for line item.", 'Calcd: ' + float(quantity) * float(priceper), 'PO: ' + POTotal])
                            otherError = True
                        elif date[3] and self.datecheck:
                            self.errors.append([company, PONumber, originalPDF, "Problem with date.", date])
                            otherError = True
                        else:
                            date = (date[0], date[1], date[2])  # Check date above outputs tuple with 4 entries - remake as 3 (ie, remove error message)
                            tempitems.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])

                        sumofitems += float(itemtotal)

                except TypeError:
                    self.errors.append([company, 'unknown', originalPDF,
                                        "PO values could not be found. The file was likely miscategorized or the format has changed."])
                    otherError = True

                # Check comes at end of PO. Above self.errors occur on line items, this check is for total PO price.
                if float(sumofitems) != float(POTotal):
                    self.errors.append([company, PONumber, originalPDF, "Total price not sum of individual items. This error may appear due to a different error in the price checking.", 'Calcd: ' + str(sumofitems), 'PO: ' + POTotal])
                elif self.checkPOdictionary(PONumber, company) and not otherError:
                    #####output#####
                    self.POContents.extend(tempitems)
                    processed = True
                elif not otherError:
                    self.errors.append([company, PONumber, originalPDF, "File appears to be a duplicate of an already processed PO."])
            # Vest04 - very similar to vest02 but there are some spacing issues that are different
            elif 'Vestas Do Brasil Energia' in PDFContents:
                company = 'vest04'
                otherError: False

                try:  # Error handling for unmatched regex (NoneType is not subscriptable...)
                    # Get PO Number
                    PONumber = re.search(r'Purchase order[ ]*([0-9]+)', PDFContents)[1]

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

                    POTotal = re.search(r'Net value[ ]+([0-9.,]+)', PDFContents)[1].replace(',', '')

                    sumofitems = 0
                    for i, date in enumerate(alldates):
                        # regex only checks line items 10-90. Probably won't be 10 items but throw error just in case
                        if i > 9:
                            otherError=True
                            self.errors.append([company, PONumber, originalPDF, "More than 9 line items. File not processed correctly."])
                            break

                        date = self.checkDate(date, company)

                        # Separate the combined terms
                        priceper = re.match(r'[0-9]+.[0-9]{2}', alldata[i][1])[0]
                        itemtotal = float(alldata[i][1].replace(priceper, '').replace(',', ''))

                        # not a great way to separate the quantity from the part number
                        # comes as '290107241,000' with quantity as 1,000 or '153452600' with quantity as 600
                        # need to calculate from the item total and the priceper (both of which we have at high confidence)
                        # Should double check with price dictionary
                        # this is dangerous though because of rounding self.errors.
                        quantity = float(itemtotal) / float(priceper.replace(',', ''))

                        # insert commas so that we match the quantity and not another part of the item string
                        commaquantity = format(int(quantity), ',d')

                        partnumber = alldata[i][0].replace(commaquantity, '')

                        if re.search(r'per[ ]+10', PDFContents) and len(partnumber) > 7:  # For some reason they occasionally price things in batches but still list quantity as total
                            quantity = quantity * 10
                            priceper = float(priceper) / 10
                            commaquantity = format(int(quantity), ',d')
                            partnumber = alldata[i][0].replace(commaquantity, '')

                        pricechk = self.checkPriceDictionary(company, partnumber, priceper)
                        qtychk = self.checkQuantities(partnumber, quantity)

                        if pricechk:
                            self.errors.append([company, PONumber, originalPDF, pricechk])
                            otherError = True
                        elif qtychk:
                            self.errors.append([company, PONumber, originalPDF, qtychk])
                            otherError = True
                        elif float(priceper) * int(quantity) > float(itemtotal) + 1 or float(priceper) * int(quantity) < float(itemtotal) - 1:  # Set this as +/- 1 to deal with floating point precision self.errors
                            self.errors.append([company, PONumber, originalPDF, "Incorrect quantity or price for line item.",'Calcd: ' + float(quantity) * float(priceper), 'PO: ' + POTotal])
                            otherError = True
                        elif date[3] and self.datecheck:
                            self.errors.append([company, PONumber, originalPDF, "Problem with date.", date])
                            otherError = True
                        else:
                            date = (date[0], date[1], date[2])  # Check date above outputs tuple with 4 entries - remake as 3
                            tempitems.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])

                        # check if individual line items sum correctly
                        sumofitems += float(itemtotal)

                except TypeError:
                    self.errors.append([company, 'unknown', originalPDF, "PO values could not be found. The file was likely miscategorized or the format has changed."])
                    otherError = True

                # This check only comes at end of PO (end of for loop). The above self.errors can occur on individual line items, so we need to wait for this check.
                if float(sumofitems) != float(POTotal):
                    self.errors.append([company, PONumber, originalPDF, "Total price not sum of individual items.",
                                   'Calcd: ' + str(sumofitems), 'PO: ' + POTotal])
                elif self.checkPOdictionary(PONumber, company) and not otherError:
                    #####output#####
                    self.POContents.extend(tempitems)
                    processed = True
                elif not otherError:
                    self.errors.append([company, PONumber, originalPDF,"File appears to be a duplicate of an already processed PO."])
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

                    try: # Error handling for unmatched regex (NoneType is not subscriptable...)

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

                        for item in itemline:
                            priceper = item[0]

                            itemtotal = item[1].replace(',', '')
                            date = self.checkDate((item[2], item[3], item[4]), company)
                            partnumber = item[5] + item[6]
                            quantity = item[7].replace(',', '')
                            rev = item[8] # Does revision need to be included in part number?

                            pricechk = self.checkPriceDictionary(company, partnumber, priceper)
                            qtychk = self.checkQuantities(partnumber, quantity)

                            if pricechk:
                                self.errors.append([company, PONumber, originalPDF, pricechk])
                                otherError = True
                            elif qtychk:
                                self.errors.append([company, PONumber, originalPDF, qtychk])
                                otherError = True
                            elif float(priceper) * int(quantity) > float(itemtotal) + 1 or float(priceper) * int(quantity) < float(itemtotal) - 1:  # Set this as +/- 1 to deal with floating point precision self.errors
                                self.errors.append([company, PONumber, originalPDF, "Incorrect quantity or price for line item.", 'Calcd: ' + float(quantity) * float(priceper), 'PO: ' + POTotal])
                                otherError = True
                            elif date[3] and self.datecheck:
                                self.errors.append([company, PONumber, originalPDF, "Problem with date.", date])
                                otherError = True
                            else:
                                date = (date[0], date[1], date[2])  # Check date above outputs tuple with 4 entries - remake as 3 (ie, remove error message)
                                tempitems.append([company, PONumber, date, partnumber, quantity, priceper, POTotal])

                            # check if individual line items sum correctly
                            sumofitems += float(itemtotal)

                    except TypeError:
                        self.errors.append([company, 'unknown', originalPDF,"PO values could not be found. The file was likely miscategorized or the format has changed."])
                        otherError = True

                    # This check only comes at end of PO (end of for loop). The above self.errors can occur on individual line items, so we need to wait for this check.
                    if float(sumofitems) != float(POTotal):
                        self.errors.append([company, PONumber, originalPDF,"Total price not sum of individual items. This error may appear due to a different error in the price checking.",'Calcd: ' + str(sumofitems), 'PO: ' + POTotal])
                    elif self.checkPOdictionary(PONumber, company) and not otherError:
                        #####output#####
                        self.POContents.extend(tempitems)
                        processed = True
                    elif not otherError:
                        self.errors.append([company, PONumber, originalPDF,"File appears to be a duplicate of an already processed PO."])
            # Scanned document
            elif not PDFContents:
                company = 'unknown'
                self.errors.append([company, 'Scanned/empty document', originalPDF, "PDF appears to be a scanned file"])
            # Unidentified PO
            else:
                company = 'unknown'
                self.errors.append([company, 'unknown', originalPDF, "PO not recognized"])
            # check if a folder has been made for a day - if not, create it
            if self.movepdf and processed:
                try:
                    os.rename(originalPDF, self.processedpath + self.datestring + '_' + company + '_' + str(PONumber) + '.pdf')
                except:
                    # In general, this message should not appear. Duplicates should be blocked by PONumber dictionary
                    # Put here to avoid crashes during debugging (don't check PO dictionary flag)
                    self.errors.append([company, PONumber, originalPDF, "This file is a duplicate of an already processed file. It was not moved. There should be another associated error in the error log."])
            elif not processed:
                self.logs.append(originalPDF)

        numerrors = len(self.errors)
        if not numerrors:
            if self.printstatus:
                print('SOs successfully extracted!')
        else:
            if self.printstatus:
                print('There were %s errors. Check the error log.' % numerrors)

    def replaceWithSQLQuery(self):
        """
        Uses an EFACS xls report to collect current stock levels and open orders from the database
        At the very least, this should be replaced with a scrape of the enquiry webpage so user doesn't have to intervene
        """

        wb = xlrd.open_workbook(self.projectedstock)
        data = []

        # Collect items from workbook into empty list
        for s in wb.sheets():
            for row in range(s.nrows):
                values = []
                for col in range(s.ncols):
                    values.append(str(s.cell(row, col).value))

    def projectStock(self):
        """
        Make stock projections into the future based on open orders, forecasting and current stock.
        """

        # Need to read open orders from disk because open orders only arrive once a week and are asynchronus.
        # Reading from disk ensures that we have the latest open orders from customers
        # Can skip reading if lists are still in memory because that indicates that they arrived today
        if not self.HYopenorders:
            self.HYopenorders = readCSVtolist(self.stockprojectionpath + '\OpenOrders\HYOpenOrders.csv')

        if not self.VESTASJOopenorders:
            self.VESTASJOopenorders = readCSVtolist(self.stockprojectionpath + '\OpenOrders\VESTSJOOpenOrders.csv')

        if not self.GEopenorders:
            self.GEopenorders = readCSVtolist(self.stockprojectionpath + '\OpenOrders\GEOpenOrders.csv')

        if datetime.date.today().weekday() == 0:  # 0 = Monday 6 = Sunday --  d should be datetime object in future. (d.weekday())
            pass
            # Remove projected stock based on forecasts

        ######
        #
        #
        #  Update after I figure out best way to get data into program
        #
        #
        ######

        # searchpart = ''  # Initialize this as empty so that the rest of the code works even with no search part
        # if self.mfparts:  # Make sure we have parts to search for
        #     mfpartcalcs = {}
        #     if row[0] == 'Part :':
        #         searchpart = ''  # Reset search part on new item (mf item that we are counting stock for)
        #         for mfpart in self.mfparts:
        #             if row[1] in mfpart:
        #                 searchpart = row[1]
        #                 mfpartcalcs[searchpart] = 0  # Create quantity for the part
        #
        # if searchpart:  # if current looping item is a part we want to keep track of
        #     mfpartcalcs[searchpart] = mfpartcalcs[searchpart] + float(row[4])  # Next to itemstock in while loop

    def writeFiles(self):

        # Need to create directory
        if not os.path.exists(self.stockprojectionpath):
            os.makedirs(self.stockprojectionpath)

        # set output file name
        # including hours and minutes so that this program can be run twice in one day
        outputfilename = datetime.datetime.now().strftime(self.datepath+'%y-%m-%d_%H%M_SalesOrders.csv')
        self.logs.append(outputfilename)
        writeListToCSV(outputfilename, self.POContents)

        errorfilename = datetime.datetime.now().strftime(self.datepath + '%y-%m-%d_%H%M_ErrorLog.csv')
        self.logs.append(errorfilename)
        writeListToCSV(errorfilename, self.errors)

        if self.nopricedictentry:  # Empty by default, so test for existence should be fastest
            # set output file name
            # including hours and minutes so that this program can be run twice in one day
            baddictfilename = datetime.datetime.now().strftime(self.datepath+'%y-%m-%d_%H%M_NoPriceDictionaryEntry.csv')
            self.logs.append(baddictfilename)
            writeListToCSV(baddictfilename, self.nopricedictentry)

        if len(self.GRN) > 1:  # Has a header so we need to test the length.
            grnfilename = datetime.datetime.now().strftime(self.datepath+'%y-%m-%d_%H%M_GRNs.csv')
            self.logs.append(grnfilename)
            writeListToCSV(grnfilename, self.GRN)

        # save the PO dictionary
        writeListToCSV(self.PODICTIONARYPATH, self.polist)

        # Writing open orders to disk and overwriting previous files. Only write if some were found
        # This way we always have the most recent report even though they are
        # received on different days

        # Save HY open orders
        if self.HYopenorders:
            writeListToCSV(self.stockprojectionpath + '\OpenOrders\HYOpenOrders.csv', self.HYopenorders)

        # Save VEST and SJO open orders (arrive in the same email)
        if self.VESTASJOopenorders:
            writeListToCSV(self.stockprojectionpath + '\OpenOrders\VESTSJOOpenOrders.csv', self.VESTASJOopenorders)

        # Save GE open orders
        if self.GEopenorders:
            writeListToCSV(self.stockprojectionpath + '\OpenOrders\GEOpenOrders.csv', self.GEopenorders)

    def calculateManufacturedParts(self):
        if self.printstatus:
            print('Calculating manufactured parts...')

        # No get_sheet_by_name() in xlwt.workbook, so we need a work around.
        # Need to keep the old workbook writer open because we are going to write the output
        # of this function as a new sheet.

        bookcopy = xlrd.open_workbook(self.stockprojectionpath + 'StockProjections.xls')

        # Get the parent and child parts
        mfcalcs = []
        for mfpart in self.mfparts:
            partcalcs = []
            mfqtys = []
            if self.printstatus:
                print(mfpart[0].upper())
            # Get parent stock requests (should all be negative)
            try:
                try:  # Not entirely sure how consistent this capitalization is. Hopefully they don't use Title Case.
                    s = bookcopy.sheet_by_name(mfpart[0].upper()) # These should be unique
                except:
                    s = bookcopy.sheet_by_name(mfpart[0].lower())
            except:
                continue  # Some manufactured parts might not be on order (i.e., won't be in projected stock file). Just move on to the next part if there are no movements.
            # Put date and required inventory in list
            for row in range(s.nrows):
                # Need to pass over variable number of unfulfilled POs and SOs
                # Skipping them will make sure dates line up below
                if "Opening" in str(s.cell(row, 0).value) or self.dateTupleToDatetime(s.cell(row, 1).value) < self.today:  # change to self.today - testing on old data
                    continue
                partcalcs.append([str(s.cell(row, 1).value), str(s.cell(row, 3).value)])  # date, mfpart stock

            # Append all of the required child parts
            for i in range(len(mfpart)//2):  # Get the other parts and their stock levels - should be @ index 1,3,5, etc.
                try: # Upper vs. lower case can be unpredictable. Try upper first (generally more of them).
                    s = bookcopy.sheet_by_name(mfpart[i * 2 + 1].upper())
                except: # Throws an exception if no sheet is found with that name.
                    s = bookcopy.sheet_by_name(mfpart[i * 2 + 1].lower())
                mfqtys.append(mfpart[i * 2 + 2])
                skipped = 0
                for row in range(s.nrows):
                    date = str(s.cell(row, 1).value).split('-')
                    if "Opening" in str(s.cell(row, 0).value) or datetime.date(int(date[0]),int(date[1]),int(date[2])) < self.today:
                        skipped += 1
                        continue
                    row -= skipped
                    # If you get an error here, make sure that the stock projections excel file is current.
                    partcalcs[row].extend([str(s.cell(row, 3).value)])  # extend each row by child part stock

            # Add the smallest number of each manufactured part you can
            for row in partcalcs:
                # mfqtys should have 1 entry per item and row[2:] should be the same length (i.e., stock for each of the items)
                limitingpart = min([float(b) / float(a) for a, b in zip(mfqtys, row[2:])])
                row.extend([limitingpart, float(row[1])+limitingpart])  # row[1] should be negative (requested stock) and limiting part should be positive
            mfcalcs.append(partcalcs)

        for i,part in enumerate(mfcalcs):
            try:
                try:
                    sheet = self.get_sheet_by_name(self.mfparts[i][0].upper())  # xlwt doesn't have get sheet by name. This ensures we have the correct sheet
                    totrows = bookcopy.sheet_by_name(self.mfparts[i][0].upper()).nrows
                except:
                    sheet = self.get_sheet_by_name(self.mfparts[i][0].lower())
                    totrows = bookcopy.sheet_by_name(self.mfparts[i][0].lower()).nrows
            except:
                continue  # As before, skip manufactured parts that don't have activity
            skipped = totrows - len(part)  # skipped rows earlier. want to line up dates
            components = len(self.mfparts[i])//2  # needed to get the correct column for the output (skip number of components per mf item)
            for j,row in enumerate(part):
                sheet.write(j + skipped, 6, row[components + 2])  #Open orders come before these values
                sheet.write(j + skipped, 7, row[components + 3])

        self.book.save(self.stockprojectionpath + 'StockProjections.xls')

    def TEMP(self):
        """
        This whole thing could be refactored to be much more efficient.
        You can really see how I went step by step through the code.
        It's plenty fast though, so probably not worth it.
        Also, it might be easier to fix with all the various breakpoints if something changes.
        """

        # Open workbook and create empty list to store items from workbook
        wb = xlrd.open_workbook(self.projectedstock)
        data = []

        # Collect items from workbook into empty list
        for s in wb.sheets():
            for row in range(s.nrows):
                values = []
                for col in range(s.ncols):
                    values.append(str(s.cell(row, col).value))
                data.append(values)

        openorderdata = []
        openorderdata.extend(self.HYopenorders)
        openorderdata.extend(self.VESTASJOopenorders)
        openorderdata.extend(self.GEopenorders)

        # Only collect rows with data
        # Set collect Row to false in between items (indicated by "Nett change"
        # Set to active when a new part occurs (Part number is separated from rest so doesn't use this variable)
        collectrow = False
        activity = []
        for row in data:
            if row[0] == 'Nett change :':
                collectrow = False
                continue
            elif row[0] == 'Part :':
                activity.append(row)
            elif row[0] == 'Opening stock':
                collectrow = True
            if collectrow:
                activity.append(row)

        # Original data have all parts. Only collect parts with activity predicted.
        # Parts without activity will only have Opening stock reports.
        # Check for row+2 if it is also a Part number, skip ahead
        # Collect other parts
        activeparts = []
        skip = False
        for i, row in enumerate(activity):
            if skip:
                skip = False
                continue
            try:  # If last part in list has no activity, i+2 will cause list index exception. We'll just break here.
                if row[0] == 'Part :' and activity[i + 2][0] == 'Part :':
                    skip = True
                    continue
                else:
                    activeparts.append(row)
            except:
                break

        # Fill in missing dates and the projected stock on those dates
        # Better for visualization and excel doesn't seem to easily/quickly do this
        fulldate = []
        itemstock = 0
        d = self.today
        delta = datetime.timedelta(days=1)
        for i, row in enumerate(activeparts):
            if row[0] == 'Part :' and d <= self.today + datetime.timedelta(days=60):
                if i > 0:  # Populate dates from last PO/SO up to end of forecast range (skip if first part)
                    while d < self.today + datetime.timedelta(days=60):
                        fulldate.append(['', '', '', d, 0, itemstock])
                        d += delta
                fulldate.append(row)  # Append part now - effectively starting new part
                itemstock = 0  # Reset item stock tracker
                d = self.today
                continue
            if row[0] == 'Opening stock':
                itemstock += float(row[4])  # Add opening stock to tracker
                row[3] = self.today  # Insert correct opening date
                row[4] = 0  # Don't want to show opening stock as purchase
                fulldate.append(row)
                continue

            # Has to be defined after the above two conditionals because they don't have dates, causing index errors.
            rowdate = datetime.date(int(row[3].split()[2]), self.abr2num[row[3].split()[1][:3].lower()],
                                    int(row[3].split()[0]))

            # Sometimes there are two activities on one date
            # If the date is the same for 2 or more rows in a row, the activity will be added without incrementing the date
            if rowdate == d - delta :
                row[3] = rowdate  # Convert date format
                itemstock += float(row[4])  # Make stock adjustment
                row[5] = itemstock
                fulldate.append(row)
                continue

            if rowdate < self.today:  # Active entries with due dates before today
                row[3] = datetime.date(int(row[3].split()[2]), self.abr2num[row[3].split()[1][:3].lower()],
                                       int(row[3].split()[0]))
                fulldate.append(row)
                continue

            while d < self.today + datetime.timedelta(days=60):
                if rowdate > d:
                    fulldate.append(['', '', '', d, 0, itemstock])
                elif rowdate == d:
                    row[3] = d  # Convert date format
                    itemstock += float(row[4])  # Make stock adjustment
                    row[5] = itemstock
                    fulldate.append(row)
                    d += delta  # Need to iterate here so we don't get two rows for the same date
                    break
                d += delta

        # Converting to XLS with each part as it's own sheet.
        # Simplifies graphing procedure
        # Could consider using matplotlib or something similar

        n = 0  # Sheet number
        self.book = xlwt.Workbook()  # Create a workbook
        for row in fulldate:
            if row[0] == 'Part :':
                pastdueneeded = True
                # New part - add a sheet for it (can't have / in sheet names)
                # Need to overwrite sheets so that the open items can be added tot he third column
                self.book.add_sheet(str(row[1].replace('/', '-')), cell_overwrite_ok=True)
                sheet = self.book.get_sheet(n)
                n += 1
                rowiterator = 0  # Needed to write rows with correct index (can't use enumerate i bc we start new sheets)

                openquanties = []
                for openitem in openorderdata:  # Check if part number has open items and append to list (added to outfile below)
                    # This could probably be cleaned up some
                    if str(row[1].lower().replace('vestas', '').replace('hyster-','').replace('sjo-','').replace('ge-','')) in openitem[0]:
                        openquanties.append(openitem)

            else:
                for i, openitem in enumerate(openquanties):
                    if self.dateTupleToDatetime(openitem[2]) < self.today and pastdueneeded:  # Make sure we add past due items
                        sheet.write(rowiterator, 1, str(self.dateTupleToDatetime(openitem[2])))  # We use this convoluted approach because it's more typesafe
                        sheet.write(rowiterator, 2,
                                    0)  # Sticking zeros here because excel doesn't graph correctly if the top row has empty cells
                        sheet.write(rowiterator, 3, 0)
                        sheet.write(rowiterator, 4, float(openitem[3]) * -1)
                        sheet.write(rowiterator, 5, str(openitem[1]))
                        rowiterator += 1
                    if self.dateTupleToDatetime(openitem[2]) == row[3] and self.dateTupleToDatetime(openquanties[i - 1][2]) == row[3]: # Two actions on the same day
                        sheet.write(rowiterator, 4, float(openitem[3]) * -1)
                        sheet.write(rowiterator, 5, str(openitem[1] + ' & ' + openquanties[i - 1][1]))
                    elif self.dateTupleToDatetime(openitem[2]) == row[3]:
                        sheet.write(rowiterator, 4, float(openitem[3]) * -1)
                        sheet.write(rowiterator, 5, str(openitem[1]))

                pastdueneeded = False

                sheet.write(rowiterator, 0,
                            str(row[0]))  # Apparently need to write cells 1 at at time (row, column, value)
                sheet.write(rowiterator, 1, str(row[3]))
                sheet.write(rowiterator, 2,
                            float(row[4]))  # Had this in a loop but easier to convert type here than in excel
                sheet.write(rowiterator, 3, float(row[5]))
                rowiterator += 1

        self.book.save(self.stockprojectionpath + 'StockProjections.xls')

if __name__ == "__main__":
    bot = SOBOT()
    bot.debug()  # leaveunread=True POdictionarycheck=False originfolder='./', destfolder='./', PDFtoText=True
    bot.fetchMail()
    bot.scrapePDF()
    bot.parseExcel()
    bot.writeFiles()
    bot.projectStock()
    bot.TEMP()
    bot.calculateManufacturedParts()
    bot.sendMail()

