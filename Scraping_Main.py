#import config
from Utilities.myLogger import logger
import sys
from gmail import Gmail
from gmail.exceptions import AuthenticationError
from PageEvents import gmailEvents
from pprint import pprint
from datetime import timedelta


def doScrape(Con, q):
    try:
        g = Gmail()
        ################ LOGIN #####################################
        q.put(('Logging In', 5),)
        logger.info("Logging In")
        try:
            g.login(Con.Username, Con.Password)
        except AuthenticationError:
            logger.exception(sys.exc_info())
            q.put(('Login Failed', 100),)
            return "AUTH_ERROR"
        ############################################################

        ################ GET LABEL MAILBOX #########################
        mailbox = None
        q.put(('Getting Mailbox', 10),)
        logger.info("Getting Mailbox")
        try:
            if Con.Label.lower() == 'inbox':
                mailbox = g.inbox()
            else:
                mailbox = g.mailbox(Con.Label)
        except:
            logger.exception(sys.exc_info())
            q.put(('Problem in fetching Gmail Label', 100),)
            return "LABEL_FETCH_ERROR"
        if not mailbox:
            q.put(('Gmail Label Not Found', 100),)
            return "LABEL_NOT_FOUND"
        ############################################################

        ################ GET EMAILS ################################
        mails = None
        q.put(('Searching For Emails', 15),)
        logger.info("Searching Emails")
        try:
            afterDate = Con.FromDate - timedelta(days=1)
            beforeDate = Con.ToDate + timedelta(days=1)
            mails = mailbox.mail(subject='Fiverr: Congrats! You have a new order',
                                 after=afterDate, before=beforeDate)
            mails.extend(mailbox.mail(subject='just ordered an Extra',
                                      after=afterDate, before=beforeDate))
            # mails = mailbox.mail(after=Con.FromDate, before=Con.ToDate)
        except:
            logger.exception(sys.exc_info())
            q.put(('Problem in searching for emails', 100),)
            return "EMAILS_FETCH_ERROR"
        if len(mails) == 0:
            q.put(('No Emails Found with search criteria', 100),)
            return "NO_EMAIL_FOUND"
        ############################################################

        ################ FETCH EMAILS ##############################
        q.put(('Fetching Emails', 20),)
        logger.info("Scraping Order Data From Emails")
        Con.Orders = []
        logger.info("Num of Emails found: " + str(len(mails)))
        try:
            for mail in mails:
                msg = "Fetching Email " + str(mails.index(mail)+1) + ' of ' + str(len(mails))
                per = 20 + int((float(mails.index(mail)+1) * 100.0 * 0.6 / float(len(mails))))
                q.put((msg, per),)
                #logger.info(msg)
                mail.fetch()
                gmailEvents.extract_orders_from_email(mail, Con)
        except:
            logger.exception(sys.exc_info())
            q.put(('Problem in fetching emails', 100),)
            return "EMAIL_FETCH_ERROR"
        ############################################################

        # return 'SUCCESS'

        ################ CALCULATE TOTAL AMOUNT ####################
        q.put(('Calculating Total and Revenue', 85),)
        logger.info("Calculating Total Amount")
        gmailEvents.calculate_total_amount(Con)
        ############################################################

        ################ GENERATE XLS ##############################
        q.put(('Generating XLS', 90),)
        logger.info("Generating XLS")
        gmailEvents.generate_xls(Con)
        ############################################################

        q.put(('Logging Out of Gmail', 95),)
        g.logout()
        q.put(('SUCCESS', 100),)
        return 'SUCCESS'
    except:
        if g:
            g.logout()
        logger.exception(sys.exc_info())
        q.put(('Error Occurred', 100),)
        return "ERROR"