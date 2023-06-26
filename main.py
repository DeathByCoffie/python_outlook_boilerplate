import pandas as pd
import os
import glob
import win32com.client as win32
import jinja2
import logging
import traceback
import time
import datetime as dt

def get_latest_file(path):
    list_of_files = glob.glob(path)
    latest_file = max(list_of_files, key=os.path.getctime)

    return latest_file


def sendMail(inboxName: str, 
    sendTo: str, 
    subject: str, 
    body: str, 
    send: bool, 
    cc: str | None = '', 
    attachment: list[str] | None = []):
    '''Send email though Outlook

    Parameters
    ----------
    inboxName : str
        Name of inbox to send email out of
    sendTo : str
        List of send recipient emails
    subject : str
        Subject line of email
    body : str
        Plain text or HTML
    send : bool
        Send or Draft
    
    '''
    outlook = win32.Dispatch('outlook.application')
    for account in outlook.Session.Accounts:
        if str.upper(account.DisplayName) == str.upper(inboxName):
            mail = outlook.CreateItem(0)
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
            mail.To = sendTo
            mail.CC = cc
            mail.Subject = subject
            mail.HTMLBody = body

            mail.ReadReceiptRequested = False

            for a in attachment:
                mail.Attachments.Add(a)

            if send:
                mail.Send()
            elif not send:
                mail.Save()


def main():
    # Do stuff here
    return


if __name__ == '__main__':
    # Setup logging
    logger = logging.getLogger(__name__)
    formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    logger.setLevel(logging.DEBUG)
    # Setup stream handler
    stream_handler = logging.StreamHandler()
    stream_handler.setLevel(logging.INFO)
    stream_handler.setFormatter(formatter)
    # Setup file handler
    LOG_FILENAME = dt.datetime.now().strftime('./logs/logfile_%m%d%Y_%H%M%S.log')
    file_handler = logging.FileHandler(filename=LOG_FILENAME)
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    # add handlers to logger
    logger.addHandler(stream_handler)
    logger.addHandler(file_handler)

    try:
        main()

    except Exception as e:
        logger.critical(f'Somthing went wrong: {e}')
        logging.critical(traceback.format_exc())
        time.sleep(120)