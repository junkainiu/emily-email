# -*- coding: utf-8 -*-
"""
Created on Mon Dec  5 15:40:15 2016

@author: tanf
"""

import win32com.client as win32
import pandas as pd
import time
import os
import random
import glob
import settings

# ATTACHMENT1 = os.getcwd() + '\\attachments\\'
ATTACHMENTDIR = os.getcwd()

def outlook(olook, code, text, subject, recipient, attachments):
    mail = olook.CreateItem(win32.constants.olMailItem)
    mail.Recipients.Add(recipient)
    mail.Subject = subject
    mail.HTMLBody = text
	for attachment in attachments:
    	mail.Attachments.Add(attachment)
    mail.Send()


def main():
	app = 'Outlook'
    olook = win32.gencache.EnsureDispatch('%s.Application' % app)
    mail_data = pd.read_excel("mail_data.xlsx").set_index('Code')
    code = pd.read_excel("mail_data.xlsx", 1)
    for each_dealer in code.Code:
        VW_dealer_name = mail_data.ix[each_dealer, 'VW_dealer_name']
        each_html_text = fill_dealer_text(TEXT, VW_dealer_name)

        each_subject = settings.TITLE_PREFIX + mail_data.ix[each_dealer, 'mail_subject']
        recipient = mail_data.ix[each_dealer, 'TO']
		attachments = []
        attachment1 = ATTACHMENT1 + mail_data.ix[each_dealer, 'mail_subject'] + '.pdf'

        outlook(olook, each_dealer, each_html_text, each_subject, recipient, attachments)
        time.sleep(random.randint(1, 10))
	# olook.Quit()


if __name__ == '__main__':
	main()
