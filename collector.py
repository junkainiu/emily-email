# -*- coding: utf-8 -*-
import settings
import pandas as pd
import os


class ExcelInfoCollector(object):

    def __init__(self):
        self.mail_data = pd.read_excel("mail_data.xlsx").set_index('Code')
        self.code = pd.read_excel("mail_data.xlsx", 1)
        self.attach_dir = os.getcwd()
        self.init_settings()

    def init_settings(self):
        try:
            self.text_template = settings.TEXT_TEMPLATE
            self.title_prefix = settings.TITLE_PREFIX
            self.attach_map = settings.ATTACH_MAP
            self.allow_missing = settings.ALLOW_MISSING
        except Exception as e:
            raise Exception("Invalid settings! %s" % e)


    def fill_dealer_text(self, vm_dealer_name):
        html_text = text.replace('VW_dealer_name', vm_dealer_name)
        return html_text

    def prepare_mail_infos(self):
        all_mail_infos = []
        for dealer in self.code.Code:
            vm_dealer_name = self.mail_data.ix[dealer, 'VW_dealer_name']
            mail_text = self.fill_dealer_text(vm_dealer_name)
            subject = self.title_prefix + self.mail_data.ix[dealer, 'mail_subject']
            recipient = self.mail_data.ix[each_dealer, 'TO']
            attachments = self.make_attachments(dealer)
            info = [dealer, mail_text, subject, recipient, attachments]
            all_mail_infos.append(info)
        return all_mail_infos

    def make_attachments(self, dealer):
        attachments = []
        for dir_name, ex_column_name in self.attach_map:
            full_dir_path = os.path.join(self.attach_dir, dir_name)
            file_name = self.mail_data.ix[dealer, ex_column_name] + '.pdf'
            full_attach_path = os.path.join(self.full_dir_path, file_name)
            if not os.path.exists(full_attach_path):
                if not allow_missing:
                    raise Exception("Missing File %s" % full_attach_path)
            attachments.append(full_attach_path)
