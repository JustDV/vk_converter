import re
import openpyxl
from datetime import datetime

FILE_NAME = 'file_to_convert.txt'

class Message(object):
    def __init__(self, text):
        self.text = text
        self.type_message = self.get_type_message()
        self.type_user = self.get_type_user()
        self.name_user = self.get_name_user()
        self.date_message = self.get_date_message()
        self.time_message = self.get_time_message()
        self.url_user = self.get_url_user()
        if self.type_user == 'Чат':
            self.chat_id = self.get_chat_id
        self.delet_meta_info()
        self.attached_messages = []
    def get_type_message(self):
        patern = r'Кому:|От кого:'
        type_message_text = re.findall(patern, self.text)
        type_message_text = ''.join(type_message_text)
        return type_message_text
    def get_type_user(self):
        if re.findall('Чат.*\(идентификатор.*\)',self.text):
            type_user = 'Чат'
        if re.findall('Пользователь.*\(https://.*\)',self.text):
            type_user = 'Пользователь'
        if re.findall('Группа.*\(https://.*\)',self.text):
            type_user = 'Группа'
        return type_user
    def get_name_user(self):
        try:
            if self.type_user == 'Пользователь':
                patern = r'Пользователь(.*)\(https.*'
            elif self.type_user == 'Группа':
                patern = r'Группа(.*)\(https.*'
            elif self.type_user == 'Чат':
                if re.findall(r'Участник (.*):',self.text):
                    patern = r'Участник (.*):'
                elif re.findall(r'Пользователь \[.*\|.*\] вышел из беседы,.*', self.text):
                    patern = r'Пользователь \[.*\|(.*)\] вышел из беседы'
                elif re.findall(r'Чат.*\(идентификатор чата \d*, (.*) https://.*', self.text):
                    patern = r'Чат.*\(идентификатор чата \d*, (.*) https://.*'
                elif re.findall(r'(.*)\(https://.*\).\<a class=.*', self.text):
                    patern = r'(.*)\(https://.*\).\<a class=.*'
            name_user = re.findall(patern, self.text)
            name_user = ''.join(name_user)
            return name_user
        except:
            name_user = 'Неизвестно'
            return name_user
    def get_url_user(self):
        if self.type_user == 'Чат':
            patern = r'Чат  \(идентификатор чата.*(https.*)\)'
        if self.type_user == 'Пользователь':
            patern = r'Пользователь.*\((https.*)\)'
        if self.type_user == 'Группа':
            patern = r'Группа.*\((https://.*)\)'
        url_user = re.findall(patern, self.text)
        url_user = ''.join(url_user)
        return url_user
    def get_date_message(self):
        patern = r'\d{2}\.\d{2}\.\d{4}'
        date_message = re.findall(patern, self.text)
        date_message = ''.join(date_message)
        return date_message
    def get_time_message(self):
        patern = r'\d{2}:\d{2}:\d{2}'
        time_message = re.findall(patern, self.text)
        time_message = ''.join(time_message)
        return time_message
    def get_chat_id(self):
        patern = r'Чат  \(идентификатор чата (\d*),.* https://.*'
        chat_id = re.findall(patern, self.text)
        chat_id = ''.join(chat_id)
        return chat_id
    def delet_meta_info(self):
        if self.type_user == 'Чат':
            patern = r'(Кому:|От кого:)\n(?:(?:Чат.*https.*\n)|(?:Чат.*))\d{2}.\d{2}.\d{4}.*\n'
        else:
            patern = r'(Кому:|От кого:)\nПользователь.*\(https.*\n\d{2}.\d{2}.\d{4}.*\n'
        self.text = re.sub(patern, ' ',self.text)
def read_file(file):
    filetxt = open(file, 'r')
    text = None
    inbox = []
    for line in filetxt:
        if re.findall(r'((?:Кому:.*)|(?:От кого:.*))', line):
            if text is not None:
                inbox.append(Message(text))
                text = line
            else:
                 text = line
        elif text is not None:
            text += line
    if text is not None:
        inbox.append(Message(text))
    return inbox
def xl_write(inbox):
    row = 0
    wb = openpyxl.Workbook()
    sheet = wb.active
    for mess in inbox:
        row += 1
        sheet['A'+str(row)] = mess.type_message
        sheet['B'+str(row)] = mess.date_message
        sheet['C'+str(row)] = mess.time_message
        sheet['D'+str(row)] = mess.type_user
        sheet['E'+str(row)] = mess.name_user
        sheet['F'+str(row)] = mess.text
        sheet['G'+str(row)] = mess.url_user
    dt = datetime.now()
    filename = FILE_NAME +'__'+dt.strftime("%Y%m%d_%I%M%S") + '.xlsx'
    wb.save(filename)

if __name__ == "__main__":
    inbox_mail = read_file(FILE_NAME)
    xl_write(inbox_mail)