import fitz
import pandas as pd
import os
import re

all_parsed_data_list = []

def phone_check(phone):
    """Using regex to see if the phone is contained inside the attribute

    Args:
        phone (str): string which may contain phone

    Returns:
        bool: The return value. True for success, False otherwise.
    """
    find_phone = re.search(
        r'(\(\d{3}\)[\s\-]?\d{3}[\-\s\.]*\d{4})|(\d{10,15})|(\(\d{3}\) \d{7,15})|(\d{3}\.\d{3}\.\d{4})|(\d{3}[\-\s]*\d{3}[\-\s]*\d{4})|(\+\d.+)|(\d{1,3}\s?\d{7,15})|(\d{3,5}\s\d{3}\s\d{3}.+)|(\(404)',
        phone)
    if find_phone:
        return True
    else:
        return False

def data_processing(user_contents):
    parsed_data = []
    user_contents_length = len(user_contents)

    user_title = ''
    user_company = ''
    user_address = ''
    user_city = ''
    user_state = ''
    user_zipcode = ''
    user_phone = ''

    if user_contents_length:
        for user_content_index in range(user_contents_length):
            user_content = user_contents[user_content_index]

            if user_content_index == 0:
                # First Name & Last Name
                user_names = user_content.split(' ')
                user_first_name = ''
                user_last_name = ''
                if len(user_names):
                    for user_name_index in range(len(user_names)):
                        if user_name_index == 0:
                            user_first_name = user_names[user_name_index]
                        if user_name_index != 0:
                            if user_last_name:
                                user_last_name = user_last_name + ' '
                            user_last_name = user_last_name + user_names[user_name_index]
                parsed_data.append(user_first_name)
                parsed_data.append(user_last_name)

            if user_content_index == 1:
                # Title
                user_title = user_content

            if user_content_index == 2:
                # Company
                if user_contents_length == 3:
                    if not phone_check(user_content):
                        user_company = user_content
                else:
                    user_company = user_content

            if 3 <= user_content_index < user_contents_length - 1:
                # Address & City & State & Zip Code
                #
                if user_address:
                    user_address = user_address + ' '
                user_address = user_address + user_content

                if user_content_index == user_contents_length - 2:
                    find_zipcode = re.search(r'([0-9]{5}(?:-[0-9]{4})?)', user_content)
                    if find_zipcode:
                        user_zipcode = find_zipcode.group(0)

                    user_address_data = user_content.split(', ')
                    if len(user_address_data):
                        user_city = user_address_data[0]
                    if len(user_address_data) > 1:
                        user_address_sub_data = user_address_data[1].split(' ')
                        if len(user_address_sub_data):
                            user_state = user_address_sub_data[0]


            if user_content_index == user_contents_length - 1:
                if phone_check(user_content):
                    user_phone = user_content

        parsed_data.append(user_title)
        parsed_data.append(user_company)
        parsed_data.append(user_address)
        parsed_data.append(user_city)
        parsed_data.append(user_state)
        parsed_data.append(user_zipcode)
        parsed_data.append(user_phone)


        all_parsed_data_list.append(parsed_data)

pdf_page_content_divider = '________________________________________________________________________________';
pdf_page_footer_content = 'No part of this list can be reproduced or stored in a retrieval system in any form without prior written permission from ICSC';

pdf_file = [file for file in os.listdir() if file.endswith('.pdf')][0]
pdf_doc = fitz.open(pdf_file)
for page_no in range(pdf_doc.page_count):
    pdf_page = pdf_doc[page_no]
    pdf_wordlist = pdf_page.get_text_blocks()
    pdf_content_start = False
    for pdf_word in pdf_wordlist:
        if pdf_content_start and pdf_page_footer_content != pdf_word[4].strip():
            user_contents = pdf_word[4].split('\n')
            user_data = []
            for user_content in user_contents:
                if user_content:
                    user_data.append(user_content)
            data_processing(user_data)

        if pdf_page_content_divider == pdf_word[4].strip():
            pdf_content_start = True

columns = ['First Name', 'Last Name', 'Title', 'Company', 'Address', 'City', 'State', 'Zip Code', 'Phone']
df = pd.DataFrame(all_parsed_data_list, columns=columns)
# file_name = 'ICSC RECon NY 2020 Virtual textract.csv'
# df.to_csv(file_name, index=False)
file_name = 'ICSC RECon NY 2020 Virtual textract.xlsx'
df.to_excel(file_name, index=False)