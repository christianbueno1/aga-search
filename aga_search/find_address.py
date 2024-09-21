import os
import unidecode
# import variables from config.py
from aga_search.config import DOWNLOADS_PATH, DB_FILE_PATH, START_SHEET_NAME, END_SHEET_NAME, COLUMNS, ROW_START, BASE_PERCENTAGE_SIMILARITY, LEVEL, HEADER_ROW, FILE_PATH
import pandas as pd


def find_address_in_excel(file_name, search_address, search_county, start_sheet_name=START_SHEET_NAME, end_sheet_name=END_SHEET_NAME, columns=COLUMNS, header_row=HEADER_ROW, row_start=ROW_START, base_percentage_similarity=BASE_PERCENTAGE_SIMILARITY, level=LEVEL) -> list:
    """
    Find the address in the excel file
    """
    header = 'Unnamed: 1'
    max_similarity = base_percentage_similarity
    address_find = None
    county_find = None
    sheet_name_find = None
    count_all_sheet_coincidences = 0
    sheets_dict = pd.read_excel(file_name, sheet_name=None, header=header_row, usecols=columns)
    sheets_data = {}
    sheet_data: dict = {}
    address_eq: bool = False
    for sheet_name, df in sheets_dict.items():
        rows: list = []
        if sheet_name < start_sheet_name or sheet_name > end_sheet_name:
            continue
        print(f"Sheet namee: {sheet_name}")
        # remove the first row
        # print(f"df: {df}")
        df = df.iloc[row_start:]
        # apply compare_address_similarity function to the data
        # count_rows = len(df)
        for index, row in df.iterrows():
            # address to str
            # print(f"Row: {row} row.values: {row.values}")
            roww: dict = {}
            address = str(row.values[0])
            # parroquias
            county = str(row.values[1])
            if not pd.isna(address):
                match level:
                    case 0:
                        if is_address_equals(search_address, address):
                            similarity = 1
                            address_eq = True
                            county_eq = is_county_equals(search_county, county)
                            similarity = address_eq and county_eq
                            # print(f"Found address:: {address}, county: {county}")
                        else:
                            address_in_address = is_address_in_address_not_vice_versa(search_address, address)
                            county_eq = is_county_equals(search_county, county)
                            similarity = address_in_address and county_eq
                            if similarity:
                                if not(is_both_address_one_word(search_address, address)) and not(is_both_address_more_than_one_word(search_address, address)):
                                    similarity = False
                    case 1:
                        pass
                # use equal to know if there is a match                        
                if similarity >= max_similarity:
                    roww['address'] = address
                    roww['county'] = county
                    count_all_sheet_coincidences += 1
                    max_similarity = similarity
                    address_find = address
                    county_find = county
                    sheet_name_find = sheet_name
                    # print(f"Sheet name: {sheet_name}")
                    # print(f"Similarity: {similarity}")
                    print(f"Found address: {address}, county: {county}")
                    # print(f"max_similarity: {max_similarity}")
                    # ask to type enter
                    # input("Press Enter to continue...")
                    rows.append(roww)
            # print(f"rowsss: {rows}")
            if address_eq:
                break
        #     if similarity == 1:
        #         break
        # if max_similarity == 1:
        #     break
        # print(f"rowsss: {rows}")
        # print enter to continue
        # input("Press Enter to continue...")
        # add rows to sheet_data only if there are rows
        if len(rows) > 0:
            sheet_data[sheet_name] = rows
        if address_eq:
            break
    # add sheet_data to sheets_data only if there are rows
    if len(sheet_data) > 0:
        sheets_data[search_address] = sheet_data

    print(f"\nMax similarity: {max_similarity}")
    print(f"address: {search_address}")
    print(f"Found address(pattern): {address_find}")
    print(f"Found county(pattern): {search_county}")
    print(f"Sheet name: {sheet_name_find}")
    print(f"address_eq: {address_eq}")
    print(f"len sheet_data: {len(sheet_data)}")
    print(f"sheet_data: {sheet_data}")
    print(f"Count of all coincidences: {count_all_sheet_coincidences}\n")
    # print sheets_data
    for address, sheets in sheets_data.items():
        print(f"Address: {address}")
        for sheet_name, rows in sheets.items():
            print(f"Sheet name: {sheet_name}")
            for row in rows:
                print(f"Row: {row}")

    # ask for typing enter
    # input("Press Enter to continue...\n")
    if len(sheet_data) > 1 and not address_eq:
        sheet_name_find = None
    return sheet_name_find, address_find


# remove special characters and convert to lowercase
def remove_special_characters(phrase):
    # convert to lowercase
    phrase = phrase.lower()
    # remove accents
    phrase = unidecode.unidecode(phrase)
    # remove special characters, allow only letters, numbers and spaces
    phrase = ''.join(e for e in phrase if e.isalnum() or e.isspace())
    # remove special characters, allow only letters and spaces
    # phrase = ''.join(e for e in phrase if e.isalpha() or e.isspace())
    # remove leading and trailing whitespaces
    phrase = phrase.strip()
    return phrase

# remove words like mz, sl, etc
def remove_special_words(phrase):
    """
    Remove special words like mz, sl, etc
    """
    special_words = ['mz', 'sl']
    phrase = phrase.split()
    # remove numbers following the special words
    for i, word in enumerate(phrase):
        if word in special_words:
            if i + 1 < len(phrase):
                if phrase[i + 1].isnumeric():
                    phrase[i + 1] = ''
    phrase = [word for word in phrase if word not in special_words]
    phrase = ' '.join(phrase)
    # remove leading and trailing whitespaces
    phrase = phrase.strip()
    return phrase

# check if counties are equals
def is_county_equals(county1, county2) -> bool:
    """
    Check if county1 is equals to county2
    """
    # remove word parroquia
    county1 = remove_parroquia(county1)
    county2 = remove_parroquia(county2)
    county1 = remove_special_characters(county1)
    county2 = remove_special_characters(county2)
    return county1 == county2

# remove word parroquia
def remove_parroquia(phrase):
    """
    Remove the word parroquia
    """
    phrase = phrase.lower()
    phrase = phrase.replace('parroquia', '')
    phrase = phrase.strip()
    return phrase
    
# check if addresses are equals
def is_address_equals(address1, address2) -> bool:
    """
    Check if address1 is equals to address2
    """
    address1 = remove_special_characters(address1)
    address2 = remove_special_characters(address2)
    # remove special words
    address1 = remove_special_words(address1)
    address2 = remove_special_words(address2)
    return address1 == address2
    
# function using in operator
# check if address1 is in address2 or vice versa
def is_address_in_address(address1, address2) -> bool:
    """
    Check if address1 is in address2 or vice versa
    """
    result = False
    # remove special characters and convert to lowercase
    address1 = remove_special_characters(address1)
    address2 = remove_special_characters(address2)
    # remove special words
    address1 = remove_special_words(address1)
    address2 = remove_special_words(address2)
    # print(f"Address1: {address1}")
    # print(f"Address2: {address2}")
    return address1 in address2 or address2 in address1    

# check if address1 is in address2 but not vice versa
def is_address_in_address_not_vice_versa(address1, address2) -> bool:
    """
    Check if address1 is in address2 but not vice versa
    """
    address1 = remove_special_characters(address1)
    address2 = remove_special_characters(address2)
    # remove special words
    address1 = remove_special_words(address1)
    address2 = remove_special_words(address2)
    # print(f"Address1: {address1}")
    # print(f"Address2: {address2}")
    return address1 in address2
    

# check if address1 and address2 have both one word
def is_both_address_one_word(address1, address2) -> bool:
    """
    Check if address1 and address2 have both one word
    """
    address1 = address1.split()
    address2 = address2.split()
    return len(address1) == 1 and len(address2) == 1

# check if address1 and address2 have both more than one word
def is_both_address_more_than_one_word(address1, address2) -> bool:
    """
    Check if address1 and address2 have both more than one word
    """
    address1 = address1.split()
    address2 = address2.split()
    return len(address1) > 1 and len(address2) > 1


if __name__=="__main__":
    search_address = "Coop los ángeles II"
    county = "Ximena"
    address2 = "PRE-COOP. LOS ÁNGELES 2"
    sheet_name = find_address_in_excel(FILE_PATH, search_address, county)
    print(f"Sheet name: {sheet_name}")
