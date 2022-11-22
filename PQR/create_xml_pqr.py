import csv
import datetime
import pandas as pd
import lxml.etree
from dateutil.parser import parse


excel_table = 'PQR_APR_Control file_Template_Példákkal.xlsx'
sheet_name = 'PQR - WS'
csv_file = 'pqr.csv'
xml_file = 'PQR_WS_TEST.xml'



def csv_to_xml():

    csvData = csv.reader(open(csvFile, encoding="utf8"),skipinitialspace=True)
    xmlData = open(xmlFile, 'w', encoding="utf8")
    xmlData.write('<?xml version="1.0" encoding="utf-8"?>' + "\n")
    xmlData.write('<import xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' + "\n")
    xmlData.write("\n")
    
    row_num = 0

    for row in csvData:
        if row_num == 0:
            categories = row
            categories.append("Exit - for category closing tag")
            for i in range(len(categories)):
                categories[i] = categories[i].replace('&','&amp;')
        
        elif row_num == 1:
            tags = row
            # replace spaces w/ underscores in tag names
            for i in range(len(tags)):
                tags[i] = tags[i].replace(' ', '_')
        elif row_num == 2:
            set_options = row
            set_options.append("Exit - for category closing tag")
            for i in range(len(set_options)):
                set_options[i] = set_options[i].rstrip()
        elif row_num == 3:
            tag_options = row

        else:
            print('data')

        




def is_date(string, fuzzy=False):
    """
    Return whether the string can be interpreted as a date.

    :param string: str, string to check for date
    :param fuzzy: bool, ignore unknown tokens in string if True
    """
    try: 
        parse(string, fuzzy=fuzzy)
        return True

    except ValueError:
        return False


def xlsx_to_csv():
    cols = pd.read_excel(excel_table, sheet_name=sheet_name ,header=None,nrows=1).values[0] # read first row
    df = pd.read_excel(excel_table, sheet_name=sheet_name ,header=None, skiprows=1) # skip 1 row
    df.columns = cols
    df.to_csv (csvFile, index = None, header=True)
    

def write_header():



if __name__ == "__main__":
    xlsx_to_csv()
    csv_to_xml()
