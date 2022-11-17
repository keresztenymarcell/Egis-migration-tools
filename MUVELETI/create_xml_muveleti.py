import csv
import datetime
import pandas as pd
import lxml.etree
from dateutil.parser import parse


# excel_table = 'muveleti_adatok_control.xlsx'
# excel_table = 'muveleti_adatok_BW_DOCS_02.xlsx'
excel_table = 'Migrációs adatok - Összevont_20221006.xlsx'
sheet_name = 'DOCS'
# sheet_name = 'WS3'
csvFile = 'Muveleti_test.csv'
xmlFile = 'Muveleti_test_20221006_docs_01.xml'

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

def xlsx2csv():
    cols = pd.read_excel(excel_table, sheet_name=sheet_name ,header=None,nrows=1).values[0] # read first row
    df = pd.read_excel(excel_table, sheet_name=sheet_name ,header=None, skiprows=1) # skip 1 row
    df.columns = cols
    df.to_csv (csvFile, index = None, header=True)
    

def csv2xml():

    
    csvData = csv.reader(open(csvFile, encoding="utf8"),skipinitialspace=True)
    xmlData = open(xmlFile, 'w', encoding="utf8")
    xmlData.write('<?xml version="1.0" encoding="utf-8"?>' + "\n")
    # there must be only one top-level tag
    xmlData.write('<import xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' + "\n")
    xmlData.write("\n")
    rowNum = 0
    for row in csvData:
        if rowNum == 0:
            categories = row
            categories.append("Exit - for category closing tag")
            for i in range(len(categories)):
                categories[i] = categories[i].replace('&','&amp;')
        
        elif rowNum == 1:
            tags = row
            # replace spaces w/ underscores in tag names
            for i in range(len(tags)):
                tags[i] = tags[i].replace(' ', '_')
        elif rowNum == 2:
            set_options = row
            set_options.append("Exit - for category closing tag")
            for i in range(len(set_options)):
                set_options[i] = set_options[i].rstrip()
        elif rowNum == 3:
            tag_options = row

        else: 
            m = ""
            f = ""
                      
            for i in range(len(tags)):
                if tags[i].find("location") != -1:
                    row[i] = row[i].replace('\\',':')
                    row[i] = row[i].replace('&','&amp;')
                if tags[i] == "title":
                    row[i] = row[i].replace('&','&amp;')

            for i in range(len(tags)):
                if tags[i].find("mime") != -1:
                    if row[i] != "":
                        m = row[i]
                        xmlData.write('<node type="document" action="create">' + "\n")
                        xmlData.write('\t' + '<rmclassification filenumber="1-1-1-6-13">' + "\n")
                        xmlData.write('\t\t' + '<status>EFFECTIVE</status>' + "\n")
                        xmlData.write('\t' + '</rmclassification>' + "\n")
                        xmlData.write('\t' + '<securityclearance>' + "\n")
                        xmlData.write('\t\t' + '<securitylevel>66</securitylevel>' + "\n")
                        xmlData.write('\t' + '</securityclearance>' + "\n")

                    else:
                        xmlData.write('<node action="create" type="businessworkspace">' + "\n")
                        xmlData.write('\t' + '<otsapwksp>' + "\n")
                        xmlData.write('\t' + '<template action="createfromtemplate">Content Server Document Templates:Egis Magyarország:Operations:Formula and Processing Instructions:Gyártás-Műveleti Folyamat</template>' + "\n")
                        xmlData.write('\t' + '</otsapwksp>' + "\n")

            for i in range(len(tags)):
                if tags[i].find("file") != -1:
                    if len(row[i])>0:
                        f = row[i]
                        if(f[-4:].lower() != ".pdf"):
                            if(' ' in row[i][-1]):
                                while(row[i][-1] == ' '):
                                    row[i] = row[i][:-1]
                                    row[i] = row[i]+".pdf"
                                    print("### "+row[i])
                            elif(f[-4:].lower() != ".pdf"):
                                row[i] = row[i].strip()
                                row[i] = f+".pdf"
                                print(row[i])


            for i in range(len(tags)):
                if (len(set_options[i]) > 0):
                    first_set = i
                    break

            set_name = ''
            alone = 0
            lastCategory = ''
            for i in range(len(tags)):
                if tags[i].find("attribute") != -1:
                    if len(row[i])>0:
                        
                        if(is_date(row[i]) == True):
                            row[i] = row[i].replace(' 00:00:00','')
                            row[i] = row[i].replace(' ','')
                            row[i] = row[i].replace('/','')
                            row[i] = row[i].replace(':','')
                            row[i] = row[i].replace('.','')
                            row[i] = row[i].replace('-','')
                            # print("DÁTUM: "+row[i])

                        row[i] = row[i].replace('&','&amp;')
                        tag_options[i] = tag_options[i].rstrip()
                        
                        if (categories[i] != categories[i-1]) and (categories[i] != categories[i+1]):
                            
                            xmlData.write('	<category name="'+categories[i]+'">' + "\n")
                            xmlData.write('\t\t' + '<' + tags[i] + ' name="'+tag_options[i]+'">' + row[i] + '</' + tags[i] + '>' + "\n")
                            xmlData.write('	</category>'+"\n")
                            xmlData.write("\n")
                        
                        elif(categories[i] != categories[i-1]): 
                            xmlData.write('	<category name="'+categories[i]+'">' + "\n")
                            lastCategory = categories[i]

                            if (len(set_options[i]) > 0):
                                if(set_name == set_options[i]):
                                    # print("0set_name: "+set_name+" --i:",i)
                                    xmlData.write('\t\t\t' + '<' + tags[i] + ' name="'+tag_options[i]+'">' + row[i] + '</' + tags[i] + '>' + "\n")
                                    checker = len(tags) - i + 1
                                    for v in range(checker)[1:]:
                                        if(set_name == set_options[i+v]):
                                            if(len(row[i+v]) > 0):
                                                break
                                            else:
                                                v=v+1
                                        else:
                                            # print("1set_name: "+set_name+" set_options[i+v]: "+set_options[i+v]+" --v:",v," --i:",i)
                                            xmlData.write('\t\t' + '</setattribute>' + "\n")
                                            break
                                
                                else:
                                    if(i==first_set):
                                        # print("2set_name: "+set_name+" --i:",i)
                                        xmlData.write('\t\t' + '<setattribute name="'+set_options[i][:-2]+'">' + "\n")
                                        xmlData.write('\t\t\t' + '<' + tags[i] + ' name="'+tag_options[i]+'">' + row[i] + '</' + tags[i] + '>' + "\n")
                                        
                                    elif(i==len(tags)):
                                        # print("3set_name: "+set_name+" --i:",i)
                                        xmlData.write('\t\t\t' + '<' + tags[i] + ' name="'+tag_options[i]+'">' + row[i] + '</' + tags[i] + '>' + "\n")
                                        xmlData.write('\t\t' + '</setattribute>' + "\n")
                                        
                                    else:
                                        # print("4set_name: "+set_name+" --i:",i)
                                        xmlData.write('\t\t' + '<setattribute name="'+set_options[i][:-2]+'">' + "\n")
                                        xmlData.write('\t\t\t' + '<' + tags[i] + ' name="'+tag_options[i]+'">' + row[i] + '</' + tags[i] + '>' + "\n")
                                        
                                        checker = len(tags) - i + 1
                                        # print("checker -> ",(checker)," len(tags) - i + 1 ->",len(tags)," | ",i)
                                        # print("##")
                                        for v in range(checker)[1:]:
                                            # print("checker -> ",(checker)," v -> ",v)
                                            # print("set_options[i][:-2] = "+set_options[i][:-2]+" set_option[i+v] = "+set_options[i+v])
                                            
                                            if(set_options[i] == set_options[i+v]):
                                                # print("set_options[i][:-2] == set_options[",i,"+",v,"] -> "+set_options[i][:-2])
                                                if(len(row[i+v]) > 0):
                                                    # print("len(row[i+v]) > 0 -> BREAK" )
                                                    break
                                                else:
                                                    # print("len(row[i+v]) < 0 -> " + row[i] )
                                                    v=v+1
                                            else:
                                                # print("set_options[i] != set_options[",i,"+",v,"] SET_NAME -> "+set_options[i])
                                                # print("set_options[i] != set_options[",i,"+",v,"] SET_OPTIONS[i+v] -> "+set_options[i+v])
                                                # print("BREAK \n")
                                                xmlData.write('!\t\t' + '</setattribute>' + "\n")
                                                break
                                                
                                set_name = set_options[i]
                                
                            else:
                                xmlData.write('\t\t' + '<' + tags[i] + ' name="'+tag_options[i]+'">' + row[i] + '</' + tags[i] + '>' + "\n")
                        elif(categories[i] != categories[i+1]): 
                            if(len(row[i-1]) == 0):
                                if (categories[i] == 'Content Server Categories:K004 - Felülvizsgálat adatok') and (categories[i-1] == 'Content Server Categories:K004 - Felülvizsgálat adatok'):
                                    xmlData.write('	<category name="'+categories[i]+'">' + "\n")
                                    lastCategory = categories[i]


                            xmlData.write('\t\t' + '<' + tags[i] + ' name="'+tag_options[i]+'">' + row[i] + '</' + tags[i] + '>' + "\n")
                            xmlData.write('	</category>'+"\n")
                            xmlData.write("\n")
                        
                        else:
                            # print("len(set_options[i]): ",len(set_options))
                            # print("last setoptions: ",set_options[48])
                            # print("len(row[i]): ",len(row))
                            # lenrow = len(row)
                            # print("last row: ",row[len(row)-1])
                            # print("i: ",i)
                                          
                            if (len(set_options[i]) > 0):
                                if(set_name == set_options[i]):
                                    # print("0set_name: "+set_name+" --i:",i)
                                    if(lastCategory != categories[i]):
                                        xmlData.write('	<category name="'+categories[i]+'">' + "\n")
                                        lastCategory = categories[i]

                                    xmlData.write('\t\t\t' + '<' + tags[i] + ' name="'+tag_options[i]+'">' + row[i] + '</' + tags[i] + '>' + "\n")
                                    checker = len(tags) - i + 1
                                    for v in range(checker)[1:]:
                                        if(set_name == set_options[i+v]):
                                            if(len(row[i+v]) > 0):
                                                break
                                            else:
                                                v=v+1
                                        else:
                                            # print("1set_name: "+set_name+" set_options[i+v]: "+set_options[i+v]+" --v:",v," --i:",i)
                                            xmlData.write('\t\t' + '</setattribute>' + "\n")
                                            break
                                
                                else:
                                    if(i==first_set):
                                        # print("2set_name: "+set_name+" --i:",i)
                                        if(lastCategory != categories[i]):
                                            xmlData.write('	<category name="'+categories[i]+'">' + "\n")
                                            lastCategory = categories[i]
                                        xmlData.write('\t\t' + '<setattribute name="'+set_options[i][:-2]+'">' + "\n")
                                        xmlData.write('\t\t\t' + '<' + tags[i] + ' name="'+tag_options[i]+'">' + row[i] + '</' + tags[i] + '>' + "\n")
                                        
                                    elif(i==len(tags)):
                                        # print("3set_name: "+set_name+" --i:",i)
                                        xmlData.write('\t\t\t' + '<' + tags[i] + ' name="'+tag_options[i]+'">' + row[i] + '</' + tags[i] + '>' + "\n")
                                        xmlData.write('\t\t' + '</setattribute>' + "\n")
                                        
                                    else:
                                        # print("4set_name: "+set_name+" --i:",i)
                                        # print(row[i]+' --> ')
                                        if(lastCategory != categories[i]):
                                            xmlData.write('	<category name="'+categories[i]+'">' + "\n")
                                            lastCategory = categories[i]
                                        xmlData.write('\t\t' + '<setattribute name="'+set_options[i][:-2]+'">' + "\n")
                                        xmlData.write('\t\t\t' + '<' + tags[i] + ' name="'+tag_options[i]+'">' + row[i] + '</' + tags[i] + '>' + "\n")
                                        
                                        checker = len(tags) - i + 1
                                        # print("checker -> ",(checker)," len(tags) - i + 1 ->",len(tags)," | ",i)
                                        # print("##")
                                        for v in range(checker)[1:]:
                                            # print("checker -> ",(checker)," v -> ",v)
                                            # print("set_options[i] = "+set_options[i]+" set_option[i+v] = "+set_options[i+v])
                                            
                                            if(set_options[i] == set_options[i+v]):
                                                # print("set_options[i] == set_options[",i,"+",v,"] -> "+set_options[i])
                                                if(len(row[i+v]) > 0):
                                                    # print("len(row[i+v]) > 0 -> BREAK" )
                                                    break
                                                else:
                                                    # print("len(row[i+v]) < 0 -> " + row[i] )
                                                    v=v+1
                                            else:
                                                # print("set_options[i] != set_options[",i,"+",v,"] SET_NAME -> "+set_options[i])
                                                # print("set_options[i] != set_options[",i,"+",v,"] SET_OPTIONS[i+v] -> "+set_options[i+v])
                                                # print("BREAK \n")
                                                xmlData.write('!\t\t' + '</setattribute>' + "\n")
                                                break
                                                
                                set_name = set_options[i]
                                
                            else:
                                xmlData.write('\t\t' + '<' + tags[i] + ' name="'+tag_options[i]+'">' + row[i] + '</' + tags[i] + '>' + "\n")
                                
                    elif(i == len(categories)-2):
                        xmlData.write('	</category>'+"\n")

                    elif (len(row[i]) == 0): 
                        if (categories[i] != categories[i+1]):
                            xmlData.write('	</category>'+"\n")


                elif tags[i].find("location") != -1:
                    row[i] = row[i].replace('\\',':')
                    if len(row[i+1]) == 0:
                        f_name = row[i]
                        index = f_name.rfind(":")+1
                        # print(f_name[:index-1])
                        # print(f_name[index:])
                        xmlData.write('\t' + '<' + tags[i]+'>' + f_name[:index-1] + '</' + tags[i] + '>' + "\n")
                        xmlData.write('\t' + '<' + tags[i+1]+'>' + f_name[index:] + '</' + tags[i+1] + '>' + "\n")
                    else:
                        xmlData.write('\t' + '<' + tags[i] + ' name="'+tag_options[i]+'">' + row[i] + '</' + tags[i] + '>' + "\n")
                # elif tags[i].find("title") != -1:
                #     continue
                elif tags[i].find("mime") != -1:
                    if row[i] == 'PDF':
                        row[i] =  'application/pdf'
                    if row[i] == 'pdf':
                        row[i] =  'application/pdf'
                    if row[i] == 'doc':
                        row[i] =  'application/msword'
                    if row[i] == 'docx':
                        row[i] =  'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    
                    
                    xmlData.write('\t' + '<' + tags[i] + '>' + row[i] + '</' + tags[i] + '>' + "\n")
                    xmlData.write('\t' + '<versioncontrol>TRUE</versioncontrol>' + "\n")
                    xmlData.write('\t' + '<versiontype>MAJOR</versiontype>' + "\n")


                else:
                    if len(row[i]) != 0:
                        xmlData.write('\t' + '<' + tags[i]+'>' + row[i] + '</' + tags[i] + '>' + "\n")
            # xmlData.write('	</category>'+"\n")
            xmlData.write('</node>' + "\n")
            xmlData.write("\n")
        rowNum +=1

    xmlData.write('</import>' + "\n")
    xmlData.close()



if __name__ == "__main__":
    xlsx2csv()
    csv2xml()
