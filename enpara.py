from datetime import date
import os
import subprocess
import sys
import camelot
import ctypes
from numpy import append
import numpy as np
import pandas as pd
import xlwt
from ctypes.util import find_library
from dateutil.parser import parse
import xlsxwriter
find_library("".join(("gsdll", str(ctypes.sizeof(ctypes.c_voidp) * 8), ".dll")))

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

# the final tables
tables = []

dtypes = np.dtype(
    [
        ('Islem tarihi', np.datetime64),
        ("Aciklama", str),
        ("Taksit", str),
        ("Tutar (TL)", float),
    ]
)
df = pd.DataFrame(np.empty(0, dtype=dtypes))
 

pdflist = sorted(os.listdir(os.path.dirname(sys.argv[0])))
for pdf in pdflist:
    if ".pdf" in pdf:
        print ("processing file :" + pdf)
        # we check for tables with both methods
        lattice_tables = camelot.read_pdf(pdf, flavor='lattice', pages='1-end' , encoding='IS0-8859-9')
        stream_tables = camelot.read_pdf(pdf, flavor='stream', pages='1-end' , encoding='IS0-8859-9')

        # we check whether the number of tables are the same
        if len(lattice_tables) > 0 and len(stream_tables) > 0 and len(lattice_tables) == len(stream_tables):       
            # then we try to pick the best table
            for index in range(len(lattice_tables)):
                # we check whether the tables are both good enough
                if is_good_enough(lattice_tables[index].parsing_report) and is_good_enough(stream_tables[index].parsing_report):        
                    # they probably represent the same table
                    if lattice_tables[index].parsing_report['page'] == stream_tables[index].parsing_report['page'] and lattice_tables[index].parsing_report['order'] == stream_tables[index].parsing_report['order']:
                        total_lattice = 1
                        total_stream = 1

                        for num in lattice_tables[index].shape:
                            total_lattice *= num
                        for num in stream_tables[index].shape:
                            total_stream *= num

                        # we pick the table with the most cells
                        if(total_lattice >= total_stream):
                            tables.append(lattice_tables[index])
                        else:
                            tables.append(stream_tables[index])

                # if we have a different number of tables we just pick the object with the most number of tables
                elif is_good_enough(lattice_tables[index].parsing_report):
                    tables.append(lattice_tables[index])

                elif is_good_enough(stream_tables[index].parsing_report):
                    tables.append(stream_tables[index])
        elif len(lattice_tables) >= len(stream_tables):
            tables = lattice_tables
        else:
            tables = stream_tables    

        if tables is not None and len(tables) > 0:
            # let's check whether is TableList object or just a list of tables
            if isinstance(tables, camelot.core.TableList) is False:
                tables = camelot.core.TableList(tables)

            # uptrim = False
            for i in range(len(tables._tables)):
                print ("processing table :" + str(i+1) +"/"+ str(len(tables._tables)))
                for row in tables[i].data:
                #     if "numaralı sanal kredi kartınızla yapılan" in row[1]: 
                #         uptrim=True
                #     if(uptrim):
                    if is_date(row[0]):
                        #row[1] = row[1].replace("(cid:12)","Ö").replace("(cid:6)","ç").replace("(cid:10)","İ").replace("(cid:7)","ö").replace("(cid:11)","ş").replace("(cid:8)","Ö").replace("(cid:17)","ğ").replace("(cid:15)","Ü")        
                        row[-1] = row[-1].replace(".","").replace(",",".").replace(" TL","").replace(" ","")
                        df = df.append({'Islem tarihi': row[0], 'Aciklama': row[1], 'Taksit':row[-2], 'Tutar (TL)':row[-1]}, ignore_index = True)

print(df)
df.to_excel("enpara.xlsx") 