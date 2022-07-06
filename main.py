import pikepdf
import PyPDF2 as pypdf
import os
import pathlib
import xlsxwriter as xlw
from tkinter import Tk
from tkinter import filedialog

#guide
print("Please select the folder containing the forms ONLY")
#GUI to ask for folder in which forms are saved
rootBox = Tk()
rootBox.withdraw()

path = filedialog.askdirectory()

try:
    #initializing variables and excel sheet
    print("\nWorking....\n")
    i = 1 #i is row count
    workbook = xlw.Workbook(path + r'\demo.xlsx')
    worksheet = workbook.add_worksheet()
    
    #guide
    print("You can minimize this window and work on something else!\n")
    
    #initializing column widths
    worksheet.set_column(0, 1, 30)
    worksheet.set_column(2, 2, 70)

    #initializing column names
    worksheet.write(0, 0, 'NAME')
    worksheet.write(0, 1, 'NUMBER')
    worksheet.write(0, 2, 'ADDERESS')

    #iterating through files
    pathObject = pathlib.Path(path)
    for filePath in pathObject.iterdir():

        #guide
        print("Working on file number: ", i)
        
        #decrypting file
        fileObject = pikepdf.open(filePath,password='')
        fileObject.save(path + r'\temp.pdf')

        #reading data
        pdfObject=open(path + r'\temp.pdf', 'rb')
        pdf = pypdf.PdfFileReader(pdfObject)
        dictionary = pdf.getFormTextFields()

        #writing data to excel
        worksheet.write(i, 0, dictionary['db_aplcnt_name'])
        worksheet.write(i, 1, dictionary['db_aplcnt_phone'])
        worksheet.write(i, 2, dictionary['db_aplcnt_address_1'])
        
        #incrementing i (i.e row count)
        i += 1

        #close the temp file
        pdfObject.close()

    #closing excel sheet
    workbook.close()

    try :
        os.remove(path +r'\temp.pdf')
    except:
        pass

    #guide
    print("\nDONE!!!")


#error handling
except:
    try :
        workbook.close()
    except:
        pass
    print("\n\n!!!! ERROR OCCURRED !!!!\n\n")
    print("Check the below points to try to resolve the issue: ")
    print("1. Please check if you had provided the correct folder")
    print("2. Please check if the provided folder contains ONLY the forms in pdf format and no other file")
    print("3. Please restart your PC once if the above points have been checked but the program is still not working")
    print("4. Please contact the developer if the issue has not been solved through the above points")
        
#remove temp file 
try :
    os.remove(path +r'\temp.pdf')
except:
    pass