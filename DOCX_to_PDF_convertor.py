from docx2pdf import *
from tkinter import *
from tkinter import filedialog
import re
import PyPDF2
import pandas as pd



global numlist
numlist = []

global maillist
maillist = []

global locationlist
locationlist=[]





def browseFiles():
    global filename
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select a File",
                                          filetypes=(("Text files",
                                                      "*.txt*"),
                                                     ("All Files",
                                                      "*.*")))


    label_file_explorer.configure(text="File Opened: " + filename)




def convertFile():

    convert(filename)
    locationlist.append(filename)
    data['Location'] = locationlist
    data.to_excel('CV.xlsx')





def mobNum():

    pdfFileObj = open(filename, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    print("Number of pages:-" + str(pdfReader.numPages))
    num = pdfReader.numPages
    i = 0
    while (i < num):
        pageObj = pdfReader.getPage(i)
        text = pageObj.extractText()
        text1 = text.lower()
        for line in text1:
            numm = re.findall('[0-9]+', text1)
            for o in numm:
                if len(o) == 10:
                    print(o)
                    numlist.append(o)
                    print(numlist)
                    data['Number'] = numlist
                    data.to_excel('CV.xlsx')

            break
        i = i + 1





def getEmail():

    pdfFileObj = open(filename, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    print("Number of pages:-" + str(pdfReader.numPages))
    num = pdfReader.numPages
    i = 0
    while (i < num):
        pageObj = pdfReader.getPage(i)
        text = pageObj.extractText()
        text1 = text.lower()
        for line in text1:
            numm = re.findall(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+" , text1)
            for k in numm:
                if len(k)>10:
                    print(k)
                    maillist.append(k)
                    data['EMAIL ID'] = maillist
                    data.to_excel('CV.xlsx')
            break
        i = i + 1

    maillist.append(k)






data = pd.DataFrame()

data.to_excel('CV.xlsx')






window = Tk()


window.title('BY ARJUN NAYYAR')


window.geometry("1450x300")


window.config(background="grey")


label_file_explorer = Label(window,
                            text="FILE CONVERTER FOR OST PLACEMENT",
                            width=100, height=4,
                            fg="black",
                            font="Helvetica 18 bold",
                            bg="grey")

button_explore = Button(window,
                        text="Browse Files",
                        command=browseFiles)

button_exit = Button(window,
                     text="Exit",
                     command=exit)

button_convert = Button(window, text="Convert", command = convertFile)

button_mobile = Button(window, text="Extract Mobile Number", command = mobNum)

button_email = Button(window, text="Extract Email ID", command = getEmail)


label_file_explorer.grid(column=1, row=1)

button_explore.grid(column=1, row=2)

button_convert.grid(column=1, row=3)

button_mobile.grid(column=1, row=4)

button_email.grid(column=1, row=5)

button_exit.grid(column=1, row=6)

window.mainloop()
