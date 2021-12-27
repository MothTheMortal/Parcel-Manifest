from tkinter import *
import os
import sys
from datetime import date

from docxtpl import DocxTemplate

os.chdir(sys.path[0])
today = date.today()
a = str(date.today())
doc = DocxTemplate('ParcelManifest.docx')
Height = 500
Width = 1000
SLIST = []
JLIST = []
PLIST = []
DLIST = []
UNKNOWN = []

root = Tk()
root.title("Shopee Script")
canvas = Canvas(root, height=Height, width=Width)
canvas.pack()

background_image = PhotoImage(file='./bg1.png')
background_label = Label(root, image=background_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

e = Entry(root, width=35, borderwidth=5)
e.place(x=380, y=50)


def restart():
    os.startfile(__file__)
    sys.exit()


def get_parcels():
    result = e.get()
    e.delete(0, END)

    s = len(SLIST)
    j = len(JLIST)
    p = len(PLIST)
    d = len(DLIST)
    number_of_parcel = s + j + p + d

    def shopee():
        SLIST_A = str(SLIST1)[1:-1]

        if s > 48:
            length = len(SLIST_A)
            middle_index = length // 2
            first_half = SLIST_A[:middle_index]
            second_half = SLIST_A[middle_index:]
            context = {'id': first_half, 'id2': second_half, 'date': today, 'quantity': s, 'logistic': "Shopee Express"}
            doc.render(context)
            doc.save('Shopee-' + e2.get() + "_" + a + '.docx')
        else:
            context = {'id': SLIST_A, 'date': today, 'quantity': s, 'logistic': "Shopee Express"}
            doc.render(context)
            doc.save('Shopee-' + e2.get() + "_" + a + '.docx')

    def Jnt():
        JLIST_A = str(JLIST1)[1:-1]
        context = {'id': JLIST_A, 'date': today, 'quantity': j, 'logistic': "JnT Express"}
        doc.render(context)
        doc.save('J&T-' + e2.get() + "_" + a + '.docx')

    def Poslaju():
        PLIST_A = str(PLIST1)[1:-1]
        context = {'id': PLIST_A, 'date': today, 'quantity': p, 'logistic': "Poslaju"}
        doc.render(context)
        doc.save('Poslaju-' + e2.get() + "_" + a + '.docx')

    def DHL():
        DLIST_A = str(DLIST1)[1:-1]
        context = {'id': DLIST_A, 'date': today, 'quantity': d, 'logistic': "DHL"}
        doc.render(context)
        doc.save('DHL-' + e2.get() + "_" + a + '.docx')

    def pick():
        e2.delete(0, END)
        e2.insert(0, "Pickup")

    def drop():
        e2.delete(0, END)
        e2.insert(0, "Dropoff")

    if "done" == result:
        e2 = Entry(root, width=32, borderwidth=5, font=50)
        e2.place(x=350, y=350)

        SLIST1 = list(dict.fromkeys(SLIST))
        JLIST1 = list(dict.fromkeys(JLIST))
        PLIST1 = list(dict.fromkeys(PLIST))
        DLIST1 = list(dict.fromkeys(DLIST))
        s = len(SLIST1)
        j = len(JLIST1)
        p = len(PLIST1)
        d = len(DLIST1)
        button2 = Button(root, text="Shopee", font=40, command=shopee)
        button3 = Button(root, text="J&T", font=40, command=Jnt)
        button4 = Button(root, text="Poslaju", font=40, command=Poslaju)
        button5 = Button(root, text="DHL", font=40, command=DHL)

        button2.place(anchor="n", x=30, y=200)
        button3.place(anchor="n", x=21, y=230, width=87)
        button4.place(anchor="n", x=30, y=260, width=70)
        button5.place(anchor="n", x=23, y=290, width=83)
        label3 = Label(root, text="Print", width=7, height=1, font="italic")
        label3.place(anchor='n', x=30, y=170)
        button6 = Button(root, text="Dropoff", font=40, command=drop)
        button7 = Button(root, text="Pickup", font=40, command=pick)
        button6.place(x=350, y=390)
        button7.place(x=590, y=390)
        number_of_parcel = s + j + p + d
        a_string = "Total number of parcels = " + str(number_of_parcel)
        format_parcel = a_string.format(1)
        a_string2 = "Number of parcels" + "\nShopee = " + str(s) + "\nJ&T = " + str(j) + "\nPoslaju = " + str(
            p) + "\nDHL = " + str(d)
        format2_parcel = a_string2.format(1)

        print("Shopee", SLIST1, '\n', "J&T", JLIST1, '\n', "POSLAJU", PLIST1, '\n', "DHL", DLIST1, '\n',
              "Unknown Tracking number", UNKNOWN)
        label['text'] = a_string2
        label2['text'] = format_parcel
    elif "SPXMY" in result:
        SLIST.append(result)
    elif "620000" in result:
        JLIST.append(result)
    elif "PL" in result:
        PLIST.append(result)
    elif "59201" in result:
        DLIST.append(result)
    else:
        UNKNOWN.append(result)


def func(event):
    get_parcels()


root.bind('<Return>', func)


def done1():
    e.insert(0, "done")
    get_parcels()


button8 = Button(root, text="Done", font=40, command=done1)
button8.place(x=630, y=47)
button1 = Button(root, text="Continue (Press Done, if you're done with scanning)", command=get_parcels, font=40)

button1.place(x=320, y=102)

lower_frame = Frame(root, bg='#80c1ff', bd=10, width=300, height=180)
lower_frame.place(x=350, y=160)

label = Label(lower_frame, bg="#80c1ff", fg="#F0F0F0", font='Helvetica 12 bold')
label.place(x=60, y=45)
label2 = Label(lower_frame, bg="#80c1ff", fg="#F0F0F0", font='Helvetica 12 bold')
label2.place(x=40, y=20)

button9 = Button(root, text="Restart", command=restart, width=8, height=3)
button9.place(x=1, y=1)


root.mainloop()
