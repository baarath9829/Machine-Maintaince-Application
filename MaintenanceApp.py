import TkinterClass
from Tkinter import *
import openpyxl
import datetime
locDirectory = {}
statusDirectory = {}
tabnames = []
f = open("setting.txt","r")
filename = f.readline().strip("\n")
noOfTabs = int(f.readline())
for i in range(noOfTabs):
    tabnames.append(f.readline().strip("\n"))
def describe(code):
    displayFrame = TkinterClass.frame(main.Window)
    displayFrame.frame.grid_propagate(0)
    rowno = locDirectory[code]
    Sno = barcode["A"+str(rowno)].value
    description = barcode["C"+str(rowno)].value
    mcsno = barcode["D"+str(rowno)].value
    model = barcode["E"+str(rowno)].value
    make = barcode["F"+str(rowno)].value
    quantity = barcode["G"+str(rowno)].value
    location = barcode["H"+str(rowno)].value
    PMfreq = barcode["J"+str(rowno)].value
    remark = barcode["K"+str(rowno)].value
    capacity = barcode["L"+str(rowno)].value
    serviced = barcode["M"+str(rowno)].value
    displayFrame.display("Sno   :",str(Sno))
    displayFrame.moverow(1)
    displayFrame.display("barcode   :",str(code))
    displayFrame.moverow(1)
    displayFrame.display("Description   :",str(description))
    displayFrame.moverow(1)
    displayFrame.display("MCS SLNO  :",str(mcsno))
    displayFrame.moverow(1)
    displayFrame.display("Model     :",str(model))
    displayFrame.moverow(1)
    displayFrame.display("Make  :",str(make))
    displayFrame.moverow(1)
    displayFrame.display("Quantity  :",str(quantity))
    displayFrame.moverow(1)
    displayFrame.display("Location  :",str(location))
    displayFrame.moverow(1)
    displayFrame.display("PM Frequency  :",str(PMfreq))
    displayFrame.moverow(1)
    displayFrame.display("Remark    :",str(remark))
    displayFrame.moverow(1)
    displayFrame.display("Capacity  :",str(capacity))
    displayFrame.moverow(1)
    displayFrame.display("Last Service  :",str(serviced))
    displayFrame.moverow(1)
    displayFrame.nextbutton("back",outframe)

    
def submit():
    for i in inframe.variable.iterkeys():
        statusDirectory[i] = inframe.variable[i].get()
    for i in statusDirectory.iterkeys():
        if (statusDirectory[i] == 1):
            print (i)
            barcode["M"+str(locDirectory[i])] = str(datetime.date.today())
    print ("submited")
    wb.save(filename)#Master List of Machinery 01-08-18.xlsx
    main.Window.destroy()
    
def configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

main = TkinterClass.window()

#main part
outframe = TkinterClass.frame(main.Window)
outframe.frame.grid_propagate(0)

canvas = Canvas(outframe.frame,width=350,height=350)
inframe=TkinterClass.frame(canvas)

scrollbar=Scrollbar(outframe.frame,orient="vertical",command=canvas.yview)
canvas.configure(yscrollcommand=scrollbar.set)

scrollbar.pack(side="right",fill="y")
canvas.pack(fill="both",expand=1)
canvas.create_window((0,0),window=inframe.frame,anchor='nw')
inframe.frame.bind("<Configure>",configure)

wb = openpyxl.load_workbook(filename) #Master List of Machinery 01-08-18.xlsx
for tab in tabnames:
    barcode = wb[tab]
    for i  in range(8,1000):
        code = barcode["I" + str(i)].value
        if code is None:
            scell=barcode["H" + str(i)].value
            if scell is None:
                continue
        servicedOn = barcode["M" + str(i)].value
        if servicedOn is not None:
            l = servicedOn.split("-")
            old = datetime.date(int(l[0]),int(l[1]),int(l[2]))
            new = datetime.date.today()
            if (new==old):
                continue
            days = int(str(new-old).split("day")[0])
            if (days < 30):
                continue
        code=str(code)
        print ("code:"+code)
        locDirectory[code] = str(i)
        if (i == 1000):
            inframe.seekrow(0)
            inframe.seekcolumn(inframe.column + 2)
        inframe.checkbox(code,code)
        inframe.movecolumn(1)
        inframe.button("details",describe,code)
        inframe.moverow(1)
print (inframe.row)
button = Button(outframe.frame,text="submit",command= lambda: submit())
button.pack()

outframe.show()
inframe.show()
main.show()
