import xml.etree.ElementTree as ET
import tkinter
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
from  tkinter import filedialog
from tkinter import simpledialog
import openpyxl
import win32com.client as win32
import datetime

#Open csv file
#need to find a way to create this file and make name match wo number as input by user
#need to create a folder on desktop where all these files go but not overwrite existing folder/files

class gui:
    def selectEvent():  #User selects file to parse
        path = Tk()
        path.filename =  filedialog.askopenfilename(initialdir = "Path Where the dialog should open first",title = "Select Event File",filetypes = (("xml files","*.xml"),("all files","*.*")))
        print (path.filename)
        path.withdraw()
        global root
        tree = ET.parse(path.filename)#this is where the file location to be parsed goes
        root = tree.getroot()
        
    def callWO(): #This opens an input text gui for user to input wo number and converts it to the variable woCell for later use.
        root = Tk()
        root.withdraw()
        woCell = tkinter.simpledialog.askstring("Enter WO #", "Enter WO Number")
        print (woCell)

#this is a gui for user to select which template to use and variable selected will be used to open an existing excel template that will be included in program folder.
    def woTemplate():
                
        master = Tk()
        variable = StringVar(master)
        variable.set("Select the main fault for billable event") # default value
        w = OptionMenu(master, variable, "Out of fuel", "Dead battery from misuse", "E-stop", "Frozen", "Damage", "E-brake unplugged", "OI unplugged")
        w.pack()

        def ok():
            print ('Billable event due to ' +  variable.get()) #gui.
            master.withdraw()

        button = Button(master, text="OK", command=ok)
        button.pack()
        master.mainloop()

    def truckNumber(): #Stores user input variable truckCell for later use
        root = Tk()
        root.withdraw()
        truckCell = tkinter.simpledialog.askstring("Enter Truck  Number", "Enter Truck number")
        print (truckCell)

    def repairTimeS(): #Stores user input variable repairSdate for later use
        root = Tk()
        root.withdraw()
        repairSdate = tkinter.simpledialog.askstring("Enter Truck  Number", "Enter Truck number")
        print (repairSdate)

gui.selectEvent()                
gui.callWO()
gui.woTemplate()
gui.truckNumber()

for data in root:
    dataDict =root.attrib
    fchoursCell = (dataDict['FuelCellHours'])
    systemCell = (dataDict['SerialNumber'])
    break

for child in root:
    alarmDict = child.attrib      #This lists the alarm codes as dicts
    alarm_list = ['Out of Fuel', 'System Min Bat No Fc', 'alarm'] 
    for alarm in alarm_list:
        if alarm in alarmDict.values():
            print ('contains this alarm', alarm)
        else:
            print ('No valid operator error alarm found in event file')
            break

# this script updates excel cells.
excel = win32.gencache.EnsureDispatch('Excel.Application')
#Change this location to match the file template name as long as same folder. IE. E-stop.xlxs
wb = excel.Workbooks.Open('C:\\Users\\PlugService\\Desktop\\testy.xlsx')
# wb = excel.Workbooks.Open(r'C:\myfiles\excel\workbook2.xlsx')
ws = wb.Worksheets('Copy1')

excel.Visible = False

#TODO make customerCell and locationCell and nameCell, signatureCell not have to be entered every time.  Some way for it to hold the data unless changed by user.

dateCell = datetime.date.today()
occuredCell = #TODO autofill from xml data, must check against alarm template to specificy what time to select from the xmlfile.  IE.  OOF occured at 1500 , program checks for the string 'out of fuel' in xml data and enters that time as variable
repairSdate = #TODO create GUI for entering repair start time
repairCdate = #TODO create gui for entering repair complete time
laborCell = #TODO creat GUI for labor hour input
partCost = #TODO make this total all parts cost
fillableform_for_parts = #TODO figure out how to make a fillable form for all parts entered
 
#This is the cell formating  ws.Cells(row , column)
ws.Cells(6, 3).Value = customerCell
ws.Cells(7, 3).Value = locationCell
ws.Cells(6, 9).Value = woCell  
ws.Cells(6, 12).Values = nameCell 
ws.Cells(9, 3).Value = systemCell 
ws.Cells(10, 3).Value = truckCell 
ws.Cells(11, 3).Value = fchoursCell 
ws.Cells(12, 3).Value = dateCell 
ws.Cells(8, 7).Values = occuredCell
ws.Cells(8, 10).Values = repairSdate 
ws.Cells(8, 13).Values = repairCdate 
ws.Cells(31, 15).Values = laborCell 
ws.Cells(31, 17).Values = partCostCell  
ws.Cells(32, 15).Values = laborCell + partsCostCell
#todo add all the cells for the parts consumed form
ws.Cells(34, 1).Values = 'Technician signature:' + signatureCell 

#TODO make saveas a variable location in gui.savefile funtion
wb.SaveAs('C:\\Users\\PlugService\\Desktop\\testyUpdate.xlsx')

excel.Application.Quit()
