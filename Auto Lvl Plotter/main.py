# Program created By Andrew Pilimai Tulua
# Program created August, 2019
""" The purpose for this program is to take the data from an
auto level survey and convert it into plottable points for GIS software.
This program requires the data to be in a specific .xlsx format

(available from Andrew Pilimai Tulua). For more information please contact me
at pilimai10@gmail.com """

""" importing all neccessary modules for the program """
from tkinter import *
from tkinter import filedialog
import pandas as pd
import math

#These are all the external modules needed to run the program
fsdata = {}
bsdata = {}
rfpoint = []
bspoint = []

#Main window and its labels, buttons, properties etc
window = Tk()
window.title('auto-level converter')

#Modules for all buttons on main window (window)

#Import window to select data file and convert to dictionary
def importfn():
    global filelbl
    filelbl = filedialog.askopenfilename()
    importwind = Tk()
    df = pd.read_excel(filelbl)
    dataprevlbl = Label(importwind , text=df.columns)
    
    shotid = df['ID']
    deg = df['DEG']
    dist = df['DIST']
    elev = df['ELEV']
    
    rfx = df['X'].values[0]
    rfy = df['Y'].values[0]
    rfpoint.append(rfx)
    rfpoint.append(rfy)
    
    BSdeg = df['bsdeg']
    BSdist = df['bsdist']
    bsdata.update({0:[BSdeg[x] , BSdist[x] , 0] for x in range(1)})
    
    fsdata.update({shotid[x]:[deg[x],dist[x],elev[x]] for x in range(len(shotid))})
    calclbl = Label(importwind , text=fsdata )
    bslbl = Label(importwind , text=bsdata )
    reflbl = Label(importwind , text=rfpoint )
    #GUI Layout
    dataprevlbl.grid(column=0, row=0)
    calclbl.grid(column=0 , row=1)
    bslbl.grid(column=0 , row=2)
    reflbl.grid(column=0 , row=3)
    importwind.mainloop()

def reffn():                                                                  #Window for selecting reference point
    refwind = Tk()
    def xcelfn():                                                             #Window for entering reference point for the xcel sheet
        xlrefwin = Tk()
        xlimport = Label(xlrefwin , text=rfpoint)
        xlimport.grid(column=0 , row=0)
    def manualentry1():
        manualwind = Tk()                                                   #Window for manually entering the reference point
        northlbl = Label(manualwind , text='Northing:')
        northentry = Entry(manualwind , width=15)
        eastlbl = Label(manualwind , text='Easting:')
        eastentry = Entry(manualwind , width=15)
        def importxy():                                                       #Saving the entries into refpoint list (x,y) or (Easting , Norhting)
            rfpoint.append(eastentry)
            rfpoint.append(northentry)
            print(rfpoint)
        enterbtn = Button(manualwind , text='Enter' , command = importxy)
        #GUI Layout
        northlbl.grid(column=0 , row=1)
        northentry.grid(column=1 , row=1)
        eastlbl.grid(column=0 , row=2)
        eastentry.grid(column=1 , row=2)
        enterbtn.grid(column=1 , row=3)
        manualwind.mainloop()
    xcellbl = Label(refwind , text='Use Reference Point from Xcel sheet')
    xcelbtn = Button(refwind , text='Select' , command=xcelfn)
    manuallbl = Label(refwind , text='Manually Enter Reference Point')
    manualbtn = Button(refwind , text='Select' , command=manualentry1)
    #GUI Layout
    xcellbl.grid(column=0 , row=0)
    xcelbtn.grid(column=1 , row=0)
    manuallbl.grid(column=0 , row=1)
    manualbtn.grid(column=1 , row=1)
    refwind.mainloop()
    
def bsfn():                                                                   #Window for SElecting the Back Sight
    bswind = Tk()
    def xcelfn():                                                             #Window for importing from the xcel sheel
        xlbswin = Tk()
        xlimport = Label(xlbswin , text=bsdata)
        xlimport.grid(column=0 , row=0)
    def manualentry2():                                                         #Window for manually entering the Back SIGHT
        manualwindbs = Tk()
        def enterfn():                                                        #Converts the entery data into the dictionary for Back Sights
            deg = degentry
            dist = distentry
            elev = eleventry
            bsdata.update({0: [deg , dist , elev]})
        deglbl = Label(manualwindbs , text='Degree:')
        degentry = Entry(manualwindbs , width=15)
        distlbl = Label(manualwindbs , text='Distance:')
        distentry = Entry(manualwindbs , width=15)
        elevlbl = Label(manualwindbs , text='Elevation')
        eleventry = Entry(manualwindbs , width=15)
        enterbtn = Button(manualwindbs , text='Enter' , command = enterfn)
        #GUI layout
        deglbl.grid(column=0 , row=0)
        degentry.grid(column=1 , row=0)
        distlbl.grid(column=0 , row=1)
        distentry.grid(column=1 , row=1)
        elevlbl.grid(column=0 , row=2)
        eleventry.grid(column=1 , row=2)
        enterbtn.grid(column=1 , row=3)
        manualwindbs.mainloop()
    xcellbl = Label(bswind , text='Use Back Sight from Xcel sheet')
    xcelbtn = Button(bswind , text='Select' , command = xcelfn)
    manuallbl = Label(bswind , text='Manually Enter Back Sight')
    manualbtnbs = Button(bswind , text='Select' , command = manualentry2)
    #GUI Layout
    xcellbl.grid(column=0 , row=0)
    xcelbtn.grid(column=1 , row=0)
    manuallbl.grid(column=0 , row=1)
    manualbtnbs.grid(column=1 , row=1)
    bswind.mainloop()
    
def mathmod(points):      #This is the module that will convert the [deg,dist,elev] list into the full list [deg,dist,elev,group,x,y,e,n]
    #assigning groups
    for k,v in points.items():
        if v[0]<90:
            v.append(1)
        elif 90 < v[0] < 180:
            v.append(2)
        elif 180 < v[0] < 270:
            v.append(3)
        elif 270 < v[0] < 360:
            v.append(4)
    #converting data to coordinate values
    for k,v in points.items():
        if v[3]==1:
            x=(math.sin(math.radians(v[0])) * v[1])
            y=(math.cos(math.radians(v[0])) * v[1])
            v.append(x)
            v.append(y)
            E = (x * 0.0001) / 7.871
            N = (y * 0.0001) / 11.132
            v.append(E)
            v.append(N)
        elif v[3]==2:
            x=(math.cos(math.radians(v[0]-90)) * v[1])
            y=(math.sin(math.radians(v[0]-90)) * v[1])
            v.append(x)
            v.append(y)
            E = (x * 0.0001) / 7.871
            N = (y * 0.0001) / 11.132
            v.append(E)
            v.append(N)
        elif v[3]==3:
            x=(math.sin(math.radians(v[0]-180)) * v[1])
            y=(math.cos(math.radians(v[0]-180)) * v[1])
            v.append(x)
            v.append(y)
            E = x * (0.0001 / 7.871)
            N = y * (0.0001 / 11.132)
            v.append(E)
            v.append(N)
        elif v[3]==4:
            x=(math.cos(math.radians(v[0]-270)) * v[1])
            y=(math.sin(math.radians(v[0]-270)) * v[1])
            v.append(x)
            v.append(y)
            E = (x * 0.0001) / 7.871
            N = (y * 0.0001) / 11.132
            v.append(E)
            v.append(N)
            
def convertdata():  #Window for running the converter and saving the new data
    mathmod(bsdata)
    mathmod(fsdata)
    for k,v in bsdata.items():
        if v[3]==1:
            bspoint.append(rfpoint[0]-v[6]) # BS E
            bspoint.append(rfpoint[1]-v[7]) # BS N
        if v[3]==2:
            bspoint.append(rfpoint[0]-v[6]) # BS E
            bspoint.append(rfpoint[1]+v[7]) # BS N
        if v[3]==3:
            bspoint.append(rfpoint[0]+v[6]) # BS E
            bspoint.append(rfpoint[1]+v[7]) # BS N
        if v[3]==4:
            bspoint.append(rfpoint[0]+v[6]) # BS E
            bspoint.append(rfpoint[1]-v[7]) # BS N
    for k,v in fsdata.items():
        if v[3]==1:
            v.append(bspoint[0]+v[6]) #E
            v.append(bspoint[1]+v[7]) #N
        elif v[3]==2:
            v.append(bspoint[0]+v[6])
            v.append(bspoint[1]-v[7])
        elif v[3]==3:
            v.append(bspoint[0]-v[6])
            v.append(bspoint[1]-v[7])
        elif v[3]==4:
            v.append(bspoint[0]-v[6])
            v.append(bspoint[1]+v[7])
    for k,v in fsdata.items():
        finaldata = {x:[v[9],v[8]] for x in range(len(fsdata))}
    resultswind = Tk()
    resultslbl = Label(resultswind , text=finaldata)
    resultslbl.grid(column=0 , row=0)
    finalxl = pd.DataFrame(data=finaldata)
    finalxl = (finalxl.T)
    finalxl.to_excel('converted.xlsx')

#GUI Stuff for the main window and its buttons
importlbl = Label(window , text='Import Data')
reflbl = Label(window , text='Select Reference Point')
bslbl = Label(window , text='Select Back Sight')
convertlbl = Label(window , text='Convert Data')
importbtn = Button(window , text='Import' , command = importfn)
refbtn = Button(window , text='Configure' , command = reffn)
bsbtn = Button(window , text='Configure' , command = bsfn)
convertbtn = Button(window , text='Run' , command = convertdata)

#Main window layout
importlbl.grid(column=0 , row=0 )
reflbl.grid(column=0 , row=1 )
bslbl.grid(column=0 , row=2 )
convertlbl.grid(column=0 , row=3 )
importbtn.grid(column=1 , row=0 )
refbtn.grid(column=1 , row=1 )
bsbtn.grid(column=1 , row=2 )
convertbtn.grid(column=1 , row=3 )
window.mainloop()