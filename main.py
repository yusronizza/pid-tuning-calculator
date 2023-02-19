from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showinfo
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import tkinter as tk
from tabulate import tabulate

window = tk.Tk()
bg = "#1c1c1c"
filepath = ""

def onProgress():
    showinfo(
        title='Development Process',
        message='Dalam Pengembangan'
    )

def createMenu():
    #Main Menu bar
    menubar = tk.Menu(window)
    window.config(menu=menubar, bg=bg)
    #Create a File menu bar
    file_menu = tk.Menu(menubar, tearoff=False)
    #Add menu item to File menu and Add File menu to menubar
    file_menu.add_command(
        label='Import',
        command=importFile
    )
    file_menu.add_separator()
    file_menu.add_command(
        label='Exit',
        command=window.destroy
    )
    menubar.add_cascade(
        label="File",
        menu=file_menu
    )
    #Create Mode Menu
    mode_menu = tk.Menu(menubar, tearoff=False)
    mode_menu.add_command(
        label='Zeagler Nichols',
        command=onProgress
    )
    mode_menu.add_command(
        label='Cohen Coen',
        command=onProgress
    )
    mode_menu.add_command(
        label='Lambda Method',
        command=onProgress
    )
    menubar.add_cascade(
        label='Methods',
        menu=mode_menu
    )
    #Create an Help menu
    help_menu = tk.Menu(menubar, tearoff=False)
    help_menu.add_command(
        label='Tutorials',
        command=onProgress
    )
    file_menu.add_separator()
    help_menu.add_command(
        label='Features',
        command=onProgress
    )
    menubar.add_cascade(
        label="Help",
        menu=help_menu,
        command=onProgress
    )
    #Create an About menu
    about_menu = tk.Menu(menubar, tearoff=False)
    menubar.add_cascade(
        label="About",
        menu=about_menu
    )

    cpy_menu = tk.Menu(menubar, tearoff=False)
    menubar.add_cascade(
        label="Copyright",
        menu=cpy_menu
    )

def calculatePIController(APD, MLD, inPVP1, inPVP2, inPVT1, inPVT2, finPVP1, finPVP2, finPVT1, finPVT2, inOP, finOP, PVChangeTime, OPChangeTime):
    print("Calculating...")
    inSlope = (inPVP2-inPVP1)/(inPVT2-inPVT1)
    finSlope = (finPVP2-finPVP1)/(finPVT2-finPVT1)
    Td = abs(PVChangeTime-OPChangeTime)
    Kp = abs((finSlope-inSlope)/(finOP-inOP))
    Lambda = (2*APD)/(Kp*MLD)
    integralTime = 2*Lambda+Td
    controllerGain = integralTime/(abs(Kp)*((Lambda+Td)**2))
    print("Calculated")
    return inSlope, finSlope, Td, Kp, Lambda, integralTime, controllerGain

def onclick(event, entry, mode):
    print("Point Clicked")
    thisline = event.artist
    xdata = thisline.get_xdata()
    ydata = thisline.get_ydata()
    ind = event.ind
    points = tuple(zip(xdata[ind], ydata[ind]))
    entry.delete(0, 100)
    if mode == "PV":
        points_arr = [points[0][0], points[0][1]]
        entry.insert(0, str(points_arr[0]) + ", " + str(points_arr[1]))
    elif mode == "OP":
        points_arr = [points[0][1]]
        entry.insert(0, str(points_arr[0]) + ", " + str(points_arr[1]))
    elif mode == "Y":
        points_arr = [points[0][1]]
        entry.insert(0, str(points_arr[0]))
    else:
        points_arr = [points[0][0]]
        entry.insert(0, str(points_arr[0]))

def save_data(res):
    file = open("saved_file.txt", "w+")
    file.write(res)
    file.close()
    print("Data saved")

def load_data(data):
    return data

def importFile():
    print("Import files function executed")
    global filepath
    filetypes = (
        ('Excel', '*.xlsx'),
        ('Excel 97-2003', '*.xls'),
        ('CSV', '*.csv')
    )
    filepath = askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes
    )
    if filepath == '':
        showinfo(
            title='Selected File',
            message='Belum ada file dipilih'
        )
    else:
            showinfo(
            title='Selected File',
            message=filepath
        )
    print("Import Done")

def calculateMultiple(res):
    return res

def popupError(value):
    print("Pop up error function executed")
    showinfo(
        title='Error',
        message=str(value)
    )

def show_plot(mode, target_update_entry, choice = 0):
    print("Show Plot function executed")
    fig = plt.figure(1)
    try:
        df = pd.read_excel(filepath)
        df.head()
        t = np.arange(0.4, len(df)*0.4, 0.4)
    except (ValueError, FileNotFoundError) as e:
        if e == ValueError:
            popupError(e)
        elif e == FileNotFoundError:
            popupError(e)
        else:
            popupError("Import File Terlebih dahulu. File -> Import")
    if mode == "PV":
        try:
            PV = df[[mode]]
            PV.head()
            plt.plot(t, PV, picker = True, pickradius = 5)
            fig.canvas.mpl_connect('pick_event', lambda event: onclick(event, target_update_entry, mode))
        except (UnboundLocalError, KeyError) as e:
            if e == UnboundLocalError:
                popupError("Error Terjadi")
            elif e == KeyError:
                popupError("Error Terjadi")
            else:
                popupError("Pastikan Format Excel Sudah Sesuai")
    elif mode == "OP":
        OP = df[[mode]]
        OP.head()
        plt.plot(t, OP, picker = True, pickradius = 5)
        fig.canvas.mpl_connect('pick_event', lambda event: onclick(event, target_update_entry, mode))
    elif mode == "Y":
        OP = df[["OP"]]
        OP.head()
        plt.plot(t, OP, picker = True, pickradius = 5)
        fig.canvas.mpl_connect('pick_event', lambda event: onclick(event, target_update_entry, mode))
    elif mode == "T" and choice == 1:
        PV = df[["PV"]]
        PV.head()
        plt.plot(t, PV, picker = True, pickradius = 5)
        fig.canvas.mpl_connect('pick_event', lambda event: onclick(event, target_update_entry, mode))
    elif mode == "T" and choice == 2:
        OP = df[["OP"]]
        OP.head()
        plt.plot(t, OP, picker = True, pickradius = 5)
        fig.canvas.mpl_connect('pick_event', lambda event: onclick(event, target_update_entry, mode))
    else:
        print("Ada yang salah")
    plt.xlabel('Time(s)')
    plt.ylabel('Value')
    plt.show()

def sprtdByCommaToArray(string_input):
    res = string_input.split(", ")
    return res

def popupResultPID():
    print("Pop up result function executed")
    global resultPID
    try:
        APD = float(apd.get())
        MLD = float(mpd.get())
        inPVP1 = float(sprtdByCommaToArray(InPVP1Val.get())[0])
        inPVT1 = float(sprtdByCommaToArray(InPVP1Val.get())[1])
        inPVP2 = float(sprtdByCommaToArray(InPVP2Val.get())[0])
        inPVT2 = float(sprtdByCommaToArray(InPVP2Val.get())[1])
        finPVP1 = float(sprtdByCommaToArray(FinPVP1Val.get())[0])
        finPVT1 = float(sprtdByCommaToArray(FinPVP1Val.get())[1])
        finPVP2 = float(sprtdByCommaToArray(FinPVP2Val.get())[0])
        finPVT2 = float(sprtdByCommaToArray(FinPVP2Val.get())[1])
        inOP = float(sprtdByCommaToArray(InOPVal.get())[0])
        finOP = float(sprtdByCommaToArray(FinOPVal.get())[0])
        PVChangeTime =float(sprtdByCommaToArray(PVCTVal.get())[0])
        OPChangeTime = float(sprtdByCommaToArray(OPCTVal.get())[0])
        inSlope, finSlope, Td, Kp, Lambda, integralTime, controllerGain = calculatePIController(APD, MLD, inPVP1, inPVP2, inPVT1, inPVT2, finPVP1, finPVP2, finPVT1, finPVT2, inOP, finOP, PVChangeTime, OPChangeTime)
        #Show Table
        column_names = ["No", "Parameter", "Value"]
        data_show = [
            ["APD", APD], 
            ["MLD", MLD], 
            ["inPVP1", inPVP1], 
            ["inPVT1", inPVT1], 
            ["inPVP2", inPVP2], 
            ["inPVT2", inPVT2], 
            ["finPVP1", finPVP1], 
            ["finPVT1", finPVT1], 
            ["finPVP2", finPVP2], 
            ["finPVT2", finPVT2], 
            ["inOP", inOP], 
            ["finOP", finOP], 
            ["PVChangeTime", PVChangeTime], 
            ["OPChangeTime", OPChangeTime], 
            ["Initial Slope", inSlope], 
            ["Final Slope", finSlope],
            ["Td", Td],
            ["Kp", Kp],
            ["Lambda", Lambda],
            ["Proportional Gain", controllerGain],
            ["Integral Time", integralTime/60],
        ]
        print(tabulate(data_show, headers=column_names, tablefmt="fancy_grid", showindex="always"))
        #Pop up window result
        resultPID = tk.Toplevel(window)
        resultPID.title("Result")
        resultPID.geometry("300x250")
        resultPID.config(bg=bg)
        resultPID.resizable(False, False)
        resultPID.iconphoto(False, tk.PhotoImage(file = "logo.png"))
        label = tk.Label(resultPID, text="PI Controller Gain Result", font=('Arial', 12, 'bold'),fg="#fff", bg=bg)
        label.pack(pady=20)
        #P Gain
        tk.Label(resultPID,  text="P Gain :", bg=bg, font=('Arial',12,'bold'), fg="#fff").pack(pady=5)
        tk.Label(resultPID,  text=controllerGain, bg=bg, font=('Arial',12,'bold'), fg="#fff").pack(pady=5)
        #Integral Time
        tk.Label(resultPID,  text="I Time (1/min):", bg=bg, font=('Arial',12,'bold'), fg="#fff").pack(pady=5)
        tk.Label(resultPID,  text=integralTime/60, bg=bg, font=('Arial',12,'bold'), fg="#fff").pack(pady=5)
        #Save Data
        data = [
            [
                [APD, MLD, inPVP1,inPVT1,inPVP2,inPVT2,finPVP1,finPVT1,finPVP2,finPVT2,inOP,finOP,PVChangeTime,OPChangeTime], 
                [inSlope, finSlope, Td, Kp, Lambda, integralTime, controllerGain]
            ]
        ]
        tk.Button(resultPID, text ="Save Result", font=('Arial',10,'bold'), command=lambda: save_data(str(data))).pack(pady=5)
    except (IndexError , ZeroDivisionError) as e:
        if e == IndexError:
            popupError("Masukkan Nilai yang Benar")
        elif e == ZeroDivisionError:
            popupError("Pembagian dengan 0")
        else:
            popupError(e)

#Main Window
print("Application Started")
window.title("PID Tuner by Electrical Engineering UGM 2023")
window.geometry("335x490+400+30")
window.iconphoto(False, tk.PhotoImage(file = "logo.png"))
window.resizable(False, False)
inner_frame  =  tk.Frame(window,  width=200,  height=  200)

#Create Label Title
tk.Label(window,  text="Parameter", bg=bg, font=('Arial',13,'bold'), fg="#fff").grid(row=0,  column=0,  padx=5,  pady=5)
tk.Label(window,  text="Value", bg=bg, font=('Arial',13,'bold'), fg="#fff").grid(row=0,  column=2,  padx=5,  pady=5)

#APD
tk.Label(window,  text="Allowed Process Deviation", bg=bg, font=('Arial',10,'bold'), fg="#fff").grid(sticky="W", row=1,  columnspan=2,  padx=5,  pady=5)
apd = tk.Entry(window, font=('Arial',10,'bold'), width=16, justify='center')
apd.grid(sticky="W", row=1,  column=2,  padx=5,  pady=5)
apd.insert(0, "0")

#MLD
tk.Label(window,  text="Maximum Load Disturbance", bg=bg, font=('Arial',10,'bold'), fg="#fff").grid(sticky="W", row=2,  columnspan=2,  padx=5,  pady=5)
mpd = tk.Entry(window, font=('Arial',10,'bold'), width=16, justify='center')
mpd.grid(sticky="W", row=2,  column=2,  padx=5,  pady=5)
mpd.insert(0, "0")

#InPV1
tk.Label(window,  text="Initial PV point 1 (%)", bg=bg, font=('Arial',10,'bold'), fg="#fff").grid(sticky="W", row=3,  column=0,  padx=5,  pady=5)
InPVP1Val = tk.Entry(window, font=('Arial',10,'bold'), width=16, justify='center')
InPVP1Val.insert(0, "0")
InPVP1Val.grid(sticky="W", row=3,  column=2,  padx=5,  pady=5)
InPVP1Sel = tk.Button(window, text ="Select", font=('Arial',10,'bold'), command=lambda: show_plot("PV", InPVP1Val))
InPVP1Sel.grid(sticky="W", row=3,  column=1,  padx=5,  pady=5)

#InPV2
tk.Label(window,  text="Initial PV point 2 (%)", bg=bg, font=('Arial',10,'bold'), fg="#fff").grid(sticky="W", row=4,  column=0,  padx=5,  pady=5)
InPVP2Val = tk.Entry(window, font=('Arial',10,'bold'), width=16, justify='center')
InPVP2Val.insert(0, "0")
InPVP2Val.grid(sticky="W", row=4,  column=2,  padx=5,  pady=5)
InPVP2 = tk.Button(window, text ="Select", font=('Arial',10,'bold'), command=lambda: show_plot("PV", InPVP2Val))
InPVP2.grid(sticky="W", row=4,  column=1,  padx=5,  pady=5)

#FinPVP1
tk.Label(window,  text="Final PV point 1 (%)", bg=bg, font=('Arial',10,'bold'), fg="#fff").grid(sticky="W", row=5,  column=0,  padx=5,  pady=5)
FinPVP1Val = tk.Entry(window, font=('Arial',10,'bold'), width=16, justify='center')
FinPVP1Val.insert(0, "0")
FinPVP1Val.grid(sticky="W", row=5,  column=2,  padx=5,  pady=5)
FinPVP1 = tk.Button(window, text ="Select", font=('Arial',10,'bold'), command=lambda: show_plot("PV", FinPVP1Val))
FinPVP1.grid(sticky="W", row=5,  column=1,  padx=5,  pady=5)

#FinPVP2
tk.Label(window,  text="Final PV point 2 (%)", bg=bg, font=('Arial',10,'bold'), fg="#fff").grid(sticky="W", row=6,  column=0,  padx=5,  pady=5)
FinPVP2Val = tk.Entry(window, font=('Arial',10,'bold'), width=16, justify='center')
FinPVP2Val.insert(0, "0")
FinPVP2Val.grid(sticky="W", row=6,  column=2,  padx=5,  pady=5)
FinPVP2 = tk.Button(window, text ="Select", font=('Arial',10,'bold'), command=lambda: show_plot("PV", FinPVP2Val))
FinPVP2.grid(sticky="W", row=6,  column=1,  padx=5,  pady=5)

#InOP
tk.Label(window,  text="Initial OP", bg=bg, font=('Arial',10,'bold'), fg="#fff").grid(sticky="W", row=7,  column=0,  padx=5,  pady=5)
InOPVal = tk.Entry(window, font=('Arial',10,'bold'), width=16, justify='center')
InOPVal.insert(0, "0")
InOPVal.grid(sticky="W", row=7,  column=2,  padx=5,  pady=5)
InOP = tk.Button(window, text ="Select", font=('Arial',10,'bold'), command=lambda: show_plot("Y", InOPVal))
InOP.grid(sticky="W", row=7,  column=1,  padx=5,  pady=5)

#finOP
tk.Label(window,  text="Final OP", bg=bg, font=('Arial',10,'bold'), fg="#fff").grid(sticky="W", row=8,  column=0,  padx=5,  pady=5)
FinOPVal = tk.Entry(window, font=('Arial',10,'bold'), width=16, justify='center')
FinOPVal.insert(0, "0")
FinOPVal.grid(sticky="W", row=8,  column=2,  padx=5,  pady=5)
FinOP = tk.Button(window, text ="Select", font=('Arial',10,'bold'), command=lambda: show_plot("Y", FinOPVal))
FinOP.grid(sticky="W", row=8,  column=1,  padx=5,  pady=5)

#PVCT
tk.Label(window,  text="PV change time (s)", bg=bg, font=('Arial',10,'bold'), fg="#fff").grid(sticky="W", row=9,  column=0,  padx=5,  pady=5)
PVCTVal = tk.Entry(window, font=('Arial',10,'bold'), width=16, justify='center')
PVCTVal.insert(0, "0")
PVCTVal.grid(sticky="W", row=9,  column=2,  padx=5,  pady=5)
PVCT = tk.Button(window, text ="Select", font=('Arial',10,'bold'), command=lambda: show_plot("T", PVCTVal, 1))
PVCT.grid(sticky="W", row=9,  column=1,  padx=5,  pady=5)

#OPCT
tk.Label(window,  text="OP change time (s)", bg=bg, font=('Arial',10,'bold'), fg="#fff").grid(sticky="W", row=10,  column=0,  padx=5,  pady=5)
OPCTVal = tk.Entry(window, font=('Arial',10,'bold'), width=16, justify='center')
OPCTVal.insert(0, "0")
OPCTVal.grid(sticky="W", row=10,  column=2,  padx=5,  pady=5)
OPCT = tk.Button(window, text ="Select", font=('Arial',10,'bold'), command=lambda: show_plot("T", OPCTVal, 2))
OPCT.grid(sticky="W", row=10,  column=1,  padx=5,  pady=5)

#Calculate
tk.Button(window, text ="Calculate", font=('Arial',10,'bold'), command=popupResultPID, width=30).grid(row = 15, columnspan=4, sticky="S", pady=15)
createMenu()

window.mainloop()

print("Application Closed")
