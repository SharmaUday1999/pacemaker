import tkinter as tk                # python 3
from tkinter import font  as tkfont # python 3
import openpyxl
import os
import time


# opening the existing excel file
filename = 'userdata.xlsx'

if os.path.isfile(filename):
    wb = openpyxl.load_workbook(filename)
else:
    wb = openpyxl.Workbook()
    wb.save(filename)

# create the sheet object
ws = wb.active


# TODO: We should iterate over an array and fill each column instead in the future (when/if we have time)
if ((ws['A1'] == 'First name') and (ws['B1'] == 'Last Name') and (ws['C1'] == 'Email') and (ws['D1'] == 'Username')  and (ws['E1'] == 'Password')
    and (ws['F1'] == 'Lower Rate Limit') and (ws['G1'] == 'Upper Rate Limit') and (ws['H1'] == 'Atrial Amplitude') and (ws['I1'] == 'Atrial Pulse Width')
    and (ws['J1'] == 'Ventricular Amplitude') and (ws['K1'] == 'Ventricular Pulse Width') and (ws['L1'] == 'VRP') and (ws['M1'] == 'ARP')):
    pass
else:
    ws['A1'] = 'First name'
    ws['B1'] = 'Last name'
    ws['C1'] = 'Email'
    ws['D1'] = 'Username'
    ws['E1'] = 'Password'
    ws['F1'] = 'Lower Rate Limit'
    ws['G1'] = 'Upper Rate Limit'
    ws['H1'] = 'Atrial Amplitude'
    ws['I1'] = 'Atrial Pulse Width'
    ws['J1'] = 'Ventricular Amplitude'
    ws['K1'] = 'Ventricular Pulse Width'
    ws['L1'] = 'VRP'
    ws['M1'] = 'ARP'
wb.save(filename)



class pacemaker(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        self.title_font = tkfont.Font(family='Helvetica', size=18, weight="bold")

    
        # the container is where we'll stack a bunch of frames
        # on top of each other, then the one we want visible
        # will be raised above the others
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)


        self.frames = {}
        for F in (welcomePage, registerPage, loginPage, mainPage):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame

            # put all of the pages in the same location;
            # the one on the top of the stacking order
            # will be the one that is visible.
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("welcomePage")

    def show_frame(self, page_name):
        '''Show a frame for the given page name'''
        frame = self.frames[page_name]
        frame.tkraise()

        
class welcomePage(tk.Frame):

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.winfo_toplevel().title("Pacemaker control application")

        label = tk.Label(self, text="WELCOME TO THE PACEMAKER CONTROL APPLICATION", font=controller.title_font)
        label.pack(side="top", fill="x", pady=10, padx = 20)

        button1 = tk.Button(self, text="Register",
                            command=lambda: controller.show_frame("registerPage"))
        button2 = tk.Button(self, text="Login",
                            command=lambda: controller.show_frame("loginPage"))
        button1.pack(pady=2)
        button2.pack(pady=10)


class registerPage(tk.Frame):

    def saveData(self):
        if (firstnameLabelEntry.get() == "" or
            lastnameLabelEntry.get() == "" or
            emailLabelEntry.get() == "" or
            usernameLabelEntry.get() == "" or
            passwordLabelEntry.get() == "" ):

            regLabel = tk.Label(self, text = "")
            emptyLabel = tk.Label(self, text = "All inputs must be filled")
            emptyLabel.grid(row = 2, column = 4)
        else:

            # assigning the max row and max column
            # value upto which data is written
            # in an excel sheet to the variable
            current_row = ws.max_row
            current_column = ws.max_column

            duplicateUsernameToggle = 0

            for i in range(1,current_row+1):
                if ws.cell(row = i, column = 4).value == usernameLabelEntry.get():
                    duplicateUsernameToggle = 1
                    break

            # get method returns current text
            # as string which we write into
            # excel spreadsheet at particular location
            if duplicateUsernameToggle == 0:

                ws.cell(row=current_row + 1, column=1).value = firstnameLabelEntry.get()
                ws.cell(row=current_row + 1, column=2).value = lastnameLabelEntry.get()
                ws.cell(row=current_row + 1, column=3).value = emailLabelEntry.get()
                ws.cell(row=current_row + 1, column=4).value = usernameLabelEntry.get()
                ws.cell(row=current_row + 1, column=5).value = passwordLabelEntry.get()

                # save the file
                wb.save(filename)

                # clear the content of text entry box
                firstnameLabelEntry.delete(0, 'end')
                lastnameLabelEntry.delete(0, 'end')
                emailLabelEntry.delete(0, 'end')
                usernameLabelEntry.delete(0, 'end')
                passwordLabelEntry.delete(0, 'end')

                regLabel = tk.Label(self, text = "")
                regLabel = tk.Label(self, text = "You have been registered!")
                regLabel.grid(row = 2, column = 4)
            else:
                regLabel = tk.Label(self, text = "")
                regLabel = tk.Label(self, text = "Username not available!")
                regLabel.grid(row = 2, column = 4)

                # clear the content of text entry box
                firstnameLabelEntry.delete(0, 'end')
                lastnameLabelEntry.delete(0, 'end')
                emailLabelEntry.delete(0, 'end')
                usernameLabelEntry.delete(0, 'end')
                passwordLabelEntry.delete(0, 'end')



    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller

        firstnameLabel = tk.Label(self ,text = "First Name").grid(row = 0,column = 0, padx = 5, pady = 5)
        lastnameLabel = tk.Label(self ,text = "Last Name").grid(row = 1,column = 0, padx = 5, pady = 5)
        emailLabel = tk.Label(self ,text = "Email").grid(row = 2,column = 0, padx = 5, pady = 5)
        usernameLabel = tk.Label(self ,text = "Username").grid(row = 3,column = 0, padx = 5, pady = 5)
        passwordLabel = tk.Label(self ,text = "Password").grid(row = 4,column = 0, padx = 5, pady = 5)

        global firstnameLabelEntry
        global lastnameLabelEntry
        global emailLabelEntry
        global usernameLabelEntry
        global passwordLabelEntry

        firstnameLabelEntry = tk.Entry(self)
        firstnameLabelEntry.grid(row = 0,column = 1, padx = 5, pady = 5)
        lastnameLabelEntry = tk.Entry(self)
        lastnameLabelEntry.grid(row = 1,column = 1, padx = 5, pady = 5)
        emailLabelEntry = tk.Entry(self)
        emailLabelEntry.grid(row = 2,column = 1, padx = 5, pady = 5)
        usernameLabelEntry = tk.Entry(self)
        usernameLabelEntry.grid(row = 3,column = 1, padx = 5, pady = 5)
        passwordLabelEntry = tk.Entry(self)
        passwordLabelEntry.grid(row = 4,column = 1, padx = 5, pady = 5)


        buttonRegister = tk.Button(self, text="Register", command = self.saveData)
        buttonRegister.grid(row = 6, column = 0, padx = 5, pady = 5)

        buttonReturn = tk.Button(self, text="Return to main menu",
                           command=lambda: controller.show_frame("welcomePage"))
        buttonReturn.grid(row = 6, column = 1, padx = 5, pady = 5)


class loginPage(tk.Frame):

    def login(self):
        rowMatchPassword = 1
        for i in range(1,ws.max_row+1):
            if ws.cell(row = i, column = 4).value == usernameLabelEntryLogin.get():
                rowMatchPassword = i
                break

        if ws.cell(row = rowMatchPassword, column = 5).value == passwordLabelEntryLogin.get():
            loggedInRow = rowMatchPassword
            self.controller.show_frame('mainPage')
            usernameLabelEntryLogin.delete(0,'end')
            passwordLabelEntryLogin.delete(0,'end')
            mainPage.setLoggedInRow(self.controller.frames['mainPage'], loggedInRow)

        #implement incorrect login text


    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        label = tk.Label(self, text="Login", font=controller.title_font).grid(row = 0, column = 0)


        usernameLabel = tk.Label(self ,text = "Username").grid(row = 1,column = 0, padx = 5, pady = 5)
        passwordLabel = tk.Label(self ,text = "Password").grid(row = 2,column = 0, padx = 5, pady = 5)

        global usernameLabelEntryLogin
        global passwordLabelEntryLogin

        usernameLabelEntryLogin = tk.Entry(self)
        usernameLabelEntryLogin.grid(row = 1,column = 1, padx = 5, pady = 5)
        passwordLabelEntryLogin = tk.Entry(self)
        passwordLabelEntryLogin.grid(row = 2,column = 1, padx = 5, pady = 5)

        buttonLogin = tk.Button(self, text="Login", command = self.login)
        buttonLogin.grid(row = 6, column = 0, padx = 5, pady = 5)

        buttonReturn = tk.Button(self, text="Return to welcome page",
                           command=lambda: controller.show_frame("welcomePage"))
        buttonReturn.grid(row = 6, column = 1, padx = 5, pady = 5)

class mainPage(tk.Frame):
    _state = 'none'
    # TODO: Should move this into a separate module, it's just an object storing constants
    _params = {
        'LOWER_RATE_LIMIT' : {
            'AOO': 'normal',
            'VOO': 'normal',
            'AAI' : 'normal',
            'VVI' : 'normal'
            },
        'UPPER_RATE_LIMIT' : {
            'AOO': 'normal',
            'VOO': 'normal',
            'AAI' : 'normal',
            'VVI' : 'normal'
            },
        'ATRIAL_AMPLITUDE' : {
            'AOO': 'normal',
            'VOO': 'disabled',
            'AAI' : 'normal',
            'VVI' : 'disabled'
            },
        'ATRIAL_PULSE_WIDTH': {
            'AOO': 'normal',
            'VOO': 'disabled',
            'AAI' : 'normal',
            'VVI' : 'disabled'
            },
        'VENTRICULAR_AMPLITUDE': {
            'AOO': 'disabled',
            'VOO': 'normal',
            'AAI' : 'disabled',
            'VVI' : 'normal'
            },
        'VENTRICULAR_PULSE_WIDTH': {
            'AOO': 'disabled',
            'VOO': 'normal',
            'AAI' : 'disabled',
            'VVI' : 'normal'
            },
        'VRP' : {
            'AOO': 'disabled',
            'VOO': 'disabled',
            'AAI' : 'disabled',
            'VVI' : 'normal'
            },
        'ARP' : {
            'AOO': 'disabled',
            'VOO': 'disabled',
            'AAI' : 'normal',
            'VVI' : 'disabled'
            }
        }


    def setMode(self, mode):
        #Change state and fields depending on which pacing mode is selected. Default none

        self._state = mode
        self._loggedInRow = 1

        lrlEntry.configure(state= self._params['LOWER_RATE_LIMIT'][mode])
        urlEntry.configure(state= self._params['UPPER_RATE_LIMIT'][mode])
        atrialAmpEntry.configure(state= self._params['ATRIAL_AMPLITUDE'][mode])
        atrialPWEntry.configure(state= self._params['ATRIAL_PULSE_WIDTH'][mode])
        venAmpEntry.configure(state= self._params['VENTRICULAR_AMPLITUDE'][mode])
        venPWEntry.configure(state= self._params['VENTRICULAR_PULSE_WIDTH'][mode])
        vrpEntry.configure(state= self._params['VRP'][mode])
        arpEntry.configure(state= self._params['ARP'][mode])

    def Save(self):

        ws.cell(row = self._loggedInRow, column = 6).value = lrlEntry.get()
        ws.cell(row = self._loggedInRow, column = 7).value = urlEntry.get()
        ws.cell(row = self._loggedInRow, column = 8).value = atrialAmpEntry.get()
        ws.cell(row = self._loggedInRow, column = 9).value = atrialPWEntry.get()
        ws.cell(row = self._loggedInRow, column = 10).value = venAmpEntry.get()
        ws.cell(row = self._loggedInRow, column = 11).value = venPWEntry.get()
        ws.cell(row = self._loggedInRow, column = 12).value = vrpEntry.get()
        ws.cell(row = self._loggedInRow, column = 13).value = arpEntry.get()
        wb.save(filename)

    def populateUserData(self):

        #Clear any data that might exist
        lrlEntry.delete(0, 'end')
        urlEntry.delete(0, 'end')
        atrialAmpEntry.delete(0, 'end')
        atrialPWEntry.delete(0, 'end')
        venAmpEntry.delete(0, 'end')
        venPWEntry.delete(0, 'end')
        vrpEntry.delete(0, 'end')
        arpEntry.delete(0, 'end')
        
        lrlEntry.insert(0, ws.cell(row = self._loggedInRow, column = 6).value if type(ws.cell(row = self._loggedInRow, column = 6).value) == str else 0)
        urlEntry.insert('end', ws.cell(row = self._loggedInRow, column = 7).value if type(ws.cell(row = self._loggedInRow, column = 7).value) == str else 0)
        atrialAmpEntry.insert('end', ws.cell(row = self._loggedInRow, column = 8).value if type(ws.cell(row = self._loggedInRow, column = 8).value) == str else 0)
        atrialPWEntry.insert('end', ws.cell(row = self._loggedInRow, column = 9).value if type(ws.cell(row = self._loggedInRow, column = 9).value) == str else 0)
        venAmpEntry.insert('end', ws.cell(row = self._loggedInRow, column = 10).value if type(ws.cell(row = self._loggedInRow, column = 10).value) == str else 0)
        venPWEntry.insert('end', ws.cell(row = self._loggedInRow, column = 11).value if type(ws.cell(row = self._loggedInRow, column = 11).value) == str else 0)
        vrpEntry.insert('end', ws.cell(row = self._loggedInRow, column = 12).value if type(ws.cell(row = self._loggedInRow, column = 12).value) == str else 0)
        arpEntry.insert('end', ws.cell(row = self._loggedInRow, column = 13).value if type(ws.cell(row = self._loggedInRow, column = 13).value) == str else 0)

    def setLoggedInRow(self, row):
        self._loggedInRow = row
        self.populateUserData()
        

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        label = tk.Label(self, text="WELCOME TO THE PACEMAKER PORTAL", font=controller.title_font).grid(row = 0, column = 0, columnspan = 5)

        #Page Navigation
        logoutButton = tk.Button(self, text="Logout",
                           command=lambda: controller.show_frame('welcomePage'))
        logoutButton.grid(row = 20, column = 3, padx = 5, pady = 5)

        #Pacing Modes
        aooButton = tk.Button(self, text="AOO",
                           command=lambda: self.setMode('AOO'))
        aooButton.grid(row = 10, column = 0, padx = 5, pady = 5)

        vooButton = tk.Button(self, text="VOO",
                           command=lambda: self.setMode('VOO'))
        vooButton.grid(row = 10, column = 2, padx = 5, pady = 5)

        aaiButton = tk.Button(self, text="AAI",
                           command=lambda: self.setMode('AAI'))
        aaiButton.grid(row = 10, column = 4, padx = 5, pady = 5)

        vviButton = tk.Button(self, text="VVI",
                           command=lambda: self.setMode('VVI'))
        vviButton.grid(row = 10, column = 6, padx = 5, pady = 5)


        #Pacing Mode Parameters
        lrlLabel = tk.Label(self ,text = "Lower Rate Limit",).grid(row = 15,column = 0, padx = 1, pady = 1, columnspan=2)
        urlLabel = tk.Label(self ,text = "Upper Rate Limit").grid(row = 16,column = 0, padx = 1, pady = 1, columnspan=2)
        atrialAmpLabel = tk.Label(self ,text = "Atrial Amplitude").grid(row = 17,column = 0, padx = 1, pady = 1, columnspan=2)
        atrialPWLabel = tk.Label(self ,text = "Atrial Pulse Width").grid(row = 18,column = 0, padx = 1, pady = 1, columnspan=2)
        venAmpLabel = tk.Label(self ,text = "Ventricular Amplitude").grid(row = 15,column = 3, padx = 1, pady = 1, columnspan=2)
        venPWLabel = tk.Label(self ,text = "Ventricular Pulse Width",).grid(row = 16,column = 3, padx = 1, pady = 1, columnspan=2)
        vrpLabel = tk.Label(self ,text = "VRP",).grid(row = 17,column = 3, padx = 1, pady = 3, columnspan=2)
        arpLabel = tk.Label(self ,text = "ARP",).grid(row = 18,column = 3, padx = 1, pady = 3, columnspan=2)


        global lrlEntry
        global urlEntry
        global atrialAmpEntry
        global atrialPWEntry
        global venAmpEntry
        global venPWEntry
        global vrpEntry
        global arpEntry


        lrlEntry = tk.Entry(self, width=5, disabledbackground='grey')
        lrlEntry.grid(row = 15,column = 2)
        urlEntry = tk.Entry(self, width=5, disabledbackground='grey')
        urlEntry.grid(row = 16,column = 2, padx = 1, pady = 1)
        atrialAmpEntry = tk.Entry(self, width=5, disabledbackground='grey')
        atrialAmpEntry.grid(row = 17,column = 2, padx = 1, pady = 1)
        atrialPWEntry = tk.Entry(self, width=5, disabledbackground='grey')
        atrialPWEntry.grid(row = 18,column = 2, padx = 1, pady = 1)
        venAmpEntry = tk.Entry(self, width=5, disabledbackground='grey')
        venAmpEntry.grid(row = 15,column = 6, padx = 1, pady = 1)
        venPWEntry = tk.Entry(self, width=5, disabledbackground='grey')
        venPWEntry.grid(row = 16,column = 6, padx = 1, pady = 1)
        vrpEntry = tk.Entry(self, width=5, disabledbackground='grey')
        vrpEntry.grid(row = 17,column = 6, padx = 1, pady = 1)
        arpEntry = tk.Entry(self, width=5, disabledbackground='grey')
        arpEntry.grid(row = 18,column = 6, padx = 1, pady = 1)


        buttonSave = tk.Button(self, text="Save", command = self.Save)
        buttonSave.grid(row = 19, column = 6, padx = 5, pady = 5)

    

if __name__ == "__main__":
    app = pacemaker()
    app.mainloop()
