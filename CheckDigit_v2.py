###############################################################################
## Author: Jonathan Wolfe
## Email: jawhf4@mail.umkc.edu
## Class: CS 101 0002
## Assignment: Program 3 - Python - Check Digit
## Due: 9/30/18
###############################################################################

import random
import csv
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

class AssignmentTrackerApp(tk.Frame):
    '''Assignment Tracker window class'''

    
    def __init__(self, master = None):
        '''Initialization of the Tracker app'''

        # Setup window
        tk.Frame.__init__(self, master)
        self.master.title('Assignment Tracker')
        self.master.resizable(0,0)
        self.master.protocol("WM_DELETE_WINDOW", self.deleteWindow)

        # Initialize file name and save progress
        # for current state
        self.fileName = ''
        self.progressSaved = True

        # Create and position widgets and window
        self.createWindowWidgets(self.master)
        self.createTrackerTV(self.master)
        self.positionWidgets(self.master)        
        self.createFileMenu(self.master)
        self.centerWindow(self.master)


    def createWindowWidgets(self, win):
        ''' 
           Args: win (tkinter.Tk): Root window'''
        self.AssignmentLabel = tk.Label(win, text="Assignment Number")
        self.AssignmentBox = tk.Entry()
        self.StudentLabel = tk.Label(win, text='Student Name')
        self.StudentBox = tk.Entry()
        self.GradeLabel = tk.Label(win, text='Assignment Grade')
        self.GradeBox = tk.Entry()
        self.AddEntryButton = tk.Button(win, text="Add/Edit", command=self.addEntry)
        self.clearBoxesButton = tk.Button(win, text = 'Clear', command=self.clearBoxes)
        self.removeEntryButton = tk.Button(win, text = 'Remove', command = self.removeEntry)
        self.getEntryButton = tk.Button(win, text = 'Retrieve', command = self.getAssignment)
        self.randomButton = tk.Button(win, text = 'Random', command = self.getRandom)


    def createTrackerTV(self, win):
        ''' 
           Args: win (tkinter.Tk): Root window'''
        self.trackerTV = ttk.Treeview(win, columns=('Assignment', 'Class', 'Type', 'Student', 'Grade'), selectmode='browse', show='headings')
        self.trackerTV.heading('#1', text='Assignment', command=lambda: self.sortColumn('Assignment', '#1', False))
        self.trackerTV.heading('#2', text='Class', command=lambda: self.sortColumn('Class', '#2', False))
        self.trackerTV.heading('#3', text='Type', command=lambda: self.sortColumn('Type', '#3', False))
        self.trackerTV.heading('#4', text='Student', command=lambda: self.sortColumn('Student', '#4', False))
        self.trackerTV.heading('#5', text='Grade', command=lambda: self.sortColumn('Grade', '#5', False))
        self.trackerTV.column('#1', minwidth = 125, width = 125, stretch = False, anchor=tk.W)
        self.trackerTV.column('#2', minwidth = 125, width = 125, stretch = False, anchor=tk.W)
        self.trackerTV.column('#3', minwidth = 125, width = 125, stretch = False, anchor=tk.W)
        self.trackerTV.column('#4', minwidth = 125, width = 125, stretch = False, anchor=tk.W)
        self.trackerTV.column('#5', minwidth = 125, width = 125, stretch = False, anchor=tk.W)
        self.scrollbar = tk.Scrollbar(win, orient="vertical", command=self.trackerTV.yview)
        self.trackerTV.configure(yscrollcommand=self.scrollbar.set)
        self.trackerTV.bind('<Button>', self.separatorClick)
        self.trackerTV.bind('<ButtonRelease>', self.separatorClick)
        self.trackerTV.bind('<Motion>', self.separatorClick)


    def positionWidgets(self, win):
        ''' 
           Args: win (tkinter.Tk): Root window'''
        self.AssignmentLabel.grid(row=0, column=0, sticky=tk.E)
        self.AssignmentBox.grid(row=0, column=1, sticky=tk.W+tk.E)
        self.StudentLabel.grid(row = 1, column = 0, sticky=tk.E)
        self.StudentBox.grid(row = 1, column = 1, sticky=tk.W+tk.E)
        self.GradeLabel.grid(row = 2, column = 0, sticky=tk.E)
        self.GradeBox.grid(row = 2, column = 1, sticky=tk.W+tk.E)
        self.clearBoxesButton.grid(row = 0, column = 2, sticky=tk.W+tk.E)
        self.randomButton.grid(row = 0, column = 3, sticky=tk.W+tk.E)
        win.grid_rowconfigure(3, minsize = 25)
        self.AddEntryButton.grid(row=4, column=0, sticky=tk.W+tk.E)
        self.getEntryButton.grid(row = 4, column = 1, sticky=tk.W+tk.E)
        self.removeEntryButton.grid(row = 4, column = 2, sticky=tk.W+tk.E)
        self.trackerTV.grid(row=5, columnspan=4)
        self.scrollbar.grid(row=5, column=4, sticky=tk.N+tk.S)


    def createFileMenu(self, win):
        ''' 
           Args: win (tkinter.Tk): Root window'''
        self.menuBar = tk.Menu(win)
        self.fileMenu = tk.Menu(self.menuBar, tearoff = 0)
        self.fileMenu.add_command(label='New', command = self.newProject)
        self.fileMenu.add_command(label='Open', command = self.openProject)
        self.fileMenu.add_command(label='Save', command = self.saveProject)
        self.fileMenu.add_command(label='Save As', command = self.saveAsProject)
        self.fileMenu.add_command(label='Exit', command = self.deleteWindow)
        self.menuBar.add_cascade(label = 'File', menu = self.fileMenu)
        win.config(menu = self.menuBar)


    def centerWindow(self, win):
        '''Centers the  root window in the primary screen
           Args: win (tkinter.Tk): Root window'''
        try:
            # Find window heigth and width
            win.update_idletasks()
            winWidth = win.winfo_width()
            winHeight = win.winfo_height()

            # Find the amount to offset the window
            xOffset = (win.winfo_screenwidth() // 2) - (winWidth // 2)
            yOffset = (win.winfo_screenheight() // 2) - (winHeight // 2)

            # Set the window 
            win.geometry('{}x{}+{}+{}' \
                         .format(winWidth, winHeight, xOffset, yOffset))            
        except Exception as exception:
            messagebox.showerror('Unexpected Error', \
                                 'Unable to center window \n' \
                                 'Exception: {}'.format(exception))


    def getAssignment(self):
        '''Populates the entry boxes with the values of
            the selected Assignment'''
        selectedItems = self.trackerTV.selection()
        if len(selectedItems) == 1:
            self.clearBoxes()
            values = self.trackerTV.item(selectedItems[0],"values")
            self.AssignmentBox.insert(0, values[0])
            self.StudentBox.insert(0, values[3])
            self.GradeBox.insert(0,values[4])
        else:
            messagebox.showerror('No Selection', \
                                 'No items have been selected to retrieve')


    def addEntry(self):
        '''Adds the Assignment, Class, Type, Student and Grade
            to the Tracker'''
        validateStr = self.AssignmentBox.get().upper()

        if self.validateAssignment(validateStr):
            className = self.getAssignmentClass(validateStr)
            typeName = self.getAssignmentType(validateStr)
            student = self.StudentBox.get()
            grade = self.GradeBox.get()
            values = (validateStr, className, typeName, student, grade)

            try:
                if student.count(';') > 0:
                    raise ValueError('Student Name cannot contain semicolons')
                
                if grade.count(';') > 0:
                    raise ValueError('Grade cannot contain semicolons')
                
                if validateStr in self.trackerTV.get_children(''):
                    self.trackerTV.item(validateStr, values = values)
                else:
                    self.trackerTV.insert('', 'end', iid=validateStr, \
                                          values = values)
                self.clearBoxes()

                if self.progressSaved:
                    self.progressSaved = False
                
            except ValueError as exception:
                messagebox.showerror('Invalid Character', exception)

    def validateAssignment(self, validateStr):
        '''Return whether Assignment meets all requirements: 13 characters
           long, 0-9, A-F, first charater is A-D, second character is A-E,
           last character is a digit and the check digit is correct
           Args: validateStr (str): Assignment number
           Returns:  (bool) Assignment number is valid '''
        validEntry = False

        # Check length, alphanumeric, first is A-D, second is A-E and
        # check digit is a digit
        if len(validateStr) != 13:
            messagebox.showerror('Invalid Assignment Number', \
                                 'Assignment Number must be 13 characters ' \
                                 'in length')
        elif not validateStr.isalnum():
            messagebox.showerror('Invalid Assignment Number', \
                                 'Assignment Number must consist of 0 ' \
                                 'through 9 and A through F')
        elif validateStr[0] not in ['A', 'B', 'C', 'D']:
            messagebox.showerror('Invalid Assignment Number', \
                                 'The first digit must be a valid class ' \
                                 'identifier, A through D.')
        elif validateStr[1] not in ['A', 'B', 'C', 'D', 'E']:
            messagebox.showerror('Invalid Assignment Number', \
                                 'The second digit must be a valid ' \
                                 'assignment type identifier, A through E.')
        elif not validateStr[-1].isdigit():
            messagebox.showerror('Invalid Assignment Number', \
                                 'Last digit must be a number')
        else:
            validEntry = self.validateCheckDigit(validateStr)
            
        return validEntry

    def validateCheckDigit(self, validateStr):
        ''' Return whether the Assignment number agrees with the check digit
           Args: validateStr (str): Assignment Number
           Returns:  (bool) Assignment number agrees with check digit'''

        validEntry = False
        numberSum = 0

        # Sum the assignment while also checking for invalid characters
        for weight in range(0, 12):            
            if validateStr[weight].isdigit():
                numberSum += int(validateStr[weight]) * weight
                
            elif validateStr[weight] in ['A', 'B', 'C', 'D', 'E', 'F']:
                numberSum += (ord(validateStr[weight]) - 55) * weight
                
            else:
                messagebox.showerror('Invalid Assignment Number', \
                                     'String contains invalid characters "{}"' \
                                     .format(validateStr[weight]))
                break
            
        # Check the sum to the check digit
        else:            
            if (numberSum % 10) == int(validateStr[-1]):
                validEntry = True
                
            else:
                messagebox.showerror('Invalid Assignment Number', \
                                     'A character is entered incorrectly ' \
                                     'or transposed.')

        return validEntry

    def getAssignmentClass(self, AssignmentStr):
        ''' Return the class name from the assignment number
           Args: AssignmentStr (Str): Assignment Number
           Returns: (Str): Class name, return empty string
                           if no class is found'''
        className = ''

        if AssignmentStr[0] == 'A':
            className = 'CS 101'
        elif AssignmentStr[0] == 'B':
            className = 'CS 191'
        elif AssignmentStr[0] == 'C':
            className = 'CS 201'
        elif AssignmentStr[0] == 'D':
            className = 'CS 291'

        return className


    def getAssignmentType(self, AssignmentStr):
        ''' Return the assignment type from the assignment number
           Args: AssignmentStr (Str): Assignment Number
           Returns: (Str): Assignment Type, return empty string
                           if no type is found'''
        typeName = ''
        
        if AssignmentStr[1] == 'A':
            typeName = 'Test'
        elif AssignmentStr[1] == 'B':
            typeName = 'Program'
        elif AssignmentStr[1] == 'C':
            typeName = 'Quiz'
        elif AssignmentStr[1] == 'D':
            typeName = 'Final'
        elif AssignmentStr[1] == 'E':
            typeName = 'Other'
            
        return typeName


    def removeEntry(self):
        ''' Removes the selected item in the tracker'''
        
        try:
            selectedItems = self.trackerTV.selection()

            # Ask permission to delete before deleting
            canDelete = messagebox.askyesno('Delete Entry', \
                                            'Do you want to delete \n {}' \
                                            .format(selectedItems[0]))
                
            if canDelete:
                self.trackerTV.delete(selectedItems[0])
                
                if self.progressSaved:
                    self.progressSaved = False
                
        except IndexError:
            messagebox.showerror('No Selection', \
                                 'No items have been selected to retrieve')


    def clearBoxes(self):
        '''Removes text from Assignment, Student and Grade entry boxes'''
        
        self.AssignmentBox.delete(0, tk.END)
        self.StudentBox.delete(0, tk.END)
        self.GradeBox.delete(0, tk.END)


    def getRandom(self):
        '''Enters a random valid Assignment Number in the Assignment Box'''

        # Populate randomStr with valid random entries for each position
        randomStr = ''
        numberSum = 0
        randomStr += random.choice(['A', 'B', 'C', 'D'])
        randomStr += random.choice(['A', 'B', 'C', 'D', 'E'])

        for indx in range(0,10):
            randomStr += random.choice(['A', 'B', 'C', 'D', 'E', 'F', \
                                    '0', '1', '2', '3', '4', '5', \
                                    '6', '7', '8', '9',])

        # Get the sum of the generated number and add the check digit
        for weight in range(0, 12):
            if randomStr[weight].isdigit():
                numberSum += int(randomStr[weight]) * weight
            else:
                numberSum += (ord(randomStr[weight]) - 55) * weight
                
        randomStr += str(numberSum % 10)

        # Replace the Assignment box text with the generated number 
        self.AssignmentBox.delete(0, tk.END)
        self.AssignmentBox.insert(0, randomStr)


    def sortColumn(self, column, colID, reverse):
        '''Sorts the designated column contents in the Treeview
           Args:  column (str): Column name
                  colID (str): Column ID
                  reverse (bool): Is sorting in ascending order'''

        # Create a dictionary containing every item and the value
        # of each item in the column
        treeDic = {}

        for item in self.trackerTV.get_children(''):
            treeDic[item] = self.trackerTV.set(item, column)

        # Sort the dictionary and more the items to the correct places in the
        # treeview
        sortedItems = sorted(treeDic, key = treeDic.__getitem__, \
                             reverse = reverse)
        
        for indx, item in enumerate(sortedItems):
            self.trackerTV.move(item, '', indx)

        # Change the column heading to indicate the sorting direction and
        # reverse the direction of sorting in the next call
        symbol = ' (v)'
        
        if reverse:
            symbol = ' (^)'

        self.trackerTV.heading(colID, text = column + symbol, \
            command = lambda: self.sortColumn(column, colID, not reverse))


    def newProject(self):
        '''Removes all widget entries and resets the filename'''
        
        if self.saveProgress():
            self.fileName = ''
            self.progressSaved = True
            self.clearBoxes()
            self.trackerTV.delete(*self.trackerTV.get_children())


    def openProject(self):
        ''' '''
        if self.saveProgress():
            filetypes = (('CSV files', '*.csv'),("All files","*.*"))
            ext = '.csv'
            returnedName = filedialog.askopenfilename(parent = self.master, \
                                                      defaultextension = ext, \
                                                      filetypes = filetypes)
            
            if returnedName != '':
                self.fileName = returnedName
                self.progressSaved = True
                self.clearBoxes()
                self.trackerTV.delete(*self.trackerTV.get_children())
                self.openFile()


    def openFile(self):
        ''' '''
        try:
            with open(self.fileName, 'r') as csvFile:
                itemsReader = csv.reader(csvFile, delimiter=';')

                for item in itemsReader:
                    if len(item) == 5:
                        values = (item[0], item[1], item[2], item[3], item[4])
                        self.trackerTV.insert('', 'end', iid = item[0], \
                                              values = values)
        except IOError:
            messagebox.showerror('Cannot Open', '{} not found' \
                                 .format(self.fileName))
        except ValueError:
            messagebox.showerror('Reading Error', 'Unable to read file')
        except Exception as exception:
            messagebox.showerror('Unexpected Error', \
                                 'Exception: {}'.format(exception))
        
    def saveProject(self):
        ''' '''
        if self.fileName == '':
            self.saveAsProject()
        else:
            self.saveFile()
            self.progressSaved = True


    def saveFile(self):
        ''' '''
        try:
            with open(self.fileName, 'w') as csvFile:
                itemsWriter = csv.writer(csvFile, delimiter=';')
                for item in self.trackerTV.get_children(''):
                    itemsWriter.writerow(self.trackerTV.item(item, 'values'))
        except Exception as exception:
            messagebox.showerror('Unexpected Error', \
                                 'Exception: {}'.format(exception))


    def saveAsProject(self):
        '''Saves contents of the widgets to a specified file'''
        filetypes = (('CSV files', '*.csv'),("All files","*.*"))
        ext = '.csv'
        returnedName = filedialog.asksaveasfilename(parent = self.master, \
                                                    defaultextension = ext, \
                                                    filetypes = filetypes)

        if returnedName != '':
            self.fileName = returnedName
            self.saveProject()

    def saveProgress(self):
        '''Checks if a save is needed and saves
           Returns: (bool) Saving has been handled'''
        canContinue = True

        if not self.progressSaved:
            saveNeeded = messagebox \
                         .askyesno('Save Progress', \
                                   'Do you want to save your progress?')
            if saveNeeded:
                self.saveProject()
                canContinue = False

        return canContinue

    def deleteWindow(self):
        '''Final checks before closing window'''
        
        try:
            # Check if there are unsaved changes before closing
            if self.saveProgress():
                self.master.destroy()
                
        except Exception as exception:
            messagebox.showerror('Unexpected Error', \
                                 'Exception: {}'.format(exception))


    def separatorClick(self, event):
        if self.trackerTV.identify_region(event.x, event.y) == "separator":
            return 'break'

        
app = AssignmentTrackerApp()
app.mainloop()