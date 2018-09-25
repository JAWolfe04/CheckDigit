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
        self.topFrame = tk.Frame(self.master)

        # Initialize file name and save progress
        # for current state
        self.fileName = ''
        self.progressSaved = True

        # Create and position widgets and window
        self.createWindowWidgets(self.topFrame)
        self.createKeyboardSupport()
        self.createTrackerTV(self.master)
        self.positionWidgets(self.master)        
        self.createFileMenu(self.master)
        self.centerWindow(self.master)


    def createWindowWidgets(self, frame):
        '''Create widgets for entering info into the tracker
           Args: frame (tkinter.Tk): Frame for widgets'''
        self.AssignmentLabel = tk.Label(frame, text="Assignment Number")
        self.AssignmentBox = tk.Entry(frame)
        self.StudentLabel = tk.Label(frame, text='Student Name')
        self.StudentBox = tk.Entry(frame)
        self.GradeLabel = tk.Label(frame, text='Assignment Grade')
        self.GradeBox = tk.Entry(frame)
        self.AddEntryButton = tk.Button(frame, text="Add/Edit", width = 10, command=self.addEntry)
        self.AddEntryButton.bind('<Return>', lambda e: self.addEntry())
        self.clearBoxesButton = tk.Button(frame, text = 'Clear', width = 10, command=self.clearBoxes)
        self.clearBoxesButton.bind('<Return>', lambda e: self.clearBoxes())
        self.removeEntryButton = tk.Button(frame, text = 'Remove', width = 10, command = self.removeEntry)
        self.removeEntryButton.bind('<Return>', lambda e: self.removeEntry())
        self.getEntryButton = tk.Button(frame, text = 'Retrieve', width = 10, command = self.getAssignment)
        self.getEntryButton.bind('<Return>', lambda e: self.getAssignment())
        self.randomButton = tk.Button(frame, text = 'Random', width = 10, command = self.getRandom)
        self.randomButton.bind('<Return>', lambda e: self.getRandom())
        self.AssignmentBox.focus()


    def createKeyboardSupport(self):
        '''Binds for navigating widgets using keyboard only'''
        self.AssignmentBox.bind('<Return>', lambda e: self.StudentBox.focus())
        self.AssignmentBox.bind('<Down>', lambda e: self.StudentBox.focus())
        self.StudentBox.bind('<Up>', lambda e: self.AssignmentBox.focus())
        self.StudentBox.bind('<Return>', lambda e: self.GradeBox.focus())
        self.StudentBox.bind('<Down>', lambda e: self.GradeBox.focus())
        self.GradeBox.bind('<Up>', lambda e: self.StudentBox.focus())
        self.GradeBox.bind('<Return>', lambda e: self.AddEntryButton.focus())
        self.GradeBox.bind('<Down>', lambda e: self.AddEntryButton.focus())
        self.AddEntryButton.bind('<Up>', lambda e: self.randomButton.focus())
        self.AddEntryButton.bind('<Left>', lambda e: self.GradeBox.focus())
        self.AddEntryButton.bind('<Right>', lambda e: self.getEntryButton.focus())
        self.clearBoxesButton.bind('<Down>', lambda e: self.getEntryButton.focus())
        self.clearBoxesButton.bind('<Left>', lambda e: self.randomButton.focus())
        self.randomButton.bind('<Left>', lambda e: self.AssignmentBox.focus())
        self.randomButton.bind('<Down>', lambda e: self.AddEntryButton.focus())
        self.randomButton.bind('<Right>', lambda e: self.clearBoxesButton.focus())
        self.getEntryButton.bind('<Up>', lambda e: self.clearBoxesButton.focus())
        self.getEntryButton.bind('<Left>', lambda e: self.AddEntryButton.focus())
        self.getEntryButton.bind('<Right>', lambda e: self.removeEntryButton.focus())
        self.removeEntryButton.bind('<Left>', lambda e: self.getEntryButton.focus())


    def createTrackerTV(self, win):
        '''Create the tracker of assignments
           Args: win (tkinter.Tk): Root window'''
        columns = ('Assignment', 'Class', 'Type', 'Student', 'Grade')
        self.trackerTV = ttk.Treeview(win, columns = columns, selectmode = 'browse', show = 'headings')
        self.trackerTV.heading('#1', text='Assignment', command = lambda: self.sortColumn('Assignment', '#1', False))
        self.trackerTV.heading('#2', text = 'Class', command = lambda: self.sortColumn('Class', '#2', False))
        self.trackerTV.heading('#3', text = 'Type', command = lambda: self.sortColumn('Type', '#3', False))
        self.trackerTV.heading('#4', text = 'Student', command = lambda: self.sortColumn('Student', '#4', False))
        self.trackerTV.heading('#5', text = 'Grade', command = lambda: self.sortColumn('Grade', '#5', False))
        self.trackerTV.column('#1', minwidth = 125, width = 125, anchor = tk.W)
        self.trackerTV.column('#2', minwidth = 125, width = 125, anchor = tk.W)
        self.trackerTV.column('#3', minwidth = 125, width = 125, anchor = tk.W)
        self.trackerTV.column('#4', minwidth = 125, width = 125, anchor = tk.W)
        self.trackerTV.column('#5', minwidth = 125, width = 125, anchor = tk.W)
        self.scrollbar = tk.Scrollbar(win, orient = "vertical", command = self.trackerTV.yview)
        self.trackerTV.configure(yscrollcommand = self.scrollbar.set)
        self.trackerTV.bind('<Button>', self.separatorClick)
        self.trackerTV.bind('<ButtonRelease>', self.separatorClick)
        self.trackerTV.bind('<Motion>', self.separatorClick)


    def positionWidgets(self, win):
        '''Position all widgets
           Args: win (tkinter.Tk): Root window'''
        self.AssignmentLabel.grid(row = 0, column = 0, sticky = tk.E)
        self.AssignmentBox.grid(row = 0, column = 1, sticky = tk.EW)
        self.StudentLabel.grid(row = 1, column = 0, sticky = tk.E)
        self.StudentBox.grid(row = 1, column = 1, sticky = tk.EW)
        self.GradeLabel.grid(row = 2, column = 0, sticky = tk.E)
        self.GradeBox.grid(row = 2, column = 1, sticky = tk.EW)
        self.clearBoxesButton.grid(row = 0, column = 3)
        self.randomButton.grid(row = 0, column = 2)
        win.grid_rowconfigure(1, minsize = 25)
        self.AddEntryButton.grid(row = 2, column = 2)
        self.getEntryButton.grid(row = 2, column = 3)
        self.removeEntryButton.grid(row = 2, column = 4)
        self.topFrame.grid(sticky = tk.EW)
        self.topFrame.grid_columnconfigure(2, minsize = 125)
        self.topFrame.grid_columnconfigure(3, minsize = 125)
        self.topFrame.grid_columnconfigure(4, minsize = 125)
        self.trackerTV.grid(row = 2, column = 0)
        self.scrollbar.grid(row = 2, column  = 1, sticky = tk.NS)


    def createFileMenu(self, win):
        '''Setup file menu
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
            win.geometry('{}x{}+{}+{}'.format(winWidth, winHeight, xOffset, yOffset))            
        except Exception as exception:
            messagebox.showerror('Unexpected Error', 'Unable to center window \nException: {}'.format(exception))


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
            messagebox.showerror('No Selection', 'No items have been selected to retrieve')


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
                    self.trackerTV.insert('', 'end', iid=validateStr, values = values)
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
            messagebox.showerror('Invalid Assignment Number', 'Assignment Number must be 13 characters in length')
        elif not validateStr.isalnum():
            messagebox.showerror('Invalid Assignment Number', 'Assignment Number must consist of 0 through 9 and A through F')
        elif validateStr[0] not in 'ABCD':
            messagebox.showerror('Invalid Assignment Number', 'The first digit must be a valid class identifier, A through D.')
        elif validateStr[1] not in 'ABCDE':
            messagebox.showerror('Invalid Assignment Number', 'The second digit must be a valid assignment type identifier, A through E.')
        elif not validateStr[-1].isdigit():
            messagebox.showerror('Invalid Assignment Number', 'Last digit must be a number')
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
                
            elif validateStr[weight] in 'ABCDEF':
                numberSum += (ord(validateStr[weight]) - 55) * weight
                
            else:
                messagebox.showerror('Invalid Assignment Number', 'String contains invalid characters "{}"'.format(validateStr[weight]))
                break
            
        # Check the sum to the check digit
        else:            
            if (numberSum % 10) == int(validateStr[-1]):
                validEntry = True
                
            else:
                messagebox.showerror('Invalid Assignment Number', 'A character is entered incorrectly or transposed.')

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
            canDelete = messagebox.askyesno('Delete Entry', 'Do you want to delete \n {}'.format(selectedItems[0]))
                
            if canDelete:
                self.trackerTV.delete(selectedItems[0])
                
                if self.progressSaved:
                    self.progressSaved = False
                
        except IndexError:
            messagebox.showerror('No Selection', 'No items have been selected to remove')


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
        randomStr += random.choice('ABCD')
        randomStr += random.choice('ABCDE')

        for indx in range(0,10):
            randomStr += random.choice('ABCDEF0123456789')

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
        sortedItems = sorted(treeDic, key = treeDic.__getitem__, reverse = reverse)
        
        for indx, item in enumerate(sortedItems):
            self.trackerTV.move(item, '', indx)

        # Change the column heading to indicate the sorting direction and
        # reverse the direction of sorting in the next call
        symbol = ' (v)'
        
        if reverse:
            symbol = ' (^)'

        self.trackerTV.heading(colID, text = column + symbol, command = lambda: self.sortColumn(column, colID, not reverse))


    def newProject(self):
        '''Removes all widget entries and resets the filename'''
        
        if self.saveProgress():
            self.fileName = ''
            self.progressSaved = True
            self.clearBoxes()
            self.trackerTV.delete(*self.trackerTV.get_children())


    def openProject(self):
        '''Handles the behavior of the open button in the file menu'''
        if self.saveProgress():
            filetypes = (('CSV files', '*.csv'),("All files","*.*"))
            returnedName = filedialog.askopenfilename(parent = self.master, defaultextension = '.csv', filetypes = filetypes)
            
            if returnedName != '':
                self.fileName = returnedName
                self.progressSaved = True
                self.clearBoxes()
                self.trackerTV.delete(*self.trackerTV.get_children())
                self.openFile()


    def openFile(self):
        '''Opens the predesignated file into the tracker'''
        try:
            with open(self.fileName, 'r') as csvFile:
                itemsReader = csv.reader(csvFile, delimiter=';')

                for item in itemsReader:
                    if len(item) == 5:
                        self.trackerTV.insert('', 'end', iid = item[0], values = (item[0], item[1], item[2], item[3], item[4]))
        except IOError:
            messagebox.showerror('Cannot Open', '{} not found'.format(self.fileName))
        except ValueError:
            messagebox.showerror('Reading Error', 'Unable to read file')
        except Exception as exception:
            messagebox.showerror('Unexpected Error', 'Exception: {}'.format(exception))
        
        
    def saveProject(self):
        '''Handles the behavior of the save button in the file menu'''
        if self.fileName == '':
            self.saveAsProject()
        else:
            self.saveFile()
            self.progressSaved = True


    def saveFile(self):
        '''Saves the contents of the tracker to a predesignated file'''
        try:
            with open(self.fileName, 'w') as csvFile:
                itemsWriter = csv.writer(csvFile, delimiter=';')
                for item in self.trackerTV.get_children(''):
                    itemsWriter.writerow(self.trackerTV.item(item, 'values'))
        except Exception as exception:
            messagebox.showerror('Unexpected Error', 'Exception: {}'.format(exception))


    def saveAsProject(self):
        '''Handles the behavior of the save as button in the file menu'''
        filetypes = (('CSV files', '*.csv'),("All files","*.*"))
        returnedName = filedialog.asksaveasfilename(parent = self.master, defaultextension = '.csv', filetypes = filetypes)

        if returnedName != '':
            self.fileName = returnedName
            self.saveProject()


    def saveProgress(self):
        '''Checks if a save is needed and saves
           Returns: (bool) Saving has been handled'''
        canContinue = True

        if not self.progressSaved:
            saveNeeded = messagebox.askyesno('Save Progress', 'Do you want to save your progress?')
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
            messagebox.showerror('Unexpected Error', 'Exception: {}'.format(exception))


    def separatorClick(self, event):
        '''Disables the event when it is in the treeview separator
           Fix for the separator of the treeview being able to be carried off screen
           Returns: (str) "break" if it is in the separator region'''
        if self.trackerTV.identify_region(event.x, event.y) == "separator":
            return 'break'
        

# Call and run the app        
app = AssignmentTrackerApp()
app.mainloop()