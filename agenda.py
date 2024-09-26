"""
Agenda is a simple program to make writing an Agenda easier

agenda.py takes 'agenda_input.xlsx' excel file as an input that includes a Title, a start-time and a list 
of topics and durations. agenda.py converts this into an agenda by adding a individual start and end time for
each topic based on the duration of each topic. It outputs the resulting agenda to "agenda_output*.xlsx"
"""

from datetime import datetime, time, timedelta
from time import strftime, localtime
from openpyxl import Workbook, load_workbook
import string

class Agenda:

    """
    initialise a new agenda with initial variables
    agenda_name is a string
    agenda_starttime is an empty time object
    agenda_list will be used to hold the dictionarys with each agenda item data
    """
    def __init__(self):
        
        self.agenda_name = ""           
        self.agenda_starttime = "00:00:00"
        self.agenda_list = []

    """
    A function to populate the agenda_list from the excel sheet. Create a dictionary
    with all the data: index, start, end, topic, duration
    """
    def populate_agenda(self, excel_input) :

        index = 1   
        while index < len(excel_input) :        #re-write with 'for' loop?

            #extract the tuples from the excel_input and initialise the dictionary
            topic_title, topic_duration = excel_input[index - 1]
            agenda_dict = dict(index = str(index), start = "", end = "", topic = topic_title, duration = str(topic_duration))
            
            #add the start and end times using the time objects
            agenda_dict["start"] = str(self.time_object_current.time())
            self.time_object_current += timedelta(minutes = topic_duration)
            agenda_dict["end"] = str(self.time_object_current.time())

            self.agenda_list.append(agenda_dict)
            index += 1

    """
    a print function to print a header and iterate through the agenda_list to print each item
    """
    def printAgenda(self):

        print("----------------------------------------------------------")
        print("Title: ", self.agenda_name, "\nStart time: ", self.agenda_starttime)
        print("----------------------------------------------------------")
        
        #print the column headings
        print( '{:<10s} {:<10s} {:<10s} {:<15s} {:<10s}'.format('Index', 'Start', 'End', 'Title', 'duration'))
        
        for object in self.agenda_list :
            s1, s2, s3, s4, s5 = object.values()
            
            #format the strings, left align 20, left align 10 (string)
            print( '{:<10s} {:<10s} {:<10s} {:<15s} {:<10s}'.format(s1, s2, s3, s4, s5))  

    """
        A function to export the Agenda to an Excel document
    """
    def exportToExcel(self):

        #initialist the workbook
        workbook = Workbook()
        sheet = workbook.active

        #change the whitespace in agenda_name to '_' and use name and timestamp as output filename
        agenda_name_no_space = self.agenda_name.replace(" ", "_")
        timestamp = strftime("_%d_%m_%Y_%I.%M%p", localtime())
        doc_file_name = agenda_name_no_space, timestamp, '.xlsx'
        
        #set the Title and Start-time at the top of the sheet
        sheet['A1'] = 'Title:'
        sheet['B1'] = self.agenda_name
        sheet['A2'] = 'Start-time:'
        sheet['B2'] = self.agenda_starttime
        sheet['A4'] = 'Index'                       #can this be done any better using the keys?    
        sheet['B4'] = 'Start'
        sheet['C4'] = 'End'
        sheet['D4'] = 'Topic'
        sheet['E4'] = 'Duration'      

        #iterate though the excel cells to store the agenda item data into each cell
        #item is the dictionary, i need o unpack the values and store in excel

        j = 5                          #to start at row 5 in excel sheet
        for item in self.agenda_list :

            i = 0                       #to start at letter A
            for key in item :
                cell = string.ascii_uppercase[i] + str(j)
                sheet[cell] = item[key]
                i += 1
            j += 1

        #save to the workbook using the filename argument. 
        file_path = 'C:\\Users\\Steve\Documents\\Coding\\python\\agenda_project\\agenda\\outputs\\'
        workbook.save(file_path + doc_file_name)

        #output to terminal to confirm workbook has been saved
        print("\nData has been exported and saved with filename: " + doc_file_name)

        """
        a function to extract data from the 'input.xlsx' sheet, including Title, starttime and for each agenda
        item the session title and duration. Returns a list of tuples.
        """
    def inputFromExcel(self, filename):

        #initialise the workbook with filename and initialise an empty list
        workbook = load_workbook(filename)
        sheet = workbook.active
        value_list = []

        #set the title and startime of the agenda
        self.agenda_name = sheet["B2"].value        
        self.agenda_starttime = str(sheet["B3"].value)

        #iterate through the rows and add the tuple of values to the list and return the list
        for row in sheet.iter_rows(min_row = 5, max_row = 14, min_col = 1, max_col = 2, values_only = True):

            value_list.append(row)

        return value_list
            
if __name__ =="__main__":

    agenda1 = Agenda()

    #populate a list with session titles and durations from the excel input file
    excel_input = agenda1.inputFromExcel("agenda_input.xlsx")

    #format the starttime date.time object
    agenda1.time_object_current = datetime.strptime(agenda1.agenda_starttime, "%H:%M:%S")

    #populate the agenda_list with the dictionaries containing all the agenda items
    agenda1.populate_agenda(excel_input)

    #print the agenda to the commandline but also export to excel
    agenda1.printAgenda()
    agenda1.exportToExcel()