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
    agenda_starttime = is an empty time object
    agenda_list will be used to hold the topics and durations of each item

    """
    def __init__(self):
        
        self.agenda_name = ""           
        self.agenda_starttime = "00:00:00"
        self.agenda_list = []

        #convert the agenda_starttime string into a time object 
        #to keep track of cumulative time
        #self.time_object_current = datetime.strptime(self.agenda_starttime, "%H:%M:%S")
        #self.string_current_time = str(self.time_object_current.time())
    """
    A function to populate the agenda_list from the excel_input. Creat a dictionary
    with data; index, topic, duration, starttime, endtime, 
    """
    def populate_agenda(self, excel_input) :

        index = 1   
        while index < len(excel_input) :        #re-write with for loop?

            topic_title, topic_duration = excel_input[index - 1]

            agenda_dict = dict(index = index, topic = topic_title, duration = topic_duration)
            agenda_dict["start"] = str(self.time_object_current.time())
            self.time_object_current += timedelta(minutes = topic_duration)
            agenda_dict["end"] = str(self.time_object_current.time())

            self.agenda_list.append(agenda_dict)
            index += 1

        print(self.agenda_list)

    """
    A function to store the data from the Excel document (topic and duration) and 
    calculate the start-time and end-time for each item. And update the current time. 
    The data stored in the list is: Index, Starttime, Endtime, Topic, duration.
    """  
    def addAgendaItem(self, name, duration):
      
        agenda_item = []
        index = len(self.agenda_list)
        agenda_item.append(index+1)

        #using the cumulative time tracker as next items start-time
        agenda_item.append(str(self.time_object_current.time()))

        #add the duration to get 'endtime' and update 'current_time'
        #duration_to_add = int(duration)
        time_object_end = self.time_object_current + timedelta(minutes = int(duration))
        self.time_object_current = time_object_end

        agenda_item.append(str(time_object_end.time()))
        agenda_item.append(name)
        agenda_item.append(duration)
       
        #add agenda_item to the agenda_list
        self.agenda_list.append(agenda_item)

    """
    a print function to print a header and iterate through the agenda_list to print each item

    """
    def printAgenda(self):
                
        #print the Title and Start time
        print("--------------------------------")
        print("Title: " + self.agenda_name)
        print("Start time: " + self.agenda_starttime)
        print("--------------------------------")
        
        #print the column headings
        print( '{:<10s} {:<10s} {:<10s} {:<15s} {:<10s}'.format('Index', 'Start', 'End', 'Title', 'duration'))
        
        for object in self.agenda_list :
            s1 = str(object['index'])
            s2 = object['start']
            s3 = object['end']
            s4 = object['topic']
            s5 = str(object['duration'])

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
        doc_file_name = agenda_name_no_space + timestamp + '.xlsx'
        
        #set A1 as 'Title' and A2 as the agenda_title
        #set B1 as 'starttime' and B2 as the agenda_start
        sheet['A1'] = 'Title'
        sheet['B1'] = self.agenda_name
        sheet['A2'] = 'Starttime'
        sheet['B2'] = self.agenda_starttime

        #iterate though the excel cells to store the afenda item data into each cell
        for item in self.agenda_list :
            
            i = 0
            for x,y in item :
                cell = string.ascii_uppercase[0] + str(1)
                sheet[cell] = item[i]
        """
        #loop through each agenda item contained in agenda_list to create a temp list.
        #Loop through the excel cell descriptors (A1, B1, C1...) to transfer
        #the agenda item data into each row in the excel sheet
        for l in range(0 , len(self.agenda_list)):

            agenda_item = self.agenda_list[l]

            for k in range(0, len(agenda_item)):
                cell_description = (string.ascii_uppercase[k] + str(l+1))
                sheet[cell_description] = agenda_item[k]
        """
        #save to the workbook using the filename argument
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
        self.agenda_starttime = str(sheet["B3"].value)   #does it matter if this is a date.time object?             

        #new_title, new_starttime = value_list[0]    #unpack the values (use later)

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

    #iterate through the list and add the titles and durations to the Agenda using 'addAgendaItem' function
    """
    list_index = 2
    while list_index < len(excel_input):
        agenda1.addAgendaItem(excel_input[list_index], (excel_input[list_index+1]))
        list_index = list_index + 2
    """
    #this whole loop could be written as
    agenda1.populate_agenda(excel_input)

    #print the agenda to the commandline but also export to excel to allow copy and paste to another table
    agenda1.printAgenda()
    agenda1.exportToExcel()
