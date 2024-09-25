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
        
        #self.agenda_name = ""           #do i need to initialise these?
        #self.agenda_starttime = "00:00:00"
        self.agenda_list = []

        #convert the agenda_starttime string into a time object 
        #to keep track of cumulative time
        #self.time_object_current = datetime.strptime(self.agenda_starttime, "%H:%M:%S")
        #self.string_current_time = str(self.time_object_current.time())
    """
    A function to populate the agenda_list from the excel_input. Creat a dictionary
    with data; index, starttime, endtime, topic and duration
    """
    def populate_agenda(self, excel_input) :
        #pass

        index = 1   #index needs to only go up 1
                    #there needs to be another iterator that goes up 2
        while index < len(excel_input) :
            agenda_dict = dict(index = index, duration = 0, start = 0, end = 0, topic = "")
            agenda_dict["duration"] = excel_input[index+2]
            agenda_dict["start"] = str(self.time_object_current.time())
            self.time_object_current += timedelta(minutes = int(excel_input[index+2]))
            agenda_dict["end"] = str(self.time_object_current.time())
            agenda_dict["topic"] = excel_input[index+1]
            print(agenda_dict)

            self.agenda_list.append(agenda_dict)
            index += 2

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

        #iterate throught the agenda_list and print each individual list
        for j in range(0, len(self.agenda_list)):

            new_list = self.agenda_list[j]

            string_index = str(new_list[0])
            string_starttime = new_list[1]
            string_endtime = new_list[2]
            string_topic = new_list[3]
            string_duration = new_list[4]
        
            #format the strings, left align 20, left align 10 (string)
            print( '{:<10s} {:<10s} {:<10s} {:<15s} {:<10s}'.format(string_index, string_starttime, string_endtime, string_topic, string_duration))

    """
        A function to export the Agenda to an Excel document
    """
    def exportToExcel(self):

        #initialist the workbook
        workbook = Workbook()
        sheet = workbook.active

        #change the whitespace in agenda_name to '_' and use title and timestamp as output filename
        agenda_name_no_space = self.agenda_name.replace(" ", "_")
        timestamp = strftime("_%d_%m_%Y_%I.%M%p", localtime())
        doc_file_name = agenda_name_no_space + timestamp + '.xlsx'
        
        #loop through each agenda item contained in agenda_list to create a temp list.
        #Loop through the excel cell descriptors (A1, B1, C1...) to transfer
        #the agenda item data into each row in the excel sheet
        for l in range(0 , len(self.agenda_list)):

            agenda_item = self.agenda_list[l]

            for k in range(0, len(agenda_item)):
                cell_description = (string.ascii_uppercase[k] + str(l+1))
                sheet[cell_description] = agenda_item[k]

        #save to the workbook using the filename argument
        file_path = 'C:\\Users\\Steve\Documents\\Coding\\python\\agenda_project\\agenda\\outputs\\'
        workbook.save(file_path + doc_file_name)

        #output to terminal to confirm workbook has been saved
        print("\nData has been exported and saved with filename: " + doc_file_name)

    def inputFromExcel(self, filename):

        #initialise the workbook with filename
        workbook = load_workbook(filename)
        sheet = workbook.active

        #iterate through the sheet, return the cell values and add them to a value list and return
        value_list = []

        title = sheet["B2"].value
        value_list.append(title)
        start = sheet["B3"].value
        value_list.append(str(start))

        for row in sheet.iter_rows(min_row = 5, max_row = 14, min_col = 1, max_col = 2, values_only = True):

            value = row[0]
            value_list.append(value)
            value2 = str(int(row[1]))
            value_list.append(value2)

        return value_list
            
if __name__ =="__main__":

    agenda1 = Agenda()

    #populate a list with session titles and durations from the excel input file
    excel_input = agenda1.inputFromExcel("agenda_input.xlsx")

    #set the title and starting and set the time object tracking cumulative time
    agenda1.agenda_name = excel_input[0]
    agenda1.agenda_starttime = excel_input[1]
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
