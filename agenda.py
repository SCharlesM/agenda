"""
Agenda is a simple program to make writing an Agenda easier

agenda.py takes 'agenda_input.xlsx' excel file as an input that includes a Title, a start-time and a list 
of topics and durations. agenda.py converts this into an agenda by adding a individual start and end time for
each topic based on the duration of each topic. It outputs the resulting agenda to "agenda_output*.xlsx"
"""

from datetime import datetime, time, timedelta, date
from openpyxl import Workbook, load_workbook
import string

class Agenda:

    """
    initialise a new agenda with initial variables
    agenda_name is a string
    agenda_starttime = is an empty time object
    agendaList will be used to hold the topics and durations of each item

    """
    def __init__(self):
        
        self.agenda_name = ""
        self.agenda_starttime = "00:00:00"
        self.agendaList = []

        #convert the agenda_starttime string into a time object 
        #to keep track of cumulative time
        self.time_object_current = datetime.strptime(self.agenda_starttime, "%H:%M:%S")
        self.string_current_time = str(self.time_object_current.time())

    """
    A function to store the data from the Excel document (topic and duration) and 
     calculate the start-time and end-time for each item. And update the 
      current time. The data stored ins the list is: Index, Starttime, Endtime, Topic, duration.
    """  
    def addAgendaItem(self, name, duration):
      
        agendaItemList = []
        index = len(self.agendaList)
        agendaItemList.append(index+1)

        #using the cumulative time tracker as next items start-time
        #self.time_object_current = datetime.strptime(self.agenda_starttime, "%H:%M:%S") #this strips the start-time everytime
        agendaItemList.append(str(self.time_object_current.time()))

        #add the duration to get 'endtime' and update 'current_time'
        #duration_to_add = int(duration)
        time_object_end = self.time_object_current + timedelta(minutes = int(duration))
        self.time_object_current = time_object_end

        agendaItemList.append(str(time_object_end.time()))
        agendaItemList.append(name)
        agendaItemList.append(duration)
       
        #add agendaItemList to the agendaList
        self.agendaList.append(agendaItemList)

    """
    a print function to print a header and iterate through the agendaList to print each item

    """
    def printAgenda(self):
                
        #print the Title and Start time
        print("--------------------------------")
        print("Title: " + self.agenda_name)
        print("Start time: " + self.agenda_starttime)
        print("--------------------------------")
        
        #print the column headings
        print( '{:<10s} {:<10s} {:<10s} {:<15s} {:<10s}'.format('Index', 'Start', 'End', 'Title', 'duration'))

        #iterate throught the agendaList and print each individual list
        for j in range(0, len(self.agendaList)):

            new_list = self.agendaList[j]

            string_index = str(new_list[0])
            string_starttime = new_list[1]
            string_endtime = new_list[2]
            string_topic = new_list[3]
            string_duration = new_list[4]
        
            #format the strings, left align 20, left align 10 (string)
            print( '{:<10s} {:<10s} {:<10s} {:<15s} {:<10s}'.format(string_index, string_starttime, string_endtime, string_topic, string_duration))

    """
    a function to change title of an AgendaItem, using the index of that AgendaItem

    """
    def editAgendaItemTitle(self, index, new_title):

        temp_list = self.agendaList[index]
        temp_list[3] = new_title

    """
    a function to change duration of an AgendaItem, using the index of that item

    """
    def editAgendaItemDuration(self, index, new_duration):

        temp_list = self.agendaList[index-1]
        temp_list[4] = new_duration

    """
        A function to export the Agenda to an Excel document

    """
    def exportToExcel(self):

        #initialist the workbook
        doc_file_name = self.agenda_name
        workbook = Workbook()
        sheet = workbook.active

        new_list = self.agendaList[0]

        #populate a list with the uppercase letters
        uppercase_list = []
        for j in string.ascii_uppercase:
            uppercase_list.append(j)

        for l in range(0 , len(self.agendaList)):

            new_list = self.agendaList[l]

            for k in range(0, 5):
                temp_string = (uppercase_list[k] + str(l+1))
                sheet[temp_string] = new_list[k]

        #save to the workbook using the filename argument
        workbook.save(filename=doc_file_name)

        #confirm workbook has been saved
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
    list_index = 2
    while list_index < len(list):
        agenda1.addAgendaItem(list[list_index], (list[list_index+1]))
        list_index = list_index + 2

    #print the agenda to the commandline but also export to excel to allow copy and paste to another table
    agenda1.printAgenda()
    agenda1.exportToExcel()