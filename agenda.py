"""
Agenda is a simple program to make writing an Agenda easier

When run as a script from the commandline, agenda takes 'agenda_input.xlsx' excel file as 
an input that includes a start-time, session titles and durations and produces an output excel
file with the individual session times added.
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
        #to keep track of current time
        self.time_object_current = datetime.strptime(self.agenda_starttime, "%H:%M:%S")
        self.string_current_time = str(self.time_object_current.time())

    """
    def __init__(self, name, startTime):
        self.name = name

        #take a string as input for the time
        self.startingTime = startTime

        #convert string into a time object to keep track of current_time
        self.time_object_current = datetime.strptime(self.startingTime, "%H:%M:%S")
        self.string_current_time = str(self.time_object_current.time())

        #initialise an empty list and add the header info
        self.agendaList = []
        header_list = ['Index', 'Start-time', 'Endtime', 'Topic', 'Duration']
        self.agendaList.append(header_list)
    """
    """
    A function to add an AgendaItem with a name and duration. It is then stored
    in list with 5 items: Index, Starttime, Endtime, Topic, duration
    """  
    def addAgendaItem(self, name, duration):
      
        agendaItemList = []
        index = len(self.agendaList) - 1
        agendaItemList.append(index+1)

        #initialise the start time from the Agenda class current time
        agendaItemList.append(str(self.time_object_current.time()))

        #add the duration to get 'endtime' and update 'current_time'
        duration_to_add = int(duration)
        time_object_end = self.time_object_current + timedelta(minutes = int(duration_to_add))
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
                
        #print the header
        print("--------------------------------")
        print("Title: " + self.agenda_name)
        print("Start time: " + self.agenda_starttime)
        print("--------------------------------")

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
    def exportToExcel(self, filename):

        #initialist the workbook
        doc_file_name = str(filename)
        workbook = Workbook()
        sheet = workbook.active

        #add data to the workbook
        #sheet["A1"] = str(self.five_agenda_list)

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
        workbook = load_workbook(filename='agenda_input.xlsx')
        sheet = workbook.active

        #iterate through the sheet, return the cell values and add them to a value list and return
        value_list = []

        title = sheet["B2"].value
        value_list.append(title)
        start = sheet["B3"].value
        value_list.append(start)

        for row in sheet.iter_rows(min_row = 5, max_row = 14, min_col = 1, max_col = 2, values_only = True):

            value = row[0]
            value_list.append(value)
            value2 = str(int(row[1]))
            value_list.append(value2)

        return value_list
            

if __name__ =="__main__":

"""
initialise a new agenda object
read from excel
set agenda title and start time from excel
add all agenda items (this could be its own function)
print agenda to commandline
save to excel
"""

    #agenda1 = Agenda("Test Agenda", "10:00:00")
    agenda1 = Agenda()

    #populate a list with session titles and durations from the excel input file
    list = agenda1.inputFromExcel("agenda_input.xlsx")
    print(list)

#set the title and starting
agenda1.agenda_name = list[0]
agenda1.agenda_starttime = list[1]

    #iterate through the list and add the titles and durations to the Agenda using 'addAgendaItem' function
    list_index = 0
    while list_index < len(list):
        agenda1.addAgendaItem(list[list_index], (list[list_index+1]))
        list_index = list_index + 2

    #print the agenda to the commandline but also export to excel to allow copy and paste to another table
    agenda1.printAgenda()
    agenda1.exportToExcel("agenda_output.xlsx")
    
    