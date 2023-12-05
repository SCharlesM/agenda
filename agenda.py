## Agenda is a program to make writing
## an agenda easier. Input agenda item names and
## times and have an easy way to change times
"""
Current program runs with no errors!
I need to add
1 - finish print funtion
2 - edit AgendaItem function, edit name, length, position etc
other features:
-print to CSV to copy into Excel/Word
-export straight to Excel?
-imports topic names from csv/Excel?

"""
from datetime import datetime, time, timedelta, date
#import csv.py

class Agenda:
    def __init__(self, name, startTime):
        self.name = name
        self.agendaList = []
        
        ##take a string as input for the time
        self.startingTime = startTime

    def addAgendaItem(self, name, duration):
        """
        a function to add an AgendaItem, as a list of 2 variables, name and duration
        """        
        agendaItemList = []
        agendaItemList.append(name)
        agendaItemList.append(duration)
       
        #add these items lists to the main agenda list
        self.agendaList.append(agendaItemList)

    def printAgenda(self):
        """
        a function to print the agenda by moving through the
        index and printing each item and adding on the
        duration to make a starttime and endtime
        """
        print("--------------------------------")
        print("Title: " + self.name)
        print("Start time: " + self.startingTime)
        print("--------------------------------")
        s1 = "Index"
        s2 = "Startime"
        s3 = "Endtime"
        s4 = "Topic"
        s5 = "Duration"
        print( '{:<10s} {:<10s} {:<10s} {:<15s} {:<10s} '.format(s1, s2, s3, s4, s5))
        #print(self.agendaList)

        string_starttime = self.startingTime

        #iterate throught the agendaList and print an index, starttime, endtime, name and duration
        i = 0
        while i < len(self.agendaList):
            
            newList = self.agendaList[i]
            string_index = str(i+1)

            #convert string to a time object and format. add the duration.
            time_object_start = datetime.strptime(string_starttime, "%H:%M:%S")
            duration_to_add = newList[1]
            time_object_end = time_object_start + timedelta(minutes = int(duration_to_add))
            string_endtime = str(time_object_end.time())

            string_topic = newList[0]
            string_duration = newList[1]

            #format the strings, left align 20, left align 10 (string)
            print( '{:<10s} {:<10s} {:<10s} {:<15s} {:<10s}'.format(string_index, string_starttime, string_endtime, string_topic, string_duration))
            
            string_starttime = string_endtime

            i += 1

"""  
    def editAgendaItem():

        a function to edit an AgendaItem, including Title and duration
    """
"""
    def exportToFile():
    
        A function to export the Agenda to another document potentially to print/copy
    """
"""
#main function
newName=input("Enter the Agenda title: ")
newTime=input("Enter the Agenda start time: ")
agenda1 = Agenda(newName, newTime)
    
while True:
    newItem=input("Enter a new Agenda item: ")
  
#exit the loop when user presses Enter with no input
    if newItem == "":
        print("User pressed Enter")
        break

    newTime=input("How long is this new Agenda item: ")
    
    agenda1.addAgendaItem(newItem, newTime)
    agenda1.printHeader()
    agenda1.printAgenda()
"""