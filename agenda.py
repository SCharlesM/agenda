## Agenda is a program to make writing
## an agenda easier. Input agenda item names and
## times and have an easy way to change times
"""
Current program runs with no errors!
I need to add
-a currentTime variable to keep each item adding to
at the moment they only print their own time compared to the start
-i need an exception handler for the time format at the start
-i need an edit function that changes the length of an item
-this edit function could also change
-would also be good to put this on gitHub to track changes?

"""
from datetime import datetime, time, timedelta, date
#import csv.py

class Agenda:
    def __init__(self, name, startTime):
        self.name = name
        self.agendaList = []
        self.agendaTimes = []

        ##take a string as input for the time
        self.startTime = startTime
        
        ##convert string to a time object
        self.to = datetime.strptime(self.startTime, "%H:%M")

        ##variable to track currentTime
        self.currentTime = self.to

        ##strip out the time, not the date
        self.only_t = self.currentTime.time()
        
        self.agendaTimes.append(self.only_t)
        
    def addAgendaItem(self, name, time):
        """
        a function to add an AgendaItem, with a name and time as input
        
        """
        #convert the string input into an integer
        intLength = int(time)
        #add the length in minutes to the time object
        self.currentTime = self.currentTime + timedelta(minutes=intLength)

        #strip out just the time, not the date
        only_time = self.currentTime.time()
        
        self.agendaList.append(name)
        self.agendaTimes.append(only_time)  

    def printHeader(self):
        """
        a function to print the header of the agenda
        """
        print("--------------------------------")
        print("Title: " + self.name)
        print("Start time: " + str(self.startTime))
        print("--------------------------------")
        
    def printAgenda(self):
        """
        a function to print the full agenda eg
        ["item1", "item2" , "item3"]
        [9.30, 10.00, 11.45, 12.00]
        """
        i=0
        while i < len(self.agendaList):
            print(self.agendaTimes[i], end=" ")
            print(self.agendaList[i])
            i+=1
            
    def changeAgendaItem():
        """
        a function to edit an AgendaItem, including Title and length
        """
    def exportToFile():
        """
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

