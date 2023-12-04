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
        
        ##take a string as input for the time
        self.startTime = startTime
        
        ##convert string to a time object
        self.time_object = datetime.strptime(self.startTime, "%H:%M")

        ##variable to track currentTime
        ##im not sure i need this
        self.currentTime = self.time_object

        ##strip out the time, and remove date
        self.startingTime = self.currentTime.time()
        
"""
##initialise a class with name and start time
    def __init__(self, name, startTime):
        self.name = name
        self.startTime 
"""
    def addAgendaItem(self, name, duration):
        """
        a function to add an AgendaItem, as a list of 2 variables, name and duration
        """
        #a list to store agendaItem name an duration
        #agendaItemList = []

        #convert the string input into an integer
       # intDuratoon = int(duration)
        
        agendaItemList[name, duration]
        agendaList.append(agendaItemList)

    def printAgenda(self):
        """
        a function to print the agenda by moving through the
        index and printing each item and adding on the
        duration to make a starttime and endtime
        """
        print("--------------------------------")
        print("Title: " + self.name)
        print("Start time: " + str(self.startingTime))
        print("--------------------------------")
        
        #variables to hold starttime and endtime
        st = self.startingTime
        et = 0

        i=0
        while i<len(self.agendaList):
          #  print((i+1) + "  .  " + st + " -  " + st + 
          print



        
        """    
    def printAgenda(self):
        
        a function to print the full agenda eg
        ["item1", "item2" , "item3"]
        [9.30, 10.00, 11.45, 12.00]
        """
        i=0
        while i < len(self.agendaList):
            print(self.agendaTimes[i], end=" ")
            print(self.agendaList[i])
            i+=1
      """      
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

