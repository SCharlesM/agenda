"""
a second version of main to practice loading/saving to CSV file

"""
from agenda import Agenda
import csv

"""
open the saved data from csv files and print it
as a test for importing data
"""
"""
with open('agenda1.txt') as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
            
    for row in csv_reader:
        print(f'\t{row[1]} , {row[2]} , {row[3]} , {row[4]} , {row[5]} ')
"""
def openSavedAgendas():
    with open('saved_agendas.txt') as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')

##this did work but now give 'list index out of range
##error, I think its to do with the delimiters
        for row in csv_reader:
            print(f'\t{row[1]}')


openSavedAgendas()

"""
agenda5 = Agenda("Test 5", "08:00:00")
agenda5.addAgendaItem("Intro", "60")
agenda5.printAgenda()
agenda5.saveAgenda(agenda5.name, agenda5.startingTime, agenda5.agendaList)

agenda6 = Agenda("Test 6", "10:00:00")
agenda6.addAgendaItem('introduction', "20")
agenda6.addAgendaItem('topic 1', '60')
agenda6.printAgenda()
agenda6.saveAgenda(agenda6.name, agenda6.startingTime, agenda6.agendaList)
"""