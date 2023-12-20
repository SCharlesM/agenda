"""
agenda test file, to test the agenda program
or should this be called agenda main or something
"""

from agenda import Agenda

agenda1 = Agenda("Test Agenda", "10:00:00")

agenda1.addAgendaItem("Introduction", "15")
agenda1.addAgendaItem("Topic 1", "45")
agenda1.addAgendaItem("Topic 2", "60")
agenda1.addAgendaItem("Break", "15")
agenda1.addAgendaItem("Topic 3", "25")
agenda1.addAgendaItem("Topic 4", "45")
agenda1.addAgendaItem("Lunch", "60")
agenda1.addAgendaItem("Topic 5", "120")
agenda1.addAgendaItem("Topic 6", "60")
agenda1.addAgendaItem("Closing", "15")
agenda1.printAgenda()

#change the 1st entry to "Intro"
#agenda1.editAgendaItemTitle(1, "Intro")
#agenda1.printAgenda()

#change the 1st entry duration to 60
#agenda1.editAgendaItemDuration(1, "60")
#agenda1.printAgenda()

agenda1.exportToExcel("Agenda1.xlsx")

