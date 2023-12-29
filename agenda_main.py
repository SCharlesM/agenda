"""
The main function for the agenda program, taking an input from the command line
"""
from agenda import Agenda

print("\nWelcome to Agenda, here are the following options")

while True :

    print("\n1. Start a new Agenda Project" +
          "\n2. Add an Agenda Item" +
          "\n3. Edit an Agenda Item Title" +
          "\n4. Edit an Agenda Iten Duration" +
          "\n5. Load a Agenda list from Excel" +
          "\n6. Save an Agenda Project" +
          "\n7. Open a saved Agenda Project"
          "\nx. Close program")

    choice = input("\nType 1, 2, 3, 4, 5, 6, 7 or x: ")
    
    if choice == "1":
        print("You chose 1, Start a new agenda project")
        title = input("What is the new Agenda Title?")
        start = input("What is the new Agenda start-time? **for now this has to be xx:xx:xx**")
        agenda1 = Agenda(title, start)
        agenda1.printAgenda()

    elif choice == "2":
        print("You chose 2, add an Agenda Item")
        title = input("What is the title of the agenda item?")
        duration = input("what is the duration of the agenda item?")
        agenda1.addAgendaItem(title, duration)
        agenda1.printAgenda()

    elif choice == "3":
        print("You chose 3, edit an Agenda Item Title")
        index = str(input("Type the index of the item you would like to edit"))
        title = input("What is the new title?")
        agenda1.editAgendaItemTitle(index, title)
        agenda1.printAgenda()

    elif choice == "4":
        print("You chose 4, edit an Agenda Item duration")
        index = input("Type the index of the item you want to edit")
        new_duration = input("What is the new duration?")
        agenda1.editAgendaItemDuration(index, new_duration)
        agenda1.printAgenda()

    elif choice == "5":
        print("You chose 5. Load an agenda list from Excel")
    
    elif choice == "6":
        print("you chose option 6. Save an Agenda Project")
        agenda1.saveToCSV()
    
    elif choice == "7":
        print("You chose option 7. Open a saved Agenda Project")
        print("for now we are just opening 1 saved agenda as a test")
        agenda1 = Agenda()
        #agenda1.openCSV()
        #agenda1.printAgenda()
    
    elif choice == "x":
        print("\nYou chose x, Agenda is closing")
        break
    else :
        print("\nThat input is not recognised, try again!")

    