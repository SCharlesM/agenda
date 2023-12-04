"""
practice using the date, time and datetime fuctions
"""

from datetime import datetime, time, timedelta, date

# define a variable with the current datetime object
x = datetime.now()
print(x)

#string format the object, show just the day name (%A)
print("Today is: " + x.strftime("%A"))

#string format. %H = 24 hour clock, %M = minute as a decimal number
print("The current time is: " + x.strftime("%H" + " %M"))

# parse datetime string and add minutes to datetime
time = "10:30"
#convert string to time object
timeObject = datetime.strptime(time, "%H:%M")

print(time)
print(timeObject)

#add 29 minutes to time object
newTime = timeObject + timedelta(minutes=29)
print(newTime)

#strip out just the time, not the date
only_t = newTime.time()
print(only_t)

anotherTime = newTime + timedelta(minutes=29)
print(anotherTime.time())

#function to take a string input start_time and duration to
# make an end_time and print out
def addDuration(start_time, duration):

    #convert string to a time object and format
    time_object_start = datetime.strptime(start_time, "%H:%M")

    #add the duration 
    time_object_end = time_object_start + timedelta(minutes = 60)

    print(time_object_start.time())
    print(time_object_end.time())

addDuration("11:00", 60)
