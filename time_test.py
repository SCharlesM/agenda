"""
practice using the date, time and datetime fuctions
"""

from datetime import datetime, time, timedelta, date

x = datetime.now()
print(x)

print("Today is: " + x.strftime("%A"))

print("The current time is: " + x.strftime("%H" + " %M"))


# parse datetime string and add minutes to datetime
# time string
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

"""

"""
