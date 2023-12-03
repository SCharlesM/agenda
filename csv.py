import csv

#open the file in 'write' mode
f = open('Steve\Documents\Computers\Coding\python\agenda\csv', 'w')

#create the csv writer
writer = csv.writer(f)

#write a row to the csv file
writer.writerow(row)

#close the file
f.close()
