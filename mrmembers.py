#!/usr/bin/python

import sys
import csv
import string

if len(sys.argv) > 1:
  filename = sys.argv[1]
  print ("csv file used: " + str(filename) + "\n")
else:
  print "please provide csv file!"
  raise sys.exit()

# DEBUG dump 1st 200 bytes
#print "### DEBUG: 1st 200 bytes... ###"
#print repr(open('members_export_f664863d54.csv', 'rb').read(200))
#print ""

# cleanup csv file -> replace NUL characters 
fi = open(filename, 'rb')
data = fi.read()
fi.close()
fo = open(filename, 'wb')
fo.write(data.replace('\x00', ''))
fo.close()

with open(filename, 'rb') as csvfile:
  dialect = csv.Sniffer().sniff(csvfile.read(1024))
  #print dialect.quotechar
  #print dialect.delimiter
  #print dialect.doublequote
  # what's this?? why needed?
  csvfile.seek(0)

  reader = csv.DictReader(csvfile, dialect=dialect)
  #headers = reader.fieldnames
  #print headers 
  filename2 = string.replace(filename, '.csv', '_slim.csv')

  with open(filename2, 'wb') as csvfile2:
    writer = csv.DictWriter(csvfile2, dialect=dialect, fieldnames=['E-Mail', 'First Name', 'Last Name']) 
    writer.writeheader()
    for row in reader:
      #print(row['E-Mail Adresse'], row['First Name'], row['Last Name'],)
      writer.writerow({'E-Mail': row['E-Mail Adresse'], 'First Name': row['First Name'], 'Last Name': row['Last Name']})
