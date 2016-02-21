#!/usr/bin/python

import sys
import csv
import string
import zipfile
#import operator
import pprint
from subprocess import call

# puts field data to lower case
def lower_getter(field):
  def _getter(obj):
      return obj[field].lower()
  return _getter

# used like this
#list_of_dicts.sort(key=lower_getter(key_field)) 

if len(sys.argv) > 1:
  filename = sys.argv[1]
  print ("file used: " + str(filename))
  if ".zip" in filename:
    print ("filename contains .zip, extracting "+str(filename)+" ...\n")
    with zipfile.ZipFile(filename, "r") as z:
      z.extractall()
      filename = string.replace(filename, '.zip', '.csv')
else:
  print "please provide zip or csv file!"
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

# guess csv files dialect?
with open(filename, 'rb') as csvfile:
  dialect = csv.Sniffer().sniff(csvfile.read(1024))
  #print dialect.quotechar
  #print dialect.delimiter
  #print dialect.doublequote
  # what's this?? why needed?
  csvfile.seek(0)

  reader = csv.DictReader(csvfile, dialect=dialect)
  #print reader.fieldnames 

  # how does data actually look like?
  #for row in reader:
  #  pp = pprint.PrettyPrinter(indent=4, depth=1)
  #  pp.pprint(row)
  

  # sort readers data by email field, 2 many options!!!
  # -> using itemgetter -> call like this: sorted(.... key=sortkey)
  # sortkey = operator.itemgetter('E-Mail Adresse') 
  # -> using special function lower_getter, defined on top
  #sortedlist = sorted(reader, key=lower_getter('E-Mail Adresse'), reverse=False)
  # -> using lambda, why te hell was itemgetter necessary??
  sortedlist = sorted(reader, key=lambda foo: (foo['E-Mail Adresse'].lower()), reverse=False)
  # debug output sortedlist
  #for row in sortedlist:
    #print string.lower(row['E-Mail Adresse']), row['First Name'], row['Last Name']
  #  print row['E-Mail Adresse'], row['First Name'], row['Last Name']

  # write new csv file 
  filename2 = string.replace(filename, '.csv', '_slim.csv')
  with open(filename2, 'wb') as csvfile2:
    writer = csv.DictWriter(csvfile2, dialect=dialect, fieldnames=['E-Mail', 'First Name', 'Last Name']) 
    writer.writeheader()
    for row in sortedlist:
      #print row['E-Mail Adresse'], row['First Name'], row['Last Name']
      writer.writerow({'E-Mail': row['E-Mail Adresse'], 'First Name': row['First Name'], 'Last Name': row['Last Name']})

print "file written: " + str(filename2)
print "starting libreoffice calc... "
call(['/Applications/LibreOffice.app/Contents/MacOS/soffice', '-n', '/Users/jojo/Documents/privat/mr/maillinglist_print_template.ots'])

