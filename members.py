#!/usr/bin/python

import sys
import csv
import string
import zipfile
import pprint
import operator
from subprocess import call


# returns lower case version of fields data
def lower_getter(field):
  def _getter(obj):
      return obj[field].lower()
  return _getter

# used like this
#list_of_dicts.sort(key=lower_getter(key_field))

print ("")
if len(sys.argv) > 1:
  filename = sys.argv[1]
  filenamepath=filename.rsplit('/',1)[0]
  filenamefile=filename.rsplit('/',1)[1]
  if ".zip" in filename:
    print("filename contains .zip, extracting "+str(filenamefile)+" to "+str(filenamepath)+"\n")
    with zipfile.ZipFile(filename, "r") as z:
      z.extractall(filenamepath)
      filenamefile = string.replace(filenamefile, '.zip', '.csv')
      filenamefile = "subscribed_"+filenamefile
      filename=filenamepath+"/"+filenamefile
  else:
    print("filename does not contain .zip, assuming "+str(filenamefile)+" contains csv data"+"\n")
    filename=filenamepath+"/"+filenamefile
else:
  print("please provide zip or csv file!")
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

# guess csv files dialect
with open(filename, 'rb') as csvfile:
  dialect = csv.Sniffer().sniff(csvfile.read(1024))
  #print dialect.quotechar
  #print dialect.delimiter
  #print dialect.doublequote
  # seek back to position 0
  csvfile.seek(0)

  reader = csv.DictReader(csvfile, dialect=dialect)
  #print reader.fieldnames 

  # how does data actually look like? comment out whole block for debugging
  #for row in reader:
  #  pp = pprint.PrettyPrinter(indent=4, depth=1)
  #  # each row contains a dict:
  #  pp.pprint(row)
  #  # only show the dicts key email:
  #  pprint.pprint(row["E-Mail Adresse"])
  #  # show it in lower case
  #  pprint.pprint(row["E-Mail Adresse"].lower())

  # 3 different ways of sorting
  # -> using itemgetter -> call like this: sorted(.... key=sortkey)
  #    this version is missing the make it lower case part
  #sortkey = operator.itemgetter('E-Mail Adresse')
  #sortedlist = sorted(reader, key=sortkey, reverse=False)
  # -> lambda is used as a one statement function
  #    instead of creating a named function (def) for the purpose
  #sortedlist = sorted(reader, key=lambda foo: (foo['E-Mail Adresse'].lower()), reverse=False)
  # -> or by using the special function lower_getter, defined on top
  sortedlist = sorted(reader, key=lower_getter('E-Mail Adresse'), reverse=False)

  # debug output sortedlist
  #for row in sortedlist:
  #  print row['E-Mail Adresse'], row['First Name'], row['Last Name']

  # write new csv file 
  filename2 = string.replace(filename, '.csv', '_slim.csv')
  with open(filename2, 'wb') as csvfile2:
    writer = csv.DictWriter(csvfile2, dialect=dialect, fieldnames=['E-Mail', 'First Name', 'Last Name']) 
    writer.writeheader()
    for row in sortedlist:
      #print row['E-Mail Adresse'], row['First Name'], row['Last Name']
      writer.writerow({'E-Mail': row['E-Mail Adresse'], 'First Name': row['First Name'], 'Last Name': row['Last Name']})

print("file written: " + str(filename2))
print("")
print("starting libreoffice calc... ")
call(['/Applications/LibreOffice.app/Contents/MacOS/soffice', '-n', '/Users/jojo/Documents/privat/mr/maillinglist_print_template.ods'])
#FIXME how to run second LibreOffice instance with NEW document? how does call work?
# import to NEW document, copy/paste to print_template which should also be already open
#call(['/Applications/LibreOffice.app/Contents/MacOS/soffice', '-n'])

