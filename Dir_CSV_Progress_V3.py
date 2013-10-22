"""
This script is used to get directions
using google maps API and parsing the resulting json object

Functions 
StrCheck            - PreProcess directions to remove " " and inserts "+" 
GetJson             - Returns the Json Object containing directions
ParseSummary        - Returns the route summary including start, end, and distance
ParseSteps          - Returns the steps to get from one place to another
"""
# Import required modules
import json
import urllib
import pprint
import re
import time
import csv
import sys
import webbrowser
from docx import *

# docx required stuff - maybe make into function along with stuff at bottom
# Default set of relationshipships - the minimum components of a document
relationships = relationshiplist()

# Make a new document tree - this is the main part of a Word document
document = newdocument()

# This xpath location is where most interesting content lives
body = document.xpath('/w:document/w:body', namespaces=nsprefixes)[0]

# Define Functions
def StrCheck(rawstr):
     """
     Name           : StrCheck
     Inputs         : rawstr
     Return         : CorrectStr
     Description    : This function checks input string for spaces
                         and replaces them with + for HTML
     """
     # Remove Non Alpha Numeric Characters
     # rawstr = re.sub('[^0-9a-zA-Z]+', '', rawstr)
     rawstr = re.sub(r'([^\s\w]|_)+', '', rawstr)

     # Replace White Spaces with +
     CorrectStr = rawstr.replace(" ", "+")

     return CorrectStr

def GetJson(origin, destination, sensor = "false"):
     """
     Name           : GetJson
     Inputs         : Origin, Destination, sensor(Optional)
     Return         : JsonObj
     Description    : This function returns the raw json object from googlemaps
                         Orgin and Destinations need to be valid adresses
                         Sensor is an optional input that tells google the 
                         device making the query has a gps
     """
     # Define URL for Directions
     DIR_BASE_URL = 'http://maps.googleapis.com/maps/api/directions/json'

     # PreProcess Directions
     origin      = StrCheck(origin)
     destination = StrCheck(destination)

     # Define Dictionary for URL Encoding and create URL string
     dir_args = {'origin': origin,'destination': destination,'sensor': sensor}
     url      = DIR_BASE_URL + '?' + urllib.urlencode(dir_args)
     #print "URL: ", url

     # Get json result from direction query and return
     JsonObj = json.load(urllib.urlopen(url))
     return JsonObj

def ParseSummary(JsonObj):
     """
     Name           : ParseSummary
     Inputs         : JsonObj
     Return         : Summary
     Description    : This function returns the trip summary including
                         origin, destination, and distance

     """
     for routes in JsonObj['routes']:
          for legs in routes['legs']:
               start  = "From: " + legs['start_address'] + " "
               finish = "To: " + legs['end_address'] + " "
               dist   = "Distance: " + legs['distance']['text']
               Summary = start + finish + dist

               return Summary

               # print "From: ", legs['start_address'] 
               # print "To: ", legs['end_address'] 
               # print "Distance: ", legs['distance']['text']

def ParseSteps(JsonObj, distance = False):
     """
     Name           : ParseSteps
     Inputs         : JsonObj
     Return         : FullDir
     Description    : This function returns strings with the steps inside them
     """
     FullDir = []
     for routes in JsonObj['routes']:
          for legs in routes['legs']:
               for steps in legs['steps']:
                    # html instructions can contain unicode 
                    # that cannot be interpreted by python with out encoding
                    directions = steps['html_instructions'].encode('utf8')

                    # Remove <b> HTML Instructions from strings
                    directions = re.sub(r"<([^>]*)>", "", directions)

                    if distance:
                         # Also Include Distances
                         distance   = steps['distance']['text']
                         directions = directions + " " + distance

                    if "Destination" in directions:
                         MatchObj = re.finditer(r"Destination", directions)
                         for index in MatchObj:
                              directions, destination = directions[:index.start(0)], directions[index.start(0):]
                              FullDir.append(directions)
                              FullDir.append(destination)
                    else: 
                         FullDir.append(directions)

     return FullDir


# Test Functions 
def PrintDirections(start):
     """
     Name           : PrintDirections
     Inputs         : start
     Return         : N\A
     Description    : This function prints directions to file
     """

     # Print logo
     print """______                  _              ______            _             
| ___ \                (_)             | ___ \          | |            
| |_/ /   _ _ __  _ __  _ _ __   __ _  | |_/ /___  _   _| |_ ___  ___  
|    / | | | '_ \| '_ \| | '_ \ / _` | |    // _ \| | | | __/ _ \/ __| 
| |\ \ |_| | | | | | | | | | | | (_| | | |\ \ (_) | |_| | ||  __/\__ \ 
\_| \_\__,_|_| |_|_| |_|_|_| |_|\__, | \_| \_\___/ \__,_|\__\___||___/ 
                 ______      _   __/ |   _                             
                 | ___ \    | | |___/   | |                            
                 | |_/ /___ | |__   ___ | |_                           
                 |    // _ \| '_ \ / _ \| __|                          
                 | |\ \ (_) | |_) | (_) | |_                           
                 \_| \_\___/|_.__/ \___/ \__|                          
                                                                       
                                                                       """

     # Open output file
     fo = open("Running Routes.txt", "w")
     output_file = fo.name

     # Set starting point
     origin = start

     # Prompt - comment when using sublime since it throws a shit fit when asked for input
#     raw_input()
     print "Starting at", origin

     # Open input file
     with open("Address File.csv", "rb") as f:
         
         fi = open("Address File.csv")
         lines = len(fi.readlines())
         fi.close()

         # File and program info
         print "Calculating", lines, "running routes"
         print "Saving to", output_file, "\n"

         # CSV reader object from file
         reader = csv.reader(f)
             
         # setup toolbar
         toolbar_width = 50

         # Counter for percentage complete
         i = 0

         # Table for docx
         tbl_rows = [['Incident', 'Address','Box','Map','Directions']]
         
         # For every line in the CSV file, find the address and get directions
         for row in reader:         

              # Increment counter
              i += 1

              # Manipulate original string - add MD, remove spaces and dashes
              destination    = row[1] + ", MD\n"
              destination    = destination.replace("On ","")
              destination    = destination.replace(" --","")
              destination    = destination.replace(" -","")
              result         = GetJson(origin, destination)
              Directions     = ParseSteps(result)
              fo.write(destination + "\n")              
              
              # Create docx table
              row[1] = destination
              row.append(Directions)
              tbl_rows.append(row)

              # Print directions to file
              for steps in Directions:
                   steps = steps.replace("Head southeast on","Turn right onto")
                   steps = steps.replace("Head northwest on","Turn left onto")
                   fo.write(steps + "\n")
              fo.write("\n\n")
               
              # Percent complete
              percent = i*(100/float(lines))

              percent_tick = (float(1)/toolbar_width)*100
              tick = int(percent/percent_tick)
              
              sys.stdout.write("\r")
              sys.stdout.write("[%-50s] %d%%" % ('='*tick, percent))
              sys.stdout.flush()

              # Sleep for half a second to let API catch up     
              time.sleep(0.5)

         # Append docx table to document body
         body.append(table(tbl_rows))

     # Close output file
     fo.close()

     # Prompt - comment when using sublime since it throws a shit fit when asked for input 
#     raw_input("\n\nProcess complete. Press any key to exit...")
     webbrowser.open(output_file)

     

PrintDirections("10155 Old Columbia Rd, Columbia, MD 21046")

# docx required stuff - maybe make into function along with stuff at bottom
# Create our properties, contenttypes, and other support files
title    = 'Running Routes'
subject  = 'Written directions for various incident addresses'
creator  = ''
keywords = []

coreprops = coreproperties(title=title, subject=subject, creator=creator,
                         keywords=keywords)
appprops = appproperties()
contenttypes = contenttypes()
websettings = websettings()
wordrelationships = wordrelationships(relationships)

# Save our document
savedocx(document, coreprops, appprops, contenttypes, websettings,
       wordrelationships, 'Routes Test.docx')
