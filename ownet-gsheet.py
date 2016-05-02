#!/usr/bin/python -u

#
import os, glob, time, gspread, sys, datetime # all normal imports and the gspread module
import gspread
from oauth2client.service_account import ServiceAccountCredentials

from time import localtime, strftime
from pprint import pprint

from pyownet import protocol

creds_json = "/opt/ownet2gsheet/creds.json"
DEBUG=False
TIMEOUT=5

hostname="rpi-loft"
port=4304


#
# *** MUST share the sheet with the email name in the creds.json ***
#
spreadsheet = 'Temperature_Log' #the name of the spreadsheet already created

scope = ['https://spreadsheets.google.com/feeds']

column_headers = ["Date","CH_feed", "CH_return", "Hot_water", "Cold_water", "One", "Two"]

def get_RPI_owtemp(host, port):
    try:
        owproxy = protocol.proxy(host, port, verbose=DEBUG)
    except protocol.ConnError as error:
        print "Error: Unable to open connection to host:port"
        return null
    except protocol.ProtocolError as error:
        print "Protocol error",error
        return null
    
    try:
      owdir =  owproxy.dir(slash=False, bus=False, timeout=TIMEOUT)
      #print owdir
    except protocol.OwnetError as error:
      print "ownet error",error
      return null
    except protocol.ProtocolError as error:
      print "Protocol error getting owdir",error
      return null
    
    sensor_data = {}
    for sensor in owdir:
      try:
        stype = owproxy.read(sensor + '/type', timeout=TIMEOUT).decode()
       
        if stype in [ 'DS18S20' , 'DS18B20' ]:
          data = owproxy.read(sensor+'/temperature', timeout=TIMEOUT)
    
          #print "read: %s %.1f" %  ( sensor[1:], float(data) )
          
          sensor_data[ sensor[1:] ] = float(data)
    
      except protocol.OwnetError as error:
        print "ownet error",error
        next
      except protocol.ProtocolError as error:
        print "Protocol error",error
        next

    return sensor_data


timestamp = strftime("%Y-%m-%d %H:%M:%S", localtime())

print "Starting: ",timestamp

# creds.json 
try:
  credentials = ServiceAccountCredentials.from_json_keyfile_name(creds_json, scope)
except Exception as e:
  print "Error: ",e

if not credentials:
  print "Error: creds not loaded"
  sys.exit()


headers = gspread.httpsession.HTTPSession(headers={'Connection':'Keep-Alive'})

try:
    gc = gspread.authorize(credentials)
except Exception as e:
    print("Error: Login failed",e)
    sys.exit()


#
#open the spreadsheet
try:
    sheet = gc.open(spreadsheet)
except Exception as e:
    print("Error: sheet not found",e)
    sys.exit()

try:
    ws = sheet.get_worksheet(0)
except Exception as e:
    print("Error: worksheet not found",e)
    sys.exit()
    
# format the sheet to the way we want it    
val = ws.acell('A1').value

if val != "Date":
    ws.resize(1,8) # resize to one row high, 8 wide
        
    # add headers
    cell_list = ws.range('r1c1:r1c8')  # needs to be R1C1:R2C7 format not A1:Z8
    
    cell_values = column_headers
    
    for i, val in enumerate(cell_values):   #gives us a tuple of an index and value
        cell_list[i].value = val            #use the index on cell_list and the val from cell_values
    
    ws.update_cells(cell_list)


# normal update - add a row of data
data=get_RPI_owtemp(hostname, port)

print "data:",data

columns_out = []

columns_out.append( timestamp )

for h in column_headers:
    
    if h != 'Date':
      columns_out.append( data[h] )
    
print "added: ", columns_out

ws.append_row( columns_out )


print "Done"

