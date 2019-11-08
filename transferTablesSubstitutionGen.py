#!/bin/env python

import pandas as pd
import openpyxl as xl

if hasattr(__builtins__, 'raw_input'):
    input=raw_input
# prompt user for file name of Workbook
fileName  =   input('Enter file name of XLSX File: ')
# prompt user for BLEPS beamline name
beamLine    =   input('Enter beamline name: ')
# open workbook
book = xl.load_workbook( fileName, data_only=True)

# create list of sheet names
sheetNames = book.sheetnames
# initialize dictionary for storing process variables keyed by sheet name
sheetDict = {}
# row index constant for assigning dictionary index according to process variable attribute (type, base name, pv name, epics ethernet tag, description)
cRo = 1
# initialize dictionary of DataFrames
dataframeDict = {}
#
recordTypes = ['Int','Bool','Iint','Dint','Real']
# list of new column orders for PV data frames
newOrder = ['PV Name','EPICS Ethernet Tag', 'Short Description', 'Base Name', 'Type']
# for binary input: scan, zero name, one name, zero severity, one severity
biOrder  = ['P','N','TAG','SCAN', 'ZNAM','ONAM', 'ZSV', 'OSV','DESC' ]
# for binary output: high, zero name, one name
boOrder  = [ 'P' ,'N' ,'TAG','HIGH' , 'ZNAM','ONAM', 'DESC' ]
# for analog input: scan, precision, engineering units, hihi, hi, lo, lolo, hihi + hi severity, lolo + lo severity
aiOrder  = ['P','N','TAG','SCAN', 'PREC','EGU', 'HIHI', 'HIGH', 'LOW', 'LOLO', 'HHSV', 'HSV', 'LSV', 'LLLSV', 'DESC' ]
columnsList = [ aiOrder, biOrder, aiOrder, aiOrder, aiOrder ]


    
# iterate over each sheet of workbook
for s in range(len(sheetNames)):
    # initialize/re-initialize list of current sheets used process variables
    dictList = []    
    # find the max height of each sheet; i.e. the last row in each sheet
    maxRow = book[sheetNames[s]].max_row
    # column index constant (each process variable has 5 attributes listed previously)
    maxCol = 5
    # set current row index for each page iteration (start at row 2, as row 1 is column titles/process variable attribute title ; rows are 1-indexed )
    currRow = 2
    # iterate over column A for all process variables 
    for row in book[sheetNames[s]].iter_rows(2,maxRow,1,1,True):
        # initialize/re-initalize current dictionary containing the current process variable's attributes
        currDict = {}
        # set current column index for each row/process variable iteration ( start at column 1, 'Type' attribute ; columns are 0-indexed )
        currCol = 1
        # check each process variable's 'Used' field for 'X'
        for cell in row:
            # if 'X' character found in 'Used' field, add row to current dictionary
            if cell == 'X':
                # iterate over used process variable's attributes/columns
                for col in book[ sheetNames[s] ].iter_cols(1,maxCol,currRow,currRow,True):
                    # add current attribute to current dictionary, using attribute name/column title as dictionary key
                    currDict[ book[ sheetNames[0] ][cRo][currCol].value ] = book[ sheetNames[s] ][currRow][currCol].value
                    # increment column index to add next attribute
                    currCol += 1
                # add process variable's attribute-dictionary to list of all process variables
                dictList.append( currDict )
            # increment row index to check if next process variable is used
            currRow += 1
    # dictList now contains current sheet's used process variables ; add to dictionary of all used process variables, keyed by sheet name
    sheetDict[ sheetNames[s] ] = dictList

# remove dictionary keys/pages that have no process variables present
i = 0
while i < len( sheetNames ):
    if not sheetDict[ sheetNames[i] ]:
        del sheetDict[ sheetNames[i] ]
        del sheetNames[i]
    i += 1
        

j = 0
eiFlag = 0
# convert each process variable dictionary to a DataFrame, store in dictionary keyed by sheet name
for i in range(  len(sheetNames) ):
    for i2 in range( len(sheetNames-1) ):
        j = 0
        eiFlag = False
        dataframeDict[ sheetNames[i] ] = pd.DataFrame( sheetDict[ sheetNames[i] ] )
        # reorder columns for writing substitution file
        newColumns = newOrder + ( dataframeDict[ sheetNames[i] ].columns.drop(newOrder).tolist() )
        dataframeDict[ sheetNames[i] ] = dataframeDict[ sheetNames[i] ][newColumns]
        # insert column holding beamline identifier 
        dataframeDict[ sheetNames[i] ].insert(0,'P,', str(beamLine) )
        # rename existing columns to database pattern identifiers 
        dataframeDict[ sheetNames[i] ] = dataframeDict[ sheetNames[i] ].rename(columns={ "PV Name":"N","EPICS Ethernet Tag":"TAG","Short Description":"DESC" } )
        # insert columns for each record type parameters
    
        # -.- # check for special case for 'EPICS_Inputs' sheet, use binary output columns
        if sheetNames[i] == 'EPICS_Inputs':
            #insert binary output record column names
            for jj in range(3,6):
                dataframeDict[ sheetNames[i] ].insert(jj,boOrder[jj], '"",')
                eiFlag = True # special case signal      
        # -.- #                              # -.- #                              # -.- #        
    
        # checks if 
        dataframeDict[ sheetNames[i] ].iloc[:,-1] == 
        result = dataframeDict[ sheetNames[i] ].ne( dataframeDict[ sheetNames[i2] ], axis='columns' )
        # check if dataframe is of homogenous 'Type' identifier; if not continue, if yes, sort out types
        if result.at[i,'Type']: # if another record type is present, raise flag
            #
    
    
    
    #  otherwise, use record columns corresponding to record type
    if not eiFlag:
        for ii in range( len(dataframeDict[ sheetNames[i] ].get( 'Type' ) ) ):
            while j < len(recordTypes)-1 and dataframeDict[ sheetNames[i] ].at[ii,'Type'] != recordTypes[j]: # do nothing
                j += 1
        for k in range( 3, len( columnsList[j] )-1 ):
            dataframeDict[ sheetNames[i] ].insert(k, columnsList[j][k], '"",' )

## prob with above is all contents of a dictionary have to have same column headings.
## 
    
    
# now parse out these data frames into formatted text for substitution file
# add columns to each data frame for respective data ty
    
    


    
                      
