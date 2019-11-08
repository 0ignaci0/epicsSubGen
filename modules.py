#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Thu Nov  7 08:56:48 2019

@author: rguerra
"""

from tabulate import tabulate


def write_file( fn, wl, writeMode = 'a'  ):
    f = open(fn,writeMode)
    f.writelines( wl )
    f.close()
    return
def dfsearch_insert_cols(df, rowNum, colLabl, search, insertColLbl, insertLocation, val=" "):
    """ arg list: dataframe, row number and column label to search, search value, label of columns to be inserted, insert location
     and optional value to be inserted into columns """
    if df.at[rowNum,colLabl] ==  search: 
        for k in range( insertLocation, len(insertColLbl)-1 ):
            df.insert(k, insertColLbl[k], val )
        return True
    else:
        return False
    
def format_dataframe(df):
    df = df.drop(['Base Name','Type'], axis=1) # drop columns not used in substitution file
    for ii in range( 1, len( df.transpose() ) ):    
        if ii == (len(df.transpose())-1):
            df.iloc[0,ii] = "\"" + str(df.iloc[0,ii]) + "\"}"
            continue
        df.iloc[0,ii] = "\"" + str( df.iloc[0,ii] ) + "\"," 
    return df

def write_sub_file_first(subF,sheet,dbfn,curr):
    write_file(subF,"# "+sheet+"\n")
    write_file(subF,"file \"$(TOP)/db/" + dbfn +"\"\n{\npattern\n")                               
    write_file(subF,tabulate(curr, headers="keys", tablefmt="plain", showindex=False)+"\n")
    return
def write_sub_file_last(subF,curr):
    write_file(subF,tabulate(curr, tablefmt="plain", showindex=False)+"\n}\n")
    return
def write_sub_file_mid(subF,curr):
    write_file(subF, tabulate(curr, tablefmt="plain", showindex=False)+"\n" )                
    return

def search_and_sort(df, col):
    distinctVal = 1 # number of distinct values (strings, numbers, et cetera) present in specified column of dataframe ; starts at value '1'
    changeIndex = [] # list of tuples for position of value transition in column ; [[position_index,value]]
    sortedDf = df.sort_values(by=col) # takes passed data frame and sorts by specified column into local df instance
    sortedDf = sortedDf.reset_index(drop=True) # resets this local df instance's index, dropping old index
    
    for i in range(len(sortedDf)): # for all rows of sorted dataframe
        if i == (len(sortedDf)-1): # if at last row of data frame
            break                  # break loop to avoid out of bound index
        if sortedDf.at[i, col] != sortedDf.at[i+1, col]: # while current row's value at specified column is equal to next row's value in specified column
            distinctVal += 1 # when new value found, increment 
            changeIndex.append([i,sortedDf.at[i,col]]) # push location and new Type value to stack
    
    if distinctVal == 1: # if column homogenous
        return distinctVal
    if distinctVal > 1:
        return [sortedDf, changeIndex]
    


#if sheetNames[i] == 'Display': # check for speical case for 'Display' sheet, mixed data types.
#        sortedDisp = dataframeDict[sheetNames[i]].sort_values(by='Type')
#        sortedDisp = sortedDisp.reset_index(drop=True)
#        j = 0
#        while sortedDisp.at[j,'Type'] == 'Bool':
#            if j == 0:
#                firstFlag = True
#            if j == len(sortedDisp-1):
#                lastFlag = True
#            
#            if j < len(sortedDisp):
#                j += 1
#        for j in range(len(sortedDisp)):
#            
#        while j < len(dataframeDict[sheetNames[i]]):
#            currEntry = pd.DataFrame(dataframeDict[sheetNames[i]].loc[j,:]).transpose()
#            if currEntry.at[j,'Type'] == 'Int'
#        # first, for loop looking for 'Int' type, writing those to a section with ai record type, close curly brackets
#        # then, for loop looking for 'Bool' type, writing those to a section with a bi record type, close curly brackets