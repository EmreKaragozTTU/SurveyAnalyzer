# -*- coding: utf-8 -*-
"""
Created on Tue Jun  2 11:58:59 2020

@author: Emre-USA
"""

import xlrd 
import sys
import os
import statistics
from scipy.stats import ttest_ind

args=sys.argv[1:]

if (len(args))>2:
    print("Too many arguments")
    print("Progmram is being terminated...")
    exit()

if args:
    print("Pre-Survey file: "+args[0])
    print("Post-Survey file: "+args[1])
else:
    print("Missing file names as arguments")

class Question: #create class for the questions
  QuestionNumber="Qunassigned"
  pVal=0    
  preCount=0
  preMean=0
  preStdDev=0
  postCount=0
  postMean=0
  postStdDev=0
#Questions = [] # a list for questions but not needed

dir_path = os.path.dirname(os.path.realpath(__file__))
finalPathPre=dir_path+'/'+args[0]
finalPathPost=dir_path+'/'+args[1]
  
wb = xlrd.open_workbook(finalPathPre) 
sheet = wb.sheet_by_index(0) 
wb2 = xlrd.open_workbook(finalPathPost) 
sheet2 = wb2.sheet_by_index(0) 

theColumnCount=sheet.ncols
print(theColumnCount)

for x in range(17, theColumnCount):
    tempQ = Question()
    tempQ.QuestionNumber=sheet.cell(0,x).value
    tempQ.preCount=sheet.nrows-3
    tempQ.postCount=sheet2.nrows-3
    
    listPres=[]
    listPosts=[]
    for y in range(3, sheet.nrows):
        #print(sheet.cell(y,x))
        listPres.append(sheet.cell(y,x).value)
    for y in range(3, sheet2.nrows):
        #print(sheet.cell(y,x))
        listPosts.append(sheet2.cell(y,x).value)
    tempQ.preMean=statistics.mean(listPres)
    tempQ.postMean=statistics.mean(listPosts)
    tempQ.preStdDev=statistics.stdev(listPres)
    tempQ.postStdDev=statistics.stdev(listPosts)
    ttest,pval=ttest_ind(listPres,listPosts)
    tempQ.pVal=pval/2
    #Questions.append(tempQ)
    print("pVal of "+tempQ.QuestionNumber+": "+str(tempQ.pVal))
    print(" preCount: "+str(tempQ.preCount))
    print(" preMean: "+str(tempQ.preMean))
    print(" preStdDev: "+str(tempQ.preStdDev))
    print(" postCount: "+str(tempQ.postCount))
    print(" postMean: "+str(tempQ.postMean))
    print(" postStdDev: "+str(tempQ.postStdDev))
    print("")
