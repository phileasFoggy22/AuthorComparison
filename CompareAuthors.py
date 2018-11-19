# -*- coding: utf-8 -*-
"""
Created on Tue Oct 17 16:07:48 2017

@author: w1l19

Designed for use with Python Version 3.6.2 

This is a program designed to take a range of files from seperate authors and to construct 
nGram lists for letters and for words fom each file in order to establish which author wrote 
a sample file. The sample file will be checked against each set of files and will include the 
comparison of average nGrams and standard deviation from the sample file. 

This is extracted to a Spreadsheet and the program will make an educated guess
(based on results of the above) as to which author wrote the sample file.
"""
import re, string, time, math, os
#re and string needed to format text, time to measure performance, math for standard deviation, pandas for datafiles and extraction to a spreadsheets, os for error handling

import packaging.version
import pandas
import pandas.io.formats.excel

def get_format_module():
    version = packaging.version.parse(pandas.__version__)
    if version < packaging.version.parse('0.18'):
        return pandas.core.format
    elif version < packaging.version.parse('0.20'):
        return pandas.formats.format
    else:
        return pandas.io.formats.excel

SampleTextFiles ={"Doyle": ['A1.txt','A2.txt','A3.txt','A4.txt','A5.txt','A6.txt'],
                  "Kipling": ['B1.txt','B2.txt','B3.txt','B4.txt','B5.txt','B6.txt']}
SampleFile = ['C1.txt',]
"""Filenames for easy reference - Added into a dictinary so that if more authors/ files are added, this is all that needs changing"""
ngramSizes = [3,7,20]#Change these for different lengths
writer = pandas.ExcelWriter('AuthorResults.xlsx', engine='xlsxwriter') #Open the excel method to write to a spreadsheet
#pandas.formats.format.header_style = None # This prevents the program giving a default header style so that this can be modified later in the program

ConstructednGramWords = {}
ConstructednGramLetters = {}
"""In the second iteration of development, it was found that the text documents needed to be added to dictionaries so when they are repeatedly referenced, it takes less time"""

def ExtractAllFiles():
    """Extract all files into variables so that they do not have to be constantly opened from the text files"""
    
    ExtractedTexts = {} # Dictionary for all text files
    for item in SampleTextFiles.items():
        TextPerFile = {} # For all text files per Author
        for item2 in item[1]:
            try:
                f = open(item2,'r',encoding="utf8") # Open the file with UTF8 encoding, read the contents and release the file
                sample = f.read()  
                f.close()
                TextPerFile.update({item2:sample}) # Update the dictionary for the Author of the opened file
            except IOError:
                print ("Could not read file:", item2) # If the file is unreadable, throw up error
                os.sys.exit()
        ExtractedTexts.update({item[0]:TextPerFile}) #Update the dictionary for all of the files per author
    TextPerFile = {}
    for item in SampleFile: # Do the same for the sample file(s)
        try:
            f = open(item,'r',encoding="utf8")
            sample = f.read()
            f.close()
            TextPerFile.update({item:sample})
        except IOError:
            print ("Could not read file:", item2) # If the file is unreadable, throw up error
            os.sys.exit()
    ExtractedTexts.update({"Sample":TextPerFile}) # Add the sample file onto the end of the dictionary with the Author: Sample
    return (ExtractedTexts)

def CreatenGrams(ExtractedTexts):
    """This creates all of the nGrams for words and letters in one go so that they can be referenced repeatedly"""
    for key, value  in ExtractedTexts.items():
        ConstructednGramWordsperFile = {}
        ConstructednGramWordsperLetters = {} # Dictionaries for each set of nGrams
        for filename, text in value.items():
            nGramList = FormatnGram(ngramSizes, text, True) #Format the text as appropraite
            nGramSorted = WordFrequencynGram(nGramList, True) #Count the frequency
            ConstructednGramWordsperFile.update({filename:nGramSorted}) #Add to the dictionary
            nGramList = FormatnGram(ngramSizes, text, False)
            nGramSorted = WordFrequencynGram(nGramList, False)
            ConstructednGramWordsperLetters.update({filename:nGramSorted}) #Repeat for letters
        ConstructednGramWords.update({key:ConstructednGramWordsperFile})
        ConstructednGramLetters.update({key:ConstructednGramWordsperLetters}) #Add per Author
        
def WordFrequencynGram(Text, Words):                                                
    """Returns frequency of each word given a list of words."""
    frequency = {}
    ngramList = {}
    if(Words == True):
        for i in range(len(Text)-(ngramSizes[0]-1)):                            #looks through all of the words minus the last word ngram to avoid an out of bounds exception
            ngram = ""
            for ii in range(ngramSizes[0]):                                     #adds the words per nGram - nested loop used to allow for different sizes of nGram
                ngram += (Text[i+ii]+' ')                                       #adds a space for legibility           
            ngramList[i] = ngram
    if(Words == False):
        for i in range(len(Text)-(ngramSizes[1]-1)):                            #looks through all of the words minus the last letter ngram to avoid an out of bounds exception
            ngram = ""                                                          #adds a space for legibility  
            for ii in range(ngramSizes[1]):                                             
                ngram += (Text[i+ii])      
            ngramList[i] = ngram
    for w in ngramList:                                                         #allocates the string to a list                                                     
        frequency[ngramList[w]] = frequency.get(ngramList[w], 0) + 1            #Counts every instance of each word
        
    sred = sorted(frequency.items(), key=lambda value: value[1], reverse = True) #Sort the frequency descending
    return(sred)

def FormatnGram(ngramSizes, text, Words):
    """This formats the text files for words or letters ready for nGrams to be constructed"""
    if(Words == True):
        sp=string.punctuation                                                   # Get punctuation symbols
        RemovePunc = re.compile('[%s]' % re.escape(sp))                         # Compile all
        text = RemovePunc.sub('',text )
        text = text.lower()                                                     # Remove Case
        words = text.split()                                                    # split into words
    else:
        words = re.sub(re.compile(r'\s+'), '', text)                            # split into characters
    return words





    

def extractResults(resultsWords, resultsLetters,Sampleresults) :
    """Used to extract data into a spreadsheet using DataFrames because otherwise printed nGrams are very messy"""
    StartCol = 0
    StartRow = 0 # Variables for spacing within the spreadsheet
    Variance = 0
    Formatting = []

    """The following calculates the total standard deviation from the sample file for Words (per Author)"""
    results = {}
    for Author, value in resultsWords.items():
        totalDeviation = 0
        for file, item in value.items():
            for key in item:
                if(item["1. Author"] == Author):
                    Variance = item["St Dev from C"]
            totalDeviation+=Variance
        results.update({Author:totalDeviation})
    
    GuessedAuthor = min(results, key=results.get)
    ResultStatement1 = ("According to word nGrams, the Author of the Sample file is: "+GuessedAuthor +" with an nGram Cumulative Deviation from the Sample File of: "+ str(int(results[GuessedAuthor])))
    print(ResultStatement1)
    
    """The following calculates the total standard deviation from the sample file for Letters (per Author)"""
    results = {} 
    for Author, value in resultsLetters.items():
        totalDeviation = 0
        for file, item in value.items():
            for key in item:
                if(item["1. Author"] == Author):
                    Variance = item["St Dev from C"]
            totalDeviation+=Variance
        results.update({Author:totalDeviation})
    GuessedAuthor = min(results, key=results.get)
    ResultStatement2 = ("According to letter nGrams, the Author of the Sample file is: "+GuessedAuthor +" with an nGram Cumulative Deviation from the Sample File of: "+ str(int(results[GuessedAuthor])))
    print(ResultStatement2)
           
      
    """This exports the results of the text files (A1-B6) compared to the sample file to a spreadsheet using DataFrames for clarity"""
    for Author, value in resultsLetters.items():
        df = pandas.DataFrame(value)     
        df.to_excel(writer, sheet_name='Sheet1', startcol=StartCol, startrow=StartRow)
        StartCol= (StartCol+2)+(len(value))
        
    """Used to space out results in the spreadsheet"""
    for AuthorKey in SampleTextFiles:
        if(StartRow<len(SampleTextFiles[AuthorKey])):
            StartRow = len(SampleTextFiles[AuthorKey])
            
    StartRow += 9
    StartCol = 0 
    StartRow2 = 0  

    Formatting.append(StartRow)

    """As above but for Words"""
    for Author, value in resultsWords.items():
        df = pandas.DataFrame(value)      
        df.to_excel(writer, sheet_name='Sheet1', startcol=StartCol, startrow=StartRow)
        StartCol= (StartCol+2)+(len(value))
    
    """Used to space out results in the spreadsheet"""    
    for AuthorKey in SampleTextFiles:
        if(StartRow2<len(SampleTextFiles[AuthorKey])):
            StartRow2 += len(SampleTextFiles[AuthorKey])
            
    StartRow += 9+StartRow2
    StartCol = 0
    Formatting.append(StartRow)       
     
    """As above but for the same results from the sample file and splt out per Author"""   
    for Author,value in Sampleresults.items():
        for key, items in value.items():
            df = pandas.DataFrame(items)
            df.to_excel(writer, sheet_name='Sheet1', startcol=StartCol, startrow=StartRow)
            StartCol= (StartCol+2)+(len(items))       
            
        
        
        
    """Ths formats the spreadsheet"""
    workbook  = writer.book
    wrap_format = workbook.add_format({'text_wrap': True})
    worksheet = writer.sheets['Sheet1']
    worksheet.set_column('A:ZZ', 12, wrap_format)
    wrap_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'valign': 'top',
    'fg_color': '#f0f8ff',
    'border': 1})
    worksheet.set_row(0, None, wrap_format)
    worksheet.set_row(Formatting[0], None, wrap_format)
    worksheet.set_row(Formatting[1], None, wrap_format)
    
    """The following tries to guess which author wrote the sampel file"""
    SampleGuessing = {}
    SampleSubGuessing = {} # Two dictionaries used to store results per Author
    for Author,value in Sampleresults.items():
        SampleSubGuessing = {}
        for key, items in value.items(): 
            GuessingTotal = 0
            for name,records in items.items():
                GuessingTotal+=records["Total Average"] # This is the total average instances of an nGram from the Sample file occurring in the Author's Text Files
            SampleSubGuessing.update({key:round(GuessingTotal, 2)}) # Rounds to 2dp
        SampleGuessing.update({Author:SampleSubGuessing}) 

    """This reports the results of the above check and reports the guess to the user"""
    for key, value in SampleGuessing.items():
        GuessedAuthor = min(value, key=results.get)
        ResultStatement3 = ("According to word nGrams, the Author of the Sample file is: "+GuessedAuthor +" with an nGram Cumulative Deviation from the Sample File of: "+ str(value[GuessedAuthor]))
        print(ResultStatement3)
        
    writer.save()                                                               #Saves and releases the Excel file



def Averages(Author,value,File, s, LetWor):
    """Once the most frequent nGrams have been established, their frequency needs to be recorded across other wors by the same author"""
    totalList = {}  
    if(Author in SampleTextFiles):
        for key,value1 in SampleTextFiles.items():                             #For each Author                        
            if(Author == key):
                for nGramOriginal in File: 
                    sublist = {}                                               
                    quantityTotalnGram = 0
                    variancenGram = 0
                    SampleValue = 0
                    sublist["1. Author"] = Author                                 #First RESULTS value is the Author
                    sublist["2. Text File"] = value                               #RESULTS value for text file being analysed
                    sublist["3. nGram"] = nGramOriginal[0]                        #RESULTS nGram being calculated (if words)
                    if LetWor == True:
                        for SubKey, SubValue  in ConstructednGramWords.items():   #Words checked first
                            if("Sample" ==SubKey):                                #Variance calculated from sample file, so quantity of nGrams in sample file is retrieved first
                                for filenameCheck, item in SubValue.items():
                                    quantity = 0
                                    for nGramCheck in item:
                                        if(nGramCheck[0]==nGramOriginal[0]):
                                            SampleValue = nGramCheck[1]           #C1.txt : quantity of nGrams in Sample file
                                    sublist[filenameCheck] = SampleValue        
                            if(Author ==SubKey):                                  #Once the quantity of nGrams in the Sample file has been returned, do the same for each file for this Author
                                for filenameCheck, item in SubValue.items():
                                    quantity = 0
                                    for nGramCheck in item:
                                        if(nGramCheck[0]==nGramOriginal[0]):
                                            quantity = nGramCheck[1]              #Set the quantity for this nGram 
                                            variancenGram += (SampleValue-quantity)**2  #Start counting the standard deviation from the Sample File
                                            quantityTotalnGram += nGramCheck[1]   
                                    sublist[filenameCheck] = quantity
                    else:
                        for SubKey, SubValue  in ConstructednGramLetters.items(): #As above for letters. This has not been made into a seperate function due to the number of necessary variables that would need to be passed
                            if("Sample" ==SubKey):
                                for filenameCheck, item in SubValue.items():
                                    quantity = 0
                                    for nGramCheck in item:
                                        if(nGramCheck[0]==nGramOriginal[0]):
                                            SampleValue = nGramCheck[1]
                                    sublist[filenameCheck] = SampleValue 
                            if(Author ==SubKey):
                                for filenameCheck, item in SubValue.items():
                                    quantity = 0
                                    for nGramCheck in item:
                                        if(nGramCheck[0]==nGramOriginal[0]):
                                            quantity = nGramCheck[1]
                                            quantityTotalnGram += nGramCheck[1]
                                            variancenGram += (SampleValue-quantity)**2
                                    sublist[filenameCheck] = quantity                    #RESULTS Number of instances of nGram per File for Author
    
                    
                    sublist.update({"Total":quantityTotalnGram})                         #RESULTS Add the total number of occurances for this nGram for all of the Author's files
                    sublist.update({"Total Average":quantityTotalnGram/len(value1) })    #RESULTS Total average of the files above for this Author
                    sublist.update({"St Dev from C":math.sqrt(variancenGram)})           #RESULTS Deviation from the Sample file for this nGram
                    totalList.update({str(key+":"+value+":"+nGramOriginal[0]): sublist}) #RESULTS Unique Key for the dictionary
    return(totalList)  
    
def AverageSample(Author,value,File, s, LetWor):
    """This is a similar process to the Averages above but because it is accessing files for all authors, it made since to split out into a seperate function"""
    """This gets the nGrams from the sample file and sees how many times they occur in the other text files"""
    totalList = {}  
    AuthorDict = {}
    if LetWor == True:       
        for SubKey, SubValue  in ConstructednGramWords.items():
            totalList = {}  
            if(Author !=SubKey):
                if(SubKey not in AuthorDict.keys()):
                    for nGramOriginal in File: 
                        sublist = {}
                        quantityTotalnGram = 0
                        sublist.update({"1. Compared Author" : SubKey})                            #First RESULTS value is the Author being Compared
                        sublist.update({"2. Text File" : value})                                   #RESULTS value for the sample text file name
                        sublist.update({"3. nGram" : nGramOriginal[0]})                            #RESULTS value for the sample text file nGram
                        sublist.update({"4. Quantity" : nGramOriginal[1]})                         #RESULTS value for the quantity of the sample text file nGram                                    
                        for filenameCheck, item in SubValue.items():
                            quantity = 0
                            for nGramCheck in item:
                                if(nGramCheck[0]==nGramOriginal[0]):
                                    quantity = nGramCheck[1]
                                    quantityTotalnGram += nGramCheck[1]
                            sublist[filenameCheck] = quantity                                      #RESULTS value for the nGram compared to each other file
                        sublist.update({"Total":quantityTotalnGram})
                        sublist.update({"Total Average":quantityTotalnGram/len(SubValue) })
                        totalList.update({"Sample Vs "+str(SubKey)+"-'"+nGramOriginal[0]+"'":sublist}) #Totals for the sample file per nGram
                    AuthorDict.update({SubKey:totalList})
    else:
        for SubKey, SubValue  in ConstructednGramLetters.items():                                 #As above but for letters
            totalList = {} 
            if(Author !=SubKey):
                if(SubKey not in AuthorDict.keys()):
                    for nGramOriginal in File: 
                        sublist = {} 
                        quantityTotalnGram = 0
                        sublist["1. Compared Author"] = SubKey                                
                        sublist["2. Text File"] = value                            
                        sublist["3. nGram"] = nGramOriginal[0]
                        sublist["4. Quantity"] = nGramOriginal[1]
                        for filenameCheck, item in SubValue.items():
                            quantity = 0
                            for nGramCheck in item:
                                if(nGramCheck[0]==nGramOriginal[0]):
                                    quantity = nGramCheck[1]
                                    quantityTotalnGram += nGramCheck[1]
                            sublist[filenameCheck] = quantity
                        sublist.update({"Total":quantityTotalnGram})
                        sublist.update({"Total Average":quantityTotalnGram/len(SubValue) })
                        totalList.update({"Sample Vs "+str(SubKey)+"-'"+nGramOriginal[0]+"'":sublist})
    
                    AuthorDict.update({SubKey:totalList}) 
    return(AuthorDict)  
        

   

def main():
    """Runs the main program and asks user for quantities and nGram size"""
#    ngramSizes[0] = 3#int(input("How many words do you want included in the nGram?: ")) SIZE
#    ngramSizes[1] = 10##int(input("How many letters do you want included in the nGram?: ")) CHAR SIZE
#    ngramSizes[2] = 20#int(input("How many nGrams do you want to return?: ")) QUANTITY


    ExtractedTexts = ExtractAllFiles() #Get the text from the files per Author
    CreatenGrams(ExtractedTexts) #Make the nGrams and add them to dictionaries
    
    resultsWords = {} 
    resultsLetters = {}       
    Sampleresults = {}
    #Make dictionaries for all of the results
                   
    for key, value  in ConstructednGramWords.items():
        results = {}
        for filename, text in value.items():
            if(key!="Sample"):
                results.update(Averages(key,filename,text[:ngramSizes[2]],results,True))         #Get the average occurances across all files for this Author
                resultsWords.update({key:results})
                
            else:
                results.update(AverageSample(key,filename,text[:ngramSizes[2]],results,True))    #Do the same for the Sample File
                Sampleresults.update({key+"Words":results})
                

           
    for key, value  in ConstructednGramLetters.items(): #As above with letter nGrams
        results = {}    
        for filename, text in value.items():
            if(key!="Sample"):
                results.update(Averages(key,filename,text[:ngramSizes[2]],results,False))
                resultsLetters.update({key:results})
            else:
                results.update(AverageSample(key,filename,text[:ngramSizes[2]],results,False))
                
                Sampleresults.update({key+"Letters":results})
    
    
    extractResults(resultsWords, resultsLetters,Sampleresults)                                   #Stick the results into a spreadsheet


start_time = time.time()
main()
print("---Time taken to run:  %s seconds ---" % (time.time() - start_time))     #Time the process to measure efficiency
