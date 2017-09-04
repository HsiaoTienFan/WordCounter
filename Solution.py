# -*- coding: utf-8 -*-
"""
Created on Mon Aug 21 13:08:18 2017

@author: fanat
"""


import glob
import re
from collections import Counter
import pandas as pd
path = 'C:\Users\fanat\Documents\Python\Eigen Technology'
cleaned = []
sentences = []
dict_sentences = {}
table = []
for filename in glob.glob('*.txt'):
    with open (filename, "r") as myfile:
        readIn = myfile.read()
        
    #clean input text to remove new line, take into account of capitalized initials for sentence segmentation
    readIn = re.sub(r'\n', ' ', readIn)
    readIn = re.sub(r"(?<=[A-Z])\.", "_", readIn)
    readIn = re.sub(r"’", "'", readIn)
    
    #read all text into single variable for overall word count    
    sentences = list(readIn.split('.'))
    cleaned += re.findall(r'\w+', readIn.lower())   
    
    #assign document name to each sentence into dictionary    
    for i in sentences:
        dict_sentences[i] = filename
     
#find 10 most common words
top_words = Counter(cleaned).most_common(10)

#go through the top 10 words and find the occurance in the sentences. Assign the words to a dictionary along with the sentences and the documents they appear in
for i in range(len(top_words)):
    table.append({'Word':top_words[i][0],'Doc':[],'Sentences':[]})
    for j in dict_sentences:
        if top_words[i][0] in re.findall(r'\w+', j.lower())  :
            table[i]['Sentences'].append([j])
            if dict_sentences[j] not in table[i]['Doc']:
                table[i]['Doc'].append(dict_sentences[j]) 
                
    #sort document and sentences into order    
    table[i]['Doc'] =  sorted(table[i]['Doc'])
    table[i]['Sentences'] =  sorted(table[i]['Sentences'])    



#save to excel file            
out = pd.DataFrame(table)
pd.set_option('display.max_columns', None)  
cols = out.columns.tolist()
cols = ['Word', 'Doc', 'Sentences']
out = out[cols]
writer = pd.ExcelWriter('output.xlsx')
out.to_excel(writer,'Sheet1')
writer.save()


    
                       
                   

