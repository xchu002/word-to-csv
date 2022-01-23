import docx
import os
import csv

'''
This is a program that can transfer the contents of many word documents at once into a csv file with two columns: the 
Document Name and the Content.

Usage:
Download this file and put it in the same folder as the docx files. Run this file using vscode/spyder or any code editor
you wish.
There should be a file named output.csv created in the same folder, which contains the output.
'''

#function to get the text in word docs
def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

#returns all the file names in the same directory as a list of strings
files = [f for f in os.listdir('.') if os.path.isfile(f)]

#empty list that will be filled when the for loop below is run
datalist = []


#loops through all the files in the current directory
for filename in files:

    #appends the filename and the file content into datalist above
    if ".docx" in filename: 
        print(filename) #for testing purposes
        filecontent = getText(filename)   
        datalist.append([filename.removesuffix(".docx"), filecontent])


header = ['Document Name', 'Content']
#writes into csv file
with open('output.csv', 'w', encoding='UTF8', newline="") as f:
    writer = csv.writer(f)
    writer.writerow(header)
    writer.writerows(datalist)