import docx
import os
import csv

'''
This is the customised version of wordToCSV.py for the song lyrics project. There is a minor bug is that
for some reason 100<10<1 which might have been some sorting algorithm going on in the background, but its getting late
so i'm not fixing that. We can easily sort that out in excel.
'''

def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    return '\n'.join(fullText)

files = [f for f in os.listdir('.') if os.path.isfile(f)]
datalist = []
header = ['Ranking', 'Title', 'Artist', "Lyrics"]

for filename in files:
    print(filename)
    txt = filename.split("_")
    if len(txt) > 2:
        lyrics = getText(filename)
        ranking = txt[0]       
        title = txt[1]
        artist = txt[2].removesuffix(".docx")
        datalist.append([ranking, title, artist, lyrics])

print(datalist)
print(files)

with open('songs.csv', 'w', encoding='UTF8', newline="") as f:
    writer = csv.writer(f)
    writer.writerow(header)
    writer.writerows(datalist)

