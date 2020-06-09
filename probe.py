
import win32com.client as win32
from pandas import DataFrame
import glob,os

#import sys
def ln(x):
    n = 1000.0
    return n * ((x ** (1/n)) - 1)
def termFrequency(term, document):
    normalizeDocument = document.lower().split()
    #print(normalizeDocument.count(term.lower()))
    #print(float(len(normalizeDocument)))
    return normalizeDocument.count(term.lower()) / float(len(normalizeDocument))

def idfrequency(term,doccuments):
    sum=0
    for docm in doccuments:
        considered=docm.lower().split()
        c=considered.count(term.lower())
        #print(considered)
        if(c>0):
            sum+=1
    return ln(float(len(doccuments)/(1+sum)))

keywordstring = ""
Files = []
Term = []
name=[]
extensions = ['*.docx','*.doc','*.rtf']
word = win32.gencache.EnsureDispatch('Word.Application')

for e in extensions:
    for infile in glob.glob( os.path.join('',e) ):
        Files.append(infile)
def getdocument(infile):
    w = ""
    doc = word.Documents.Open(os.getcwd() + '\\' + infile)
    for each_word in doc.Words:
        text = each_word.Text
        w += " "
        for i in text:
            if ((i >= 'a') & (i <= 'z')) | ((i >= 'A') & (i <= 'Z')) | (i == '-'):
                w += i
            else:
                break
    return w

doccuments=[]
result=[]
for infile in Files:
    if '~' in infile:
        os.remove(infile)
        continue
    if(infile=="keyword.docx"):

        print("analysing requirement " + infile)
        doc = word.Documents.Open(os.getcwd() + '\\' + infile)
        for each_word in doc.Words:

            text = each_word.Text
            keywordstring += " "
            for i in text:
                if ((i >= 'a') & (i <= 'z')) | ((i >= 'A') & (i <= 'Z')) | (i == '-') | (i == "+") | (i == "#") | (i=='2'):
                    keywordstring+=i
                else:
                    break
for infile in Files:
    if '~' in infile:
        os.remove(infile)
        continue
    if infile=="keyword.docx":
        continue

    choice='Y'
    print("Resume analysing"+ infile)
    if choice == 'Y':
        w = ""
        doc = word.Documents.Open(os.getcwd()+'\\'+infile)
        for each_word in doc.Words:
            text = each_word.Text
            w+=" "
            for i in text:
                if ((i >= 'a') & (i <= 'z')) | ((i >= 'A') & (i <= 'Z')) | (i == '-'):
                    w+=i
                else:
                    break
        doccuments.append(w)

normalized=keywordstring.lower().split()
dcn={}
nm="Nmae"

#print(normalized)
for term in normalized:
    idf=idfrequency(term,doccuments)
    dcn[term]=idf

print(dcn)
for file in Files:
    if '~' in file:
        os.remove(infile)
        continue
    if file == "keyword.docx":
        continue
    doc=getdocument(file)
    sum=0
    for term in normalized:
        tf=termFrequency(term,doc)
        idf1=dcn[term]
        sum=sum+float(tf*idf1)
    result.append(sum)
    dcn[file]=sum
resultfinal=[]


for file in Files:
    if '~' in file:
        os.remove(file)
        continue
    if file == "keyword.docx":
        continue
    name.append(file)
    #if dcn[file]>0:
    resultfinal.append(dcn[file])
    #else:
        #resultfinal.append(0)

#print(len(name))
#print("Result length"+str(len(result)))
df=DataFrame({'Name': name, 'TF-IDF ': resultfinal})
print(df)
df.to_excel("final.xlsx",sheet_name="Sheet1",index=False)

print(result)
print(dcn)