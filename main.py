import urllib.request as request
import bs4 as bs
import os
import shutil
from docx import Document

#Usage:
#1. Create destination folder on the machine. For example for the month January 2019, create 'January_19 folder
#2. Change 'caseMonth' to the folder name created above ('January_19)
#3. Make sure 'letterTemplate' '.docx' template file exist in 'sourceFolder' folder. 
#4. Open a case url, right click and save as to save it to 'caseMonth' folder. It is adviced to use case id 
#for the file name. Example, case 700 will have the file name 700.html
#5. Run this script
#6. Letters will be created in the folder where html files exists.
#6. Open each output word document letter and review its content and fix any issue found.
#7. One item that need editing will be the description about the reason for this request. Example treatment,
#education etc.
#----------------------

#Month for which letters are generated.
#make sure that Month folder exists and relevant html files exists. Once all reports are created, html files
# may be deleted.
caseMonth = 'November'
sourceFolder = 'C:/Project/HelpSaveLife/'
destFolder = os.path.join( sourceFolder+ caseMonth)
letterTemplate = 'Letter.docx'
urlPrefix = "file:///C:/Project/HelpSaveLife/"+caseMonth+"/"
#read all html files and loop through each one and generate letters.
def execute_main():
    files = os.listdir(destFolder)

    for f in files:
        if not os.path.isdir(f) and f.endswith('.html'):

            sauce = request.urlopen(urlPrefix+f).read()
            soup = bs.BeautifulSoup(sauce, "lxml")

            name = soup.find('span', {'id':'CaseInfo_lblRequesterName', 'class':'Label'}).text
            id = soup.find('span', {'id':'CaseInfo_lblCaseNo', 'class':'Label'}).text
            
            #Destination file and open document
            destFile = copy_rename_file(name,id)
            document = Document(destFile)
            
            #Replace texts
            replace_doc_text(document, '<NAME>',name)
            replace_doc_text(document,'<PAYEE_NAME>', soup.find('span', {'id':'CaseInfo_lblPaymentRecieverName', 'class':'Label'}).text,False)
            replace_doc_text(document,'<ADDRESS1>',soup.find('span', {'id':'CaseInfo_lblAddress1', 'class':'Label'}).text)
            replace_doc_text(document,'<ADDRESS2>', soup.find('span', {'id':'CaseInfo_lblAddress2', 'class':'Label'}).text)
            replace_doc_text(document,'<CITY>', soup.find('span', {'id':'CaseInfo_lblCity', 'class':'Label'}).text)
            replace_doc_text(document,'<STATE>', soup.find('span', {'id':'CaseInfo_lblState', 'class':'Label'}).text+',', False)
            replace_doc_text(document,'<COUNTRY>', soup.find('span', {'id':'CaseInfo_lblCountry', 'class':'Label'}).text, False)
            replace_doc_text(document,'<PIN1>', soup.find('span', {'id':'CaseInfo_lblPostCode', 'class':'Label'}).text)
            replace_doc_text(document,'<PHONE>', "Phone: " +soup.find('span', {'id':'CaseInfo_lblPhone', 'class':'Label'}).text, False)
            replace_doc_text(document,'<MEMBER>', soup.find('span', {'id':'CaseInfo_lblCaseInitiated', 'class':'Label'}).text,False)
            replace_doc_text(document,'<AMOUNT>',soup.find('span', {'id':'CaseInfo_lblAmountPaidInUSD', 'class':'Label'}).text,False)
            replace_doc_text(document,'<ID>', id,False)

            #Save document
            document.save(destFile)


# Copy 'Letter.docx' to new folder (month), and rename the file with user name and ccase id.
# Space and . are replaced with _ on the new file name.
def copy_rename_file(name, id):
    sourceFileName = os.path.join(sourceFolder,letterTemplate)
    #destFolder = os.path.join(source_folder,caseMonth)
    fileName = name.replace(' ', '_')
    fileName = fileName.replace('.','_')
    destFileName = os.path.join(destFolder,'Letter_'+fileName+'_'+id+'.docx')
    if not os.path.exists(destFolder):
        os.mkdir(destFolder)
    
    #Copy template letter and rename.
    #TODO: Check if file exists before copy/rename
    shutil.copy(sourceFileName, destFolder)
    os.rename(os.path.join(destFolder,'Letter.docx'),destFileName)
    return destFileName

#Replace text in the document
def replace_doc_text(document, key, text, appendComma = True):
    for paragraph in document.paragraphs:
        if key in paragraph.text:      
            paragraph.text = paragraph.text.replace(key, text)
            if(appendComma):
                 paragraph.text = paragraph.text + ","

execute_main()
