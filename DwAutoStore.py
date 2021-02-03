from Logger import Logger
import docx
import os
import time
import traceback
import subprocess
import re

#runtime Params
running=True
idle=10
watchFile="\\\\mattarfs03\\sharedfiles\\storeToDocuware"
dwStorer="DWStorer\DWStorer.exe"

#objects 
class job(object):
    file=""
    status="NEW"
    casenum=0
    name=""
    provider=""
    category=""
    doctype=""
    docdate=""
    notes=""
    filename=""
    docnum=""
    provider2=""
    category2=""
    doctype2=""
    def __init__(self, file):
        self.docdate=time.strftime('%Y-%m-%d')
        self.file=file
        self.filename=file.split("\\")[-1]
    
#functions
def scanFiles(directory):
    wordFiles=[]
    try:
        files = [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]
    except:
        Logger.writeAndPrintLine("Could not get directory listing for "+directory+". Sleeping for 5 minutes. "+traceback.format_exc(),3)  
        time.sleep(300)
        return
    for file in files:
        fullPath=directory+'\\'+file
        if(".DOCX" in fullPath.upper() or ".DOC" in fullPath.upper()):
            wordFiles.append(fullPath)
    return wordFiles

def pruneDuplicates(newFiles, jobs):
    originalFiles=[]
    for file in newFiles:
        original=True
        for tempJob in jobs:
            if(file==tempJob.file):
                original=False
                break
        if(original):
            originalFiles.append(file)
    
    return originalFiles

def pruneErroredJobs(newFiles, jobs):
    for tempJob in jobs:
        if(tempJob.status=="ERROR"):
            stillExists=False
            for file in newFiles:
                if(file==tempJob.file):
                    stillExists=True
                    break
            if(not stillExists):
                #file for failed job no longer exists, remove ghost job
                jobs.remove(tempJob)
        
def isFileLocked(file):
    try:
        myfile = open(file, "r+") # or "a+", whatever you need
        myfile.close()
        return False
    except IOError:
        return True

def pullHiddenFields(job):
    #https://automatetheboringstuff.com/chapter13/
    doc = docx.Document(job.file)
    fullText = []
    for para in doc.paragraphs:
        fullText.append(para.text)
    fulltext='\n'.join(fullText)
    #print(fulltext)
    
    job.casenum=re.search('dwCasenum=(.+)',fulltext,0).group(1)
    job.name=re.search('dwName=(.+)',fulltext,0).group(1)
    job.category=re.search('dwCategory=(.+)',fulltext,0).group(1)
    job.doctype=re.search('dwDocumentType=(.+)',fulltext,0).group(1)
    job.provider=re.search('dwProvider=(.+)',fulltext,0).group(1)
    job.docnum=re.search('docNum=(.+)',fulltext,0).group(1)
    job.casenum=job.casenum.replace(',','')
    
    if(job.docnum=="6"):
        try:
            job.category2=re.search('dwCategory2=(.+)',fulltext,0).group(1)
            job.doctype2=re.search('dwDocumentType2=(.+)',fulltext,0).group(1)
            job.provider2=re.search('dwProvider2=(.+)',fulltext,0).group(1)
        except:
            job.docnum="6_1"

def storeToDocuware(tempJob):
        dwStorer="DWStorer\DWStorer.exe"
        params=[]
        params.append(dwStorer)
        params.append('-c')
        params.append('-iCASE_ID:"'+tempJob.casenum+'"')
        params.append('-iLAST_NAME:"'+tempJob.name+'"')
        params.append('-iPROVIDER:"'+tempJob.provider+'"')
        params.append('-iCATEGORY:"'+tempJob.category+'"')
        params.append('-iDOCUMENT_TYPE:"'+tempJob.doctype+'"')
        params.append('-iSTATUS:"NEW"')
        params.append('-iSTOREDBY:"AutoStorer"')
        params.append('-f"'+tempJob.file+'"')
        out=subprocess.check_output(params)
        out=out.decode('ascii')
        if(out=="success\r\n"):
            if(tempJob.docnum=="6"):
                return storeToDocuware2(tempJob)
            return None
        else:
            return out
        
def storeToDocuware2(tempJob):
        dwStorer="DWStorer\DWStorer.exe"
        params=[]
        params.append(dwStorer)
        params.append('-c')
        params.append('-iCASE_ID:"'+tempJob.casenum+'"')
        params.append('-iLAST_NAME:"'+tempJob.name+'"')
        params.append('-iPROVIDER:"'+tempJob.provider2+'"')
        params.append('-iCATEGORY:"'+tempJob.category2+'"')
        params.append('-iDOCUMENT_TYPE:"'+tempJob.doctype2+'"')
        params.append('-iSTATUS:"NEW"')
        params.append('-iSTOREDBY:"AutoStorer"')
        params.append('-f"'+tempJob.file+'"')
        out=subprocess.check_output(params)
        out=out.decode('ascii')
        if(out=="success\r\n"):
            return None
        else:
            return out

def cleanupJob(job):
    os.unlink(job.file)

def removeDeletedJobs(jobs):
    i=len(jobs)-1
    while(i>=0):
        if(jobs[i].status=='DELETED'):
            del jobs[i]
        i=i-1
    
#runtime vars
jobs=[]
#run block
Logger.writeAndPrintLine("Application started.",1)
while(running):
    files=scanFiles(watchFile)
    pruneErroredJobs(files, jobs)
    files=pruneDuplicates(files, jobs)
    for file in files:
        if(not isFileLocked(file)):
            jobs.append(job(file))
    for tempJob in jobs:
        if(tempJob.status=="NEW"):
            try:
                Logger.writeAndPrintLine("Starting job for "+tempJob.file,1)
                pullHiddenFields(tempJob)
                tempJob.status="PULLED"
            except: 
                Logger.writeAndPrintLine("Error pulling hidden fields for file "+tempJob.file+", "+traceback.format_exc(),3)
                tempJob.status="ERROR"
        if(tempJob.status=="PULLED"):
            try:
                status=storeToDocuware(tempJob)
                if(status==None):
                    tempJob.status='COMPLETE'
                    Logger.writeAndPrintLine("Stored "+tempJob.file+" to Docuware.",1)
                else:
                    Logger.writeAndPrintLine("Error running DWStorer "+tempJob.file,3)
                    tempJob.status='ERROR'
            except:
                Logger.writeAndPrintLine("Error storing file to Docuware "+tempJob.file,3)
                tempJob.status='ERROR'
        if(tempJob.status=="COMPLETE"):
            try:
                cleanupJob(tempJob)
                Logger.writeAndPrintLine("Cleaned up original file "+tempJob.file+".",1)
                tempJob.status='DELETED'
            except:       
                Logger.writeAndPrintLine("Error deleting file "+tempJob.file+", "+traceback.format_exc(),3)
                tempJob.status='ERROR'
        
        removeDeletedJobs(jobs)
    time.sleep(idle)
        