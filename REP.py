import os
import sys
import shutil
import zipfile
import re
from os.path import basename

def createFolder(foldername):

    try:
        #Create Directory
        os.mkdir(foldername)
        print("Directory " , foldername , "created! ")
    except FileExistsError:
        print("Directory" , foldername , "already exists!")

def reppydirs():
    # Get current workdir
    global reppyimport, reppytemp, reppyexport, workindir
    workindir = os.getcwd()
    # create Required Folders
    reppyimport = workindir + '\\import'
    reppytemp = workindir + '\\temp'
    reppyexport = workindir + '\\export'
    
    createFolder(reppyimport)
    createFolder(reppytemp)
    createFolder(reppyexport)



# Replaces the Password String inside the XML Sheets
def findPasswordLine(inputSource,tempdir,filename):
    input = open(inputSource)
    output = open(tempdir + '\\' + filename,'w')
    for line in input:
        # REgex Helper ;) https://pythex.org/
        output.write(re.sub('<sheetProtection.*?.>', '', line))
    input.close()
    output.close()

def writeback2excel(zip_file,workindir):
    # 7Zip Locations
    z7location = workindir + '\\7z\\7za.exe'
    # Export with 7z 
    os.system(f'{z7location} u "{zip_file}" "{workindir}\\temp\\export\\*" ')

def cleanupTemp(tempfolder):
    print('Deleting folder ' + tempfolder + ' !')
    try:
        shutil.rmtree(tempfolder)
    except:
        print('Could not delete ' + tempfolder + ' folder. Please delete manually!')

def main(rimport,rtemp,rexport):

    listOfFile = os.listdir(rimport)
    for file in listOfFile:
        print(file)
        # Copy file to work with
        src_dir = rimport + '\\' + file 
        dst_temp_dir = rtemp + '\\temp_' + file
        dst_exp_dir = rexport + '\\[REPpy]' + file
        shutil.copy(src_dir,dst_temp_dir)
        shutil.copy(src_dir,dst_exp_dir)
        # Extract XML Sheet Files
        with zipfile.ZipFile(dst_temp_dir,'r') as zip_ref:
            ziptmp = rtemp + '\\zip'
            createFolder(ziptmp)
            zip_ref.extractall(ziptmp)
            # Required to later mount it perfectly into existing Excel File
            tempexp = rtemp + '\\export'
            createFolder(tempexp)
            tempxl = tempexp + '\\xl'
            createFolder(tempxl)
            srtemp = tempxl + '\\worksheets'
            createFolder(srtemp)
            modxmlpath = ziptmp + '\\xl\\worksheets'
            modxmltree = os.listdir(modxmlpath)
            # Modify the XML Files
            for mxml in modxmltree:
                if(mxml.endswith('.xml')):
                    print(mxml)
                    findPasswordLine(modxmlpath + '\\' + mxml,srtemp,mxml)
            # Write the Files back to Excel
            writeback2excel(dst_exp_dir,workindir)
            # Cleanup the existing files to avoid corruption on multiple Excel Inputs
            cleanupTemp(ziptmp)
    # Delete the Temp folder. Noone needs it afterwards   
    cleanupTemp(reppytemp)     
    return

## Execution
reppydirs()
main(reppyimport,reppytemp,reppyexport)