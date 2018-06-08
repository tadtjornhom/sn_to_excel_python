import sys
import datetime
import re
import os
import mimetypes
import logging


#local files
import createNewSheet
import servicenowRestAPI

from os import environ
from datetime import datetime
from operator import itemgetter, attrgetter, methodcaller
from logging import error


# custom variables
filedirectory = os.getenv('sn_filedirectory')
localTemplate = ''
localfile = ''
accountName = ''
user = os.getenv('sn_user')
pwd = os.getenv('sn_pwd_prod')
newFilepath = ''
url_servicenow = os.getenv('url_servicenow')
testingStopAfter = 'false'




# Process Queue of cases to attach documents
#TODO FixFile cleanup to delete the file off the drive

def postfilecleanup(filepath):

    for root, dirs, files in os.walk("."):  
        for filename in files:
            if filename.startswith("redbrick_"):
                print filename
                os.remove(filename)

    if os.path.isfile(filepath):
        os.remove(filepath)
        return 0
    else:
        logging.info("no file")
        return 1




def main(user,pwd):
    
    # to use dev run command line python main.py 'DEV'
    
    logging.basicConfig(filename='error.log',level=logging.DEBUG, filemode="w")
    #logger = logging.getLogger(__name__)
    logger = logging.getLogger('caseapp')
    logger.setLevel(logging.DEBUG)
    fh = logging.FileHandler('error.log')
    fh.setLevel(logging.DEBUG)
    logger.addHandler(fh)

    if sys.argv[1] == 'DEV':                
        url_servicenow = 'https://redbrickhealthdev.service-now.com'
        pwd = os.getenv('sn_pwd_dev')
        logging.info('USING DEV INSTANCE: ' + url_servicenow )
    if sys.argv[1] == 'PROD':
        url_servicenow = 'https://redbrickhealth.service-now.com'
        pwd = os.getenv('sn_pwd_prod')
        logging.info('USING PROD INSTANCE: ' + url_servicenow )   

#GET THE MAINLING LIST
    meetinglist = servicenowRestAPI.getMeetingCasestoProcess(url_servicenow,user,pwd)

    logging.info('There are ' + str(len(meetinglist)) + ' meeting records to process')


    # 1. for each case to the following
    # 2. change the status to open and write a note
    # 3. download the template locally
    # 4. modify the template to have good data
    # re-attach the new file
    # close out ticket

    for item in meetinglist:

        # change the status to open and write a note
        case_sys_id = item.get("sys_id")
        accountName = item.get("account.name")
        account_sys_id = item.get("account").get("value")

        logging.info('Processing Meeting Request\n' + 'AccountName= '+ accountName + ' for Case= ' + case_sys_id )
        logging.info('account_sys_id: ' + item.get("account").get("link"))
        logging.info('https://redbrickhealthdev.service-now.com/nav_to.do?uri=sn_customerservice_case.do?sys_id=' + item.get("sys_id") + '&sysparm_view=case')

       # 'https://redbrickhealthdev.service-now.com/nav_to.do?sn_customerservice_case.do%3Fsys_id%3D3f4294c26f219300a1af77f16a3ee4ab%26sysparm_view%3Dcase%26sysparm_record_target%3Dsn_customerservice_case%26sysparm_record_row%3D1%26sysparm_record_rows%3D127%26sysparm_record_list%3Dcategory%2521%253D4%255EORDERBYDESCsys_updated_on

        # 2. change the status to open and write a note
        # TODO change from NEW TO OPEN THEN DONE TESTING
        logging.info('Updating case status = Open')
        status = 'Open'
        work_notes = 'Currenlty Processing this record to make an agenda'

        servicenowRestAPI.updateCaseStatus(url_servicenow,case_sys_id,status,work_notes,user,pwd)

        logging.info('Completing case status = Open')
        # 3. download the template locally
        # 3.1 get the one file to process

        logging.info('Get Specific attachment = *_template.xlsx')

        attachment_api_id = servicenowRestAPI.getSpecificAttachment(url_servicenow,user,pwd,'_template.xlsx',account_sys_id)

        if "ERROR" in attachment_api_id:
             servicenowRestAPI.updateCaseStatus(url_servicenow,case_sys_id,'OPEN',attachment_api_id,user,pwd)
             logging.info('Moving on to the next Meeting Records because we couldnt download a template')
             break

        logging.info('download attachment= ' + attachment_api_id )
        newFilepath = servicenowRestAPI.download_attachment(url_servicenow,account_sys_id,user,pwd,filedirectory,localfile,attachment_api_id)

        logging.info('New path for attachment= ' + newFilepath)

        # 4. modify the template to have good data
        # 4.1 get case data so that we can create a sheet

        #Processing the all the records to file
        logging.info('Before Create Sheet Processor:' + accountName + ' | ' + newFilepath + ' | ' + account_sys_id )
        createNewSheet.createSheetProcesser(url_servicenow,user,pwd,newFilepath,accountName,account_sys_id)

        logging.info('After Create Sheet:' + accountName + ' | ' + newFilepath + ' | ' + account_sys_id )

        #TODO TESTING UPDATE CASE WITH EXCEL REMOVE EXIT TO POST RECORD
        if testingStopAfter == 'true':
            logging.info('IN Stop after testing so code is STOPPED before cleanup')
            exit()

        logging.info('Before Post Cleanup')
        postfileandcleanup(url_servicenow,newFilepath,user,pwd,"sn_customerservice_case",case_sys_id)


def postfileandcleanup(url_servicenow,newFilepath,user,pwd,tableName,case_sys_id):



        # post the new file to the case that we are processing

        postreturn = servicenowRestAPI.postfiletoServiceNOW(url_servicenow,newFilepath,user,pwd,'sn_customerservice_case',case_sys_id)
        logging.info('afterupdate excel:' + postreturn)

        # Clean up and notifiy everyone this this complete
        postfilecleanup(newFilepath)
        # post file update case that the file is attached and ready

        worknotes = 'Excel file template has been generated and attached. Please validate'
        status = 'Resolved'

        response =  servicenowRestAPI.updateCaseStatus(url_servicenow,case_sys_id,status,worknotes,user,pwd,)
        logging.info('For Account:' + accountName + ' the following case is complete \n' + status )

        logging.info('Process Completed =========================')


if __name__ == '__main__':
    main(user,pwd)
