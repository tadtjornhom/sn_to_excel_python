import json
import requests
import datetime
import re
import os
import mimetypes
import logging
import random
from logging import error
from datetime import datetime

def getMeetingCasestoProcess(url_servicenow,user,pwd):
    # Get list of cases to process
    url = url_servicenow + '/api/now/table/sn_customerservice_case?' + 'sysparm_query=state=1^category%3D5%5Esubcategory%3D800%5Edue_dateRELATIVEGE%40dayofweek%40ago%405&sysparm_fields=account%2Caccount.name%2Ccase%2Csys_id'
    #active=true^category=5^subcategory=800^state=1


    headers = {"Content-Type":"application/json","Accept":"application/json"}
    # Do the HTTP request
    response = requests.get(url, auth=(user, pwd), headers=headers )

    print response
    # Check for HTTP codes other than 200
    if response.status_code != 200:
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
        exit()

    # Decode the JSON response into a dictionary and use the data
    attachmentdata = json.loads(response.text)
    caselist  = attachmentdata["result"]

    return caselist

def updateCaseStatus(url_servicenow,record_sys_id,status,work_notes,user,pwd):

    serviceNOWtable = 'sn_customerservice_case'

    # Set the request parameters
    url = url_servicenow + '/api/now/table/' + serviceNOWtable + '/' + record_sys_id + '?sysparm_input_display_value=true'

    # Set proper headers
    headers = {"Content-Type":"application/json","Accept":"application/json"}

    datavalues = "{\"state\":\"STATUS\",\"work_notes\":\"WORKNOTE\"}"
    datavalues = re.sub('STATUS',status,datavalues)
    datavalues = re.sub('WORKNOTE',work_notes,datavalues)

    try:

        response = requests.put(url, auth=(user, pwd), headers=headers ,data=datavalues)

    # Check for HTTP codes other than 200

        if response.status_code != 200:
            logging.error('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
            return response.status_code

    except requests.exceptions.RequestException as e:
        logging.error(e)
        raise
    finally:
        data = response.json()
        print 'Update Status: ' + str(data)
        return data


    # Decode the JSON response into a dictionary and use the data


    return 'true'


# used to limit the specific attacchment
def getSpecificAttachment(url_servicenow, user,pwd,file_limiter,account_sys_id):
    print 'in get specific'
    fileList_url = url_servicenow + '/api/now/attachment?table_sys_id=' + account_sys_id

    headers = {"Content-Type":"application/json","Accept":"application/json"}
    # Do the HTTP request

    response = requests.get(fileList_url, auth=(user, pwd), headers=headers )
    print response
    # Check for HTTP codes other than 200
    if response.status_code != 200:
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
        exit()

        # Decode the JSON response into a dictionary and use the data
    attachmentdata = json.loads(response.text)
    attachmentdata_list = attachmentdata["result"]

    ## check to see how many match the file name and if ther is a more then one then error
    print attachmentdata_list
    # Error check to see if there is more then one template?

    filecounter = substring_indexes(file_limiter ,str(attachmentdata_list))
    logging.info('On the Account there are: ' + str(filecounter) + ' files. There MUST be only one file')

    if  filecounter != 1:
        if filecounter == 0:
            errorMessage = 'ERROR: There is no templete attached to the Account to process'
            logging.error(errorMessage)
            return errorMessage

        if filecounter > 1:
             errorMessage = 'ERROR: There is more then one Template on the Account and therefore dont know how to process'
             logging.error(errorMessage)
             return errorMessage

    ## Find the write list and process here should only be one. otherwwise we need to re-write this area.
    # hard coded  digits to find the write file name

    for item in attachmentdata_list:
        attachment_sys_id = item.get("sys_id")
        file_name = item.get("file_name")
        table_sys_id = item.get("table_sys_id")
        print file_name + ' ' + file_name[-13:]
        if file_name[-14:] == file_limiter:
            return attachment_sys_id

def substring_indexes(substring, string):
    last_found = -1  # Begin at -1 so the next position to search from is 0
    counter = 0
    while True:
        # Find next index of substring, by starting after its last known position
        last_found = string.find(substring, last_found+1)
        if last_found !=  -1:
            counter = counter+1
        if last_found == -1:
            return counter


def getParentAccount(url_servicenow,account_sys_id,user,pwd):

    url =  url_servicenow + '/api/now/table/customer_account?sysparm_query=parent%3D' +  account_sys_id  # + '&sysparm_limit=1'

    headers = {"Content-Type":"application/json","Accept":"application/json"}
    # Do the HTTP request

    try:
        response = requests.get(url, auth=(user, pwd), headers=headers )

        if response.status_code != 200:
            logging.error('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
            exit()

        # Decode the JSON response into a dictionary and use the data
        data = json.loads(response.text)
        accountList  = data["result"]

    except requests.exceptions.RequestException as e:
        logging.error(e)
        raise
    finally:
        return accountList



def getCaseswithTasks(url_servicenow,case_sys_id,user,pwd):

    # Get list of cases to process
    url = url_servicenow + '/api/now/table/sn_customerservice_task?sysparm_display_value=all&sysparm_query=u_case='+ case_sys_id

    headers = {"Content-Type":"application/json","Accept":"application/json"}
    # Do the HTTP request
    try:
        response = requests.get(url, auth=(user, pwd), headers=headers )

        if response.status_code != 200:
            logging.error('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
            exit()

        # Decode the JSON response into a dictionary and use the data
        data = json.loads(response.text)
        caseandtasklist  = data["result"]

    except requests.exceptions.RequestException as e:
        logging.error(e)
        raise
    finally:
        return caseandtasklist


# download a specific attachment and add time stamp
def download_attachment(url_servicenow,account_sys_id,user,pwd,filedirectory,localfile,attachment_sys_id):


    # file_url = 'https://redbrickhealthdev.service-now.com/api/now/attachment/c8b78f256f255300a1af77f16a3ee46e/file'
    file_url = url_servicenow + '/api/now/attachment/' +  attachment_sys_id + '/file'
    varBinary_headers = {"Contxlsxent-Type":"application/xml","Accept":"application/xml"}

    try:
        varBinary = requests.get(file_url, auth=(user, pwd), headers=varBinary_headers,stream=True)

        if response.status_code != 200:
            logging.error('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
            exit()
    except requests.exceptions.RequestException as e:
        logging.error(e)
        raise
    finally:
        json_obj = varBinary.headers['x-attachment-metadata']
        json_objMetadata = json.loads(json_obj)

        for k,v in json_objMetadata.items():

            if k == 'file_name':
                file_name = v

        #build file dir
        file_name = re.sub('.xlsx', '',file_name)

        vardatetime = str(datetime.now().strftime('%Y-%m-%d-%I%M%S')) +'_' + str(random.randint(1,1001))

        file_name =  file_name + '_' + vardatetime + '.xlsx'


        # with open('/Users/ttjornhom/IdeaProjects/snCSMExcelDoc/template_tracker99.xlsx', 'w+') as f:
        with open(file_name, 'w+') as f:
            for chunk in varBinary:
                f.write(chunk)

        if os.path.isfile(file_name) == 'true':
            logging.info('validated File: ' + file_name)

        return file_name


def getcasesTOupdateExcel(url_servicenow,user,pwd,account_sys_id):
    caselist = {}
    url = url_servicenow + '/api/now/table/sn_customerservice_case?sysparm_display_value=true&sysparm_query=categoryIN2%2C5%2C3%2C6%2C4%2C10%5Eaccount%3D' + account_sys_id

    headers = {"Content-Type":"application/json","Accept":"application/json"}
    # Do the HTTP request
    try:
        response = requests.get(url, auth=(user, pwd), headers=headers )

        if response.status_code != 200:
            logging.error('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
            exit()

        # Decode the JSON response into a dictionary and use the data
        data = json.loads(response.text)
        caselist  = data["result"]

    except requests.exceptions.RequestException as e:
        logging.error(e)
        raise
    finally:
        return caselist



def postfiletoServiceNOW(url_servicenow,file_name,user,pwd,table_name,table_sys_id):
    #TODO find current work on uploading a file

    print 'In POST FILE'


    content_type,fileEncoding = mimetypes.guess_type(file_name)
    print 'In POST FILE:content_type:' + str(content_type)

    headers = {
        'Accept': 'application/json',
        'Content-Type': content_type,
    }

    # specify files to send as binary
    data = open(file_name, 'rb').read()

    file_url = url_servicenow + '/api/now/attachment/file?table_name=' + table_name + '&table_sys_id=' + table_sys_id+ '&file_name=' + file_name

    print file_url

    response = requests.post(file_url, auth=(user,pwd), headers=headers, data=data)

    print response

    # this rest api indicates good as 201 and not 200
    if response.status_code != 201:
        print('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
        exit()

    attachmentdata = json.loads(response.text)

    #TODO fixLOOP

    for k,v in attachmentdata.items():
        if k == 'file_name':
            file_name = v

    return file_name

def getAccountinfo(url_servicenow,account_sys_id,user,pwd):
    # Set the request parameters
    acct_url = url_servicenow + '/api/now/table/customer_account?sysparm_limit=1&sys_id='+ account_sys_id
    # Set proper headers
    headers = {"Content-Type":"application/json","Accept":"application/json"}
    # Do the HTTP request

    try:
        response = requests.get(url, auth=(user, pwd), headers=headers )

        if response.status_code != 200:
            logging.error('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
            exit()

        # Decode the JSON response into a dictionary and use the data
        data = json.loads(response.text)
        returnlist  = data["result"]

    except requests.exceptions.RequestException as e:
        logging.error(e)
        raise
    finally:

        return returnlist



def getgeneralTableinfo(url_servicenow,table,sysparms,user,pwd):

    # Set the request parameters
    url = url_servicenow + '/api/now/table/'+table + sysparms


    # Set proper headers
    headers = {"Content-Type":"application/json","Accept":"application/json"}
    try:

        response = requests.get(url, auth=(user, pwd), headers=headers )

        if response.status_code != 200:
            logging.error('Status:', response.status_code, 'Headers:', response.headers, 'Error Response:',response.json())
            exit()

        # Decode the JSON response into a dictionary and use the data
        data = json.loads(response.text)
        returnlist  = data["result"]

    except requests.exceptions.RequestException as e:
        logging.error(e)
        raise
    finally:
        return returnlist

