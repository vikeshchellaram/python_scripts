import requests
import json
import os
from requests.auth import HTTPBasicAuth
from datetime import datetime
from requests.packages.urllib3.exceptions import InsecureRequestWarning

#To disable SSL Error warning
requests.packages.urllib3.disable_warnings(InsecureRequestWarning)

application_path = os.path.dirname(os.path.realpath(__file__))
os.chdir(application_path)

def jira_connection(ticket_id, action='get'):

    headers = {'content-type': 'application/json'}
    
    username = input('Jira username: ')
    password = input('Jira password: ')
    
    authentication = HTTPBasicAuth(username, password)

    os.environ["HTTP_PROXY"] = ""
    os.environ["HTTPS_PROXY"] = ""
    
    if action=='get':
    
        response = requests.get(f"https://jira.com/rest/api/2{ticket_id}",
                                auth = authentication,
                                headers=headers,
                                verify=False)
        print(response.status_code)
        print(response.text)
        
        data = json.loads(response.text)
    
    elif action=='post':
        
        jira_ticket = create_issue()
        
        response = requests.post("https://jira.com/rest/api/2/issue",
                                 data=json.dumps(jira_ticket),
                                 headers=headers,
                                 auth=authentication,
                                 verify=False)
        
        if response.status_code != 201:
            print(response.status_code)
            print(response.json())
        else:
            print(f"Jira ticket {response.json()['key']} was uploaded successfully.")
            
    else:
        raise ValueError(f"{action} is not an accepted value. Accepted values \
                         are 'post' and 'get'.")
    
def create_issue(jira_ticket_jason='Example_jira.json'):
    
    today = datetime.today().strftime('%d/%m/%Y')
    
    with open(jira_ticket_jason) as json_file:
        jira_ticket = json.load(json_file)
    
    jira_ticket['fields']['description'] = 'Test of jira creation with python'
    jira_ticket['fields']['summary'] =  jira_ticket['fields']['summary'] + ' ' + today

    return jira_ticket
