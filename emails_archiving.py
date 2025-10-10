from win32com.client import Dispatch
import re
import os
import datetime as dt
from datetime import datetime, timedelta
from pythoncom import com_error
import configparser


def read_credentials_from_ini(ini_path, section):
    if os.path.isfile(ini_path):
        config = configparser.ConfigParser()
        try:
            config.read(ini_path)
            lookback_days = config[section]['lookback_days']
            year = config[section]['year']
            exclude_folders = config[section]['exclude_folders']
            print('%s retrieved succefully!' %section)
        except KeyError:
            lookback_days = None
            year = None
            exclude_folders = None
            print('Make sure a section [%s] exists in the ini file.' %section)
    else:
        folder, _ = os.path.split(ini_path)
        print('Make sure an ini file is located at %s' %folder)
        year = None
        lookback_days = None
        exclude_folders = None
        
    return lookback_days, year, exclude_folders


def get_folders_in_darwin_node(outdict,node):
    folder_id= node
    list_user = darwin.listdir(folder_id)
    folders_name = [node[2].name for node in list_user]
    folders_id = [node[0] for node in list_user]
    outdict = dict(map(lambda i,j : (i,j) , folders_name,folders_id))
    return outdict


def get_outlook_folders_and_subfolders(outlook_account,account, folders, subfolders, lowerfolders): 
    global folder
    for folder in outlook_account.folders:
        # print(folder.Name)
        folders.append(folder.Name)
        # for msg in folder.Items:
        #     print(str(msg.Subject))
        for subfolder in folder.Folders:
            # print(subfolder.Name)
            subfolders.append(subfolder.Name)
            # for msg in subfolder.Items:
            #     print(str(msg.Subject))
            for lowerfolder in subfolder.Folders:
                lowerfolders.append(subfolder.Name)
            
    folders = [x for x in folders if x not in exclude_dict[f"{account}"]]
    print(f"Folders: {folders}")
    subfolders = [x for x in subfolders if x not in exclude_dict[f"{account}"]]
    print(f"Subfolders: {subfolders}")
    lowerfolders = [x for x in lowerfolders if x not in exclude_dict[f"{account}"]]
    print(f"Lowerfolders: {lowerfolders}")

    
def create_local_folders(temp_folder, folder, account): 
    path = os.path.join(temp_folder, str(folder))
    # to create local folders if not existing
    try:
        os.mkdir(path)
    except FileExistsError:
        pass
    sub_folder = account + " - " + str(folder) + " - " + date.strftime("%Y")
    sub_path = os.path.join(path, sub_folder)
    try:
        os.mkdir(sub_path)
    except FileExistsError:
        pass
    return sub_folder, sub_path


def rename_emails(msg, sub_path):
    name = str(msg.subject) + "-" + msg.ReceivedTime.strftime('%d/%m/%Y %H:%M %p')
     
    # to eliminate any special charecters in the name
    name = re.sub('[^A-Za-z0-9._-]+', ' ', name)+'.msg'
    # to save the email in temp_folder
    msg_path = sub_path +'//'+name
    msg.SaveAs(msg_path)
    
    global emails_read
    emails_read+=1
    print(name) 
    return msg_path, name


def create_darwin_folder(darwin_node_id, folder):
    try: 
        darwin.create_folder(parent_node_id=darwin_node_id, name=str(folder))
    except NodeExistsError:
        pass
    
    
def create_darwin_document(sub_node_id, name, msg_path):
    try:
        darwin.create_document(parent_node_id=sub_node_id, name=name, file=msg_path)
        global emails_archived
        emails_archived+=1
    except Exception:
        print(f"{Exception} Email already archived")
        global exceptions
        exceptions+=1


def delete_email(msg_path, msg):
    os.remove(msg_path)
    global emails_deleted
    emails_deleted+=1
    msg.Delete()
    

def process_emails():
    for account in outlook_dict:
        print(f"Outlook account: {account}")
        outlook_account=outlook.Folders[account].Folders['Inbox'].Folders['Archiving']
        
        folders = []
        subfolders = []
        lowerfolders = []
        get_outlook_folders_and_subfolders(outlook_account,account,folders,subfolders, lowerfolders)
        print("")

        darwin_node_id = outlook_dict[f"{account}"]
        temp_folder = local_folder[f"{account}"]
            
        for folder in outlook_account.folders:
            if folder.Name in exclude_dict[f"{account}"]:
                continue
            else:
                sub_folder, sub_path = create_local_folders(temp_folder, folder, account) 

                print(f"Started processing folder: {folder}")
                messages = folder.items
                #filter messages based on lookback and year 
                restriction = ("[ReceivedTime] >= '{0}/01/01 00:00' AND [ReceivedTime] <= '{0}/12/31 23:59'" "AND [ReceivedTime] <= '{1}'").format(year, lookback_date_string)
                filtered_messages = messages.Restrict(restriction)
                
                create_darwin_folder(darwin_node_id, folder)
                            
                # check darwin nodes
                darwin_folders = {}
                darwin_folders = get_folders_in_darwin_node(darwin_folders,darwin_node_id)
                
                # check darwin sub-folders
                if str(folder) in darwin_folders:
                    node_id = darwin_folders[f"{folder}"]
                    try:
                        darwin.create_folder(parent_node_id=node_id, name=sub_folder)
                    except NodeExistsError:
                        pass
                                
                for msg in filtered_messages:
                    msg_path, name = rename_emails(msg, sub_path)
                    
                    darwin_sub_folders = {}
                    darwin_sub_folders = get_folders_in_darwin_node(darwin_sub_folders,node_id)
                    
                    # to add emails in sub-folders
                    for folder in darwin_sub_folders:
                        if date.strftime('%Y') in folder:
                            sub_node_id = darwin_sub_folders[f"{sub_folder}"]
                            
                            create_darwin_document(sub_node_id, name, msg_path)
                            
                            delete_email(msg_path, msg)
                print("-------------------------------------------------------------------")
                
                try:
                    for subfolder in folder.folders:
                        if subfolder.Name in exclude_dict[f"{account}"]:
                            continue
                        else:
                            sub_folder, sub_path = create_local_folders(temp_folder, subfolder, account)        
                                            
                            print(f"Started processing subfolder: {subfolder}")
                            messages = subfolder.items
                            #filter messages based on lookback and year 
                            restriction = ("[ReceivedTime] >= '{0}/01/01 00:00' AND [ReceivedTime] <= '{0}/12/31 23:59'" "AND [ReceivedTime] <= '{1}'").format(year, lookback_date_string)
                            filtered_messages = messages.Restrict(restriction)
                            
                            create_darwin_folder(darwin_node_id, subfolder)
                                        
                            # check darwin nodes
                            darwin_folders = {}
                            darwin_folders = get_folders_in_darwin_node(darwin_folders,darwin_node_id)
                            
                            # check darwin sub-folders
                            if str(subfolder) in darwin_folders:
                                node_id = darwin_folders[f"{subfolder}"]
                                try:
                                    darwin.create_folder(parent_node_id=node_id, name=sub_folder)
                                except NodeExistsError:
                                    pass
                                            
                            for msg in filtered_messages:
                                msg_path, name = rename_emails(msg, sub_path)
                                
                                darwin_sub_folders = {}
                                darwin_sub_folders = get_folders_in_darwin_node(darwin_sub_folders,node_id)
                                
                                # to add emails in sub-folders
                                for subfolder in darwin_sub_folders:
                                    if date.strftime('%Y') in subfolder:
                                        sub_node_id = darwin_sub_folders[f"{sub_folder}"]
                                        
                                        create_darwin_document(sub_node_id, name, msg_path)

                                        delete_email(msg_path, msg)
                            print("-------------------------------------------------------------------")
                            
                            try: 
                                for lowerfolder in subfolder.folders:
                                    if lowerfolder.Name in exclude_dict[f"{account}"]:
                                        continue
                                    else:
                                        sub_folder, sub_path = create_local_folders(temp_folder, lowerfolder, account)        
                                                        
                                        print(f"Started processing lowerfolder: {lowerfolder}")
                                        messages = lowerfolder.items
                                        #filter messages based on lookback and year 
                                        restriction = ("[ReceivedTime] >= '{0}/01/01 00:00' AND [ReceivedTime] <= '{0}/12/31 23:59'" "AND [ReceivedTime] <= '{1}'").format(year, lookback_date_string)
                                        filtered_messages = messages.Restrict(restriction)
                                        
                                        create_darwin_folder(darwin_node_id, lowerfolder)
                                                    
                                        # check darwin nodes
                                        darwin_folders = {}
                                        darwin_folders = get_folders_in_darwin_node(darwin_folders,darwin_node_id)
                                        
                                        # check darwin sub-folders
                                        if str(lowerfolder) in darwin_folders:
                                            node_id = darwin_folders[f"{lowerfolder}"]
                                            try:
                                                darwin.create_folder(parent_node_id=node_id, name=sub_folder)
                                            except NodeExistsError:
                                                pass
                                                        
                                        for msg in filtered_messages:
                                            msg_path, name = rename_emails(msg, sub_path)
                                            
                                            darwin_sub_folders = {}
                                            darwin_sub_folders = get_folders_in_darwin_node(darwin_sub_folders,node_id)
                                            
                                            # to add emails in sub-folders
                                            for lowerfolder in darwin_sub_folders:
                                                if date.strftime('%Y') in lowerfolder:
                                                    sub_node_id = darwin_sub_folders[f"{sub_folder}"]
                                                    
                                                    create_darwin_document(sub_node_id, name, msg_path)
                                                    
                                                    delete_email(msg_path, msg)
                                        print("-------------------------------------------------------------------")
                        
                            except AttributeError:
                                pass
                
                except AttributeError:
                    pass

    
# %%
if __name__ == "__main__":
    
    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")

    path = ""
    ini_path = f"{path}\emails_archiving.ini"
    
    lookback_days, year, exclude_folders = read_credentials_from_ini(ini_path, "Inputs")
    
    # to filter emails:
    start = dt.datetime.now()
    lookback = start - timedelta(days = int(lookback_days))
    date = dt.datetime(year=int(year), month=1, day=1)
    lookback_date_string = lookback.strftime('%d/%m/%Y %H:%M %p')
    
    # to specify outlook accounts and associated nodes
    outlook_dict = {'example@email.com': 1234567891}
    
    # local path to store emails
    local_folder = {list(outlook_dict.keys())[0]: 'C:\example'}
    
    exclude_dict = {list(outlook_dict.keys())[0]: f'{exclude_folders}'}

    emails_archived = 0
    emails_read = 0
    exceptions = 0
    emails_deleted = 0
    
    process_emails()
    
print (f"""
       ########################## Final Summary ##########################
Emails read {emails_read}
Exceptions {exceptions}
Emails archived {emails_archived}
Emails deleted from temp folder {emails_deleted}""")

end = dt.datetime.now() - start

print("------------------------------------")
print (f"""
Emails archiving script finished successfully.
It took {end} to finish.
Have a nice day {os.getlogin()}.""")

               
        
        
    

