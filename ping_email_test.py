import win32com.client
from pythonping import ping
from subprocess import Popen
import pandas as pd
from sqlalchemy import create_engine,text
import toml
from datetime import datetime
from functions.logging_insert_db import insert_db


#fetch dbcredentials
credsfile = r"C:\Users\jcohen\Documents\ping_email_script\secrets.toml"
content  = toml.load(credsfile)
conn_string = content['connection_string']
local_ip = content['local_ip']
HV1_ip = content['HV1_ip']
engine = create_engine(conn_string)


outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

# Ping Curtis-HV1 server
Response = ping(f'{HV1_ip}')

# Logging Variables
date_time = datetime.now()
avg_reply_time = str(Response).split('/')[3]


if str(Response).__contains__('Reply'):
    server_test = "success"
    print(f"{server_test}")
    
    # Create Log event dataframe
    ping_data = {'datetime' : [date_time], 'result' : [server_test], 'avg_reply_time' : [avg_reply_time]}
    ping_event = pd.DataFrame(ping_data)

    #insert row in dataframe
    insert_db(ping_event, engine, "ping_event_log")

else:
    #pass
    server_test = "failure"
    print(f"{server_test}")
    
    # Create Log event dataframe
    ping_data = {'datetime' : [date_time], 'result' : [server_test], 'avg_reply_time' : [avg_reply_time]}
    ping_event = pd.DataFrame(ping_data)

    #insert row in dataframe
    insert_db(ping_event, engine, "ping_event_log")

    # Send email to self and relevent individuals
    email_dict = {
        'jcohen@curtispackaging.com' : f'''<body>
                            <h3> Curtis-HV1 is not Responding. </h3>
                            <p> The Applications are being back-up deployed on {local_ip}. </p>
                            <p> An email is being sent to the Tablets informing users of the downage and instructing them to redirect to 10.1.1.187 </p>
                        <body>                         
                            ''', 
        'Quality@curtispackaging.com' : f'''<body>
                            <h3> The server hosting the Final Inspection Application is down. </h3>
                            <p> The Applications are being back-up deployed on {local_ip}. </p>
                            <p> 
                                <a href="http://{local_ip}:8502/">Follow this link to be redirected to the Final Inspection App</a>
                             </p>
                        <body>                         
                            '''
    }

    # new mail object must be created for each email, hence iteration through email_dict
    for recipient, message in email_dict.items():
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = 'Curtis-HV1 Not Responding'
        mail.HTMLBody = message
        mail.Send()

    # Auto-deploy apps
    p = Popen(r"C:\Users\jcohen\Documents\ping_email_script\streamlitrun.bat")
    stdout, stderr = p.communicate()

