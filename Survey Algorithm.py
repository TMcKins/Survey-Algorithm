import imaplib
import email
from email.header import decode_header
import webbrowser
import os
import pandas as pd
import numpy as np
import csv, smtplib, ssl

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr

import io

#import yagmail


print("==================================================================================")
print("==================================================================================")



def house_elec_file_matcher(i):
    #match the occupants column value to the file for that household usage
    # so if theyve selected 1 person, it retrieves 'oneperson.csv'
    answer = 0

    #files = ['OnePerson.csv', 'TwoPerson.csv', 'ThreePerson.csv', 'FourPerson.csv', 'FivePerson.csv', 'Six+Person.csv']

    xls = pd.ExcelFile('Household Electricity Consumption.xls')
    df1 = pd.read_excel(xls, 'One Person')
    df2 = pd.read_excel(xls, 'Two Person')
    df3 = pd.read_excel(xls, 'Three Person')
    df4 = pd.read_excel(xls, 'Four Person')
    df5 = pd.read_excel(xls, 'Five Person')
    df6 = pd.read_excel(xls, 'Six + Person')

    if i == '1':
        answer = df1
    elif i == '2':
        answer = df2
    elif i == '3':
        answer = df3
    elif i == '4':
        answer = df4
    elif i == '5':
        answer = df5
    elif i == '6':
        answer = df6

    return answer




print("==================================================================================")
print("==================================================================================")






#account credentials
username = "datacollection203@gmail.com"
password = "SheffieldAlgorithm"

def clean(text):
    #clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)

#create an IMAP4 class with SSL
imap = imaplib.IMAP4_SSL("imap.gmail.com")

#authenticate
imap.login(username, password)

status, messages = imap.select("INBOX")

N = 1

messages = int(messages[0])

data = []

for i in range(messages, messages-N, -1):
    # fetch the email message by ID
    res, msg = imap.fetch(str(i), "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            # parse a bytes email into a message object
            msg = email.message_from_bytes(response[1])
            # decode the email subject
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                # if it's a bytes, decode to str
                subject = subject.decode(encoding)
            # decode email sender
            From, encoding = decode_header(msg.get("From"))[0]
            if isinstance(From, bytes):
                From = From.decode(encoding)
            print("Subject:", subject)
            print("From:", From)
            # if the email message is multipart
            if msg.is_multipart():
                # iterate over email parts
                for part in msg.walk():
                    # extract content type of email
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition"))
                    print("NEED THIS PART 1 ")
                    try:
                        # get the email body
                        body = part.get_payload(decode=True).decode()
                        print("NEED THIS PART 2")
                    except:
                        pass
                        print("NEED THIS PART 3")
                    if content_type == "text/plain" and "attachment" not in content_disposition:
                        # print text/plain emails and skip attachments
                        print(body) #THIS IS THE EMAIL!!!!!
                        string = body
                        data.append(string)
                        print("NEED THIS PART 4")
                    elif "attachment" in content_disposition:
                        # download attachment
                        filename = part.get_filename()
                        if filename:
                            folder_name = clean(subject)
                            print("NEED THIS PART 5")
                            if not os.path.isdir(folder_name):
                                # make a folder for this email (named after the subject)
                                os.mkdir(folder_name)
                            filepath = os.path.join(folder_name, filename)
                            # download attachment and save it
                            open(filepath, "wb").write(part.get_payload(decode=True))
                            print("NEED THIS PART 6")
            else:
                # extract content type of email
                content_type = msg.get_content_type()
                # get the email body
                body = msg.get_payload(decode=True).decode()
                data = data.append(body)
                print("NEED THIS PART 7")
                if content_type == "text/plain":
                    # print only text email parts
                    print(body)
                    print("NEED THIS PART 8")
            if content_type == "text/html":
                # if it's HTML, create a new HTML file and open it in browser
                folder_name = clean(subject)
                print("NEED THIS PART 9")
                if not os.path.isdir(folder_name):
                    # make a folder for this email (named after the subject)
                    os.mkdir(folder_name)
                    print("NEED THIS PART 10")
                filename = "index.html"
                filepath = os.path.join(folder_name, filename)
                # write the file
                open(filepath, "w").write(body)
                # open in the default browser
                webbrowser.open(filepath)
                print("NEED THIS PART 11")
            print("="*100)
# close the connection and logout

imap.close()
imap.logout()



print("==================================================================================")
print("==================================================================================")


print(string)
print(type(string))

email_body = string.split()

#df = pd.DataFrame(data, columns=['string_values'])
#dt = df.string_values.str.split(expand = True,)
#cols = [0,1,2,3,4,5,6,7,8,9,10,11,13,14,16,17,18,19,20,21,22,24,25,26,27,28,29]
#dp = dt.drop(dt.columns[cols], axis = 1)
#dp.columns = ['name', 'email', 'occupants']
#print(dt)
#print(type(df))
#Responses = dp

#Responses.to_csv(r'F:\PhD\Data Collection\Google Forms\responses.csv', index=False)

#print(email_body)







print("==================================================================================")
# EXTRACTING EMAIL RESULTS INTO USABLE DATAFRAME - CONSISTENT FOR ALL RESPONSES TO BE USED BY ALGORITHM

#Making Dataframes to be used by Algorithm
column_names_consent = ['1.1', '1.2', '1.3', '1.4', '1.5', '1.6', '1.7', '1.8']
Consent = pd.DataFrame(columns = column_names_consent)

column_names_demographic = ['Name', 'Email Address', 'Local Area']
Demographic = pd.DataFrame(columns = column_names_demographic)

column_names_social = ['Aware of Transition', 'Currently own EV', 'Next Car EV?', 'Reasons For', 'Reasons Against']
Social = pd.DataFrame(columns = column_names_social)

column_names_numcars = ['How Many Cars']
Numcars = pd.DataFrame(columns = column_names_numcars)

column_names_carinfo = ['Public Transport', 'Fuel Type Car 1', 'Type of Vehicle Car 1', 'Replacement Type Car 1', 'Trip Purposes Car 1', 'Fuel Type Car 2', 'Type of Vehicle Car 2', 'Replacement Type Car 2', 'Trip Purposes Car 2', 'Fuel Type Car 3', 'Type of Vehicle Car 3', 'Replacement Type Car 3', 'Trip Purposes Car 3', 'Fuel Type Car 4', 'Type of Vehicle Car 4', 'Replacement Type Car 4', 'Trip Purposes Car 4', 'Fuel Type Car 5', 'Type of Vehicle Car 5', 'Replacement Type Car 5', 'Trip Purposes Car 5']
Carinfo = pd.DataFrame(columns = column_names_carinfo)

column_names_car1specific = ['Leave for Work', 'Return from Work', 'Work miles', 'Working Days', 'Leave for School', 'Return from School', 'Car Remain at School?', 'School Miles', 'School Days', 'Leave for Shopping', 'Return from Shopping', 'Shopping Miles', 'Shopping Days', 'Leave for Day Trip', 'Return from Day Trip', 'Day Trip Miles', 'Day Trip Days', 'Leave for Other', 'Return from Other', 'Other Miles', 'Other Days']
Car1 = pd.DataFrame(columns = column_names_car1specific)

column_names_car2specific = ['Leave for Work', 'Return from Work', 'Work miles', 'Working Days', 'Leave for School', 'Return from School', 'Car Remain at School?', 'School Miles', 'School Days', 'Leave for Shopping', 'Return from Shopping', 'Shopping Miles', 'Shopping Days', 'Leave for Day Trip', 'Return from Day Trip', 'Day Trip Miles', 'Day Trip Days', 'Leave for Other', 'Return from Other', 'Other Miles', 'Other Days']
Car2 = pd.DataFrame(columns = column_names_car2specific)

column_names_car3specific = ['Leave for Work', 'Return from Work', 'Work miles', 'Working Days', 'Leave for School', 'Return from School', 'Car Remain at School?', 'School Miles', 'School Days', 'Leave for Shopping', 'Return from Shopping', 'Shopping Miles', 'Shopping Days', 'Leave for Day Trip', 'Return from Day Trip', 'Day Trip Miles', 'Day Trip Days', 'Leave for Other', 'Return from Other', 'Other Miles', 'Other Days']
Car3 = pd.DataFrame(columns = column_names_car3specific)

column_names_car4specific = ['Leave for Work', 'Return from Work', 'Work miles', 'Working Days', 'Leave for School', 'Return from School', 'Car Remain at School?', 'School Miles', 'School Days', 'Leave for Shopping', 'Return from Shopping', 'Shopping Miles', 'Shopping Days', 'Leave for Day Trip', 'Return from Day Trip', 'Day Trip Miles', 'Day Trip Days', 'Leave for Other', 'Return from Other', 'Other Miles', 'Other Days']
Car4 = pd.DataFrame(columns = column_names_car4specific)

column_names_car5specific = ['Leave for Work', 'Return from Work', 'Work miles', 'Working Days', 'Leave for School', 'Return from School', 'Car Remain at School?', 'School Miles', 'School Days', 'Leave for Shopping', 'Return from Shopping', 'Shopping Miles', 'Shopping Days', 'Leave for Day Trip', 'Return from Day Trip', 'Day Trip Miles', 'Day Trip Days', 'Leave for Other', 'Return from Other', 'Other Miles', 'Other Days']
Car5 = pd.DataFrame(columns = column_names_car5specific)

column_names_totalcars = ['Answer to 6.1']
Totalcars = pd.DataFrame(columns = column_names_totalcars)

column_names_householdelec = ['Number of Occupants', 'Annual Usage']
Householdelec = pd.DataFrame(columns = column_names_householdelec)

column_names_electariffs = ['Monitoring', 'Type of Meter', 'Different Meter?', 'Aware of EV tariffs']
Electariffs = pd.DataFrame(columns = column_names_electariffs)

column_names_charging = ['Parking Facilities', 'Number of Charge Points', 'Public Charging', 'Type of Charger', 'Aware of Bat Degrade', 'Plugin Whenever', 'Change Habits for Bat Degrade?', 'Change Habits for Costs?']
Charging = pd.DataFrame(columns = column_names_charging)





print("==================================================================================")
#POPULATING THE ABOVE DATAFRAMES WITH ANSWERS FROM QUESTIONNAIRE EMAIL
#Dataframe method
df = pd.DataFrame(data, columns=['string_values'])
dt = df.string_values.str.split(expand = True,)
dp = dt.transpose() #make it a column vector
print(dp)


#CONSENT DATAFRAME








#DEMOGRAPHIC DATAFRAME
#Get 'NAME' of participant
x = dp[dp[0].str.contains(r'\bName\b')] #\b either side of has so that it finds only has and not other words with the letters of has in
y = dp.index[dp[0].str.contains(r'\bName\b')].item() #gets the index number of that word in the dataframe column
t = y+1
p = dp.iloc[t]
print(p)
#Adding the data from participants responses to final dataframe to be used by Algorithm
Demographic['name'] = dp.iloc[t]

#Get 'EMAIL' of participant
y1 = dp[dp[0].str.contains(r'\bAddress\b')]
print(y1)
y2 = dp.index[dp[0].str.contains(r'\bAddress\b')].item()
print(y2) #gets the index number of that word in the dataframe column
y3 = y2+1
y4 = dp.iloc[y3]
print(y4)
Demographic['Email Address'] = dp.iloc[y3]


#Get 'Local Area' of participant
#Demographic['Local Area'] =











print("==================================================================================")
fghfds = fgnhj*kk6
print(gg)




#SOCIAL DATAFRAME



#NUMCARS DATAFRAME



#CARINFO DATAFRAME



#CAR1 DATAFRAME
#Car 1 - Usage Information
x1 = dp.index[dp[0].str.contains(r'\bCar\b')].item()
x2 = dp.index[dp[0].str.contains(r'\b1\b')]

if isinstance(x1, int):
    x1 = [x1]

if isinstance(x2, int):
    x2 = [x2]

if not isinstance(x1, list):
    x1 = x1.tolist()

if not isinstance(x2, list):
    x2 = x2.tolist()

print(type(x1))
print(type(x2))

print(x1)
print(x2)

#GO TO WORK TIME
for i in range(len(x1)):
    if x1[i]+1 in x2:
        print('true')
        print(type(x1[i]))
        x2index = x2.index(x1[i]+1)
        print(x2.index(x1[i]+1))
        t = x1[i]+13
        p = dp.iloc[t]
        print(p)

#responses['area'] = p



#Car 1 - HOME FROM WORK TIME
for i in range(len(x1)):
    if x1[i]+1 in x2:
        print('true')
        print(type(x1[i]))
        x2index = x2.index(x1[i]+1)
        print(x2.index(x1[i]+1))
        t = x1[i]+22
        p = dp.iloc[t]
        print(p)

#responses['occupants'] = p







#CAR2 DATAFRAME



#CAR3 DATAFRAME



#CAR4 DATAFRAME



#CAR5 DATAFRAME


















print("==================================================================================")
#ALGORITHM




fghfds = fgnhj*kk6
print(gg)





#Now code below will pic a picture of the graph of the correct household
#occupancy electricity consumption - working example of the algorithm

#basically just need to make the files dataframe with the correct filenames in each row
#based on the responses in the dataframe 'Responses'



# files = ['OnePerson.csv', 'TwoPerson.csv', 'ThreePerson.csv', 'FourPerson.csv', 'FivePerson.csv', 'Six+Person.csv']
#
# xls = pd.ExcelFile('Household Electricity Consumption.xls')
# df1 = pd.read_excel(xls, 'One Person')
# print(df1)
# df2 = pd.read_excel(xls, 'Two Person')
# df3 = pd.read_excel(xls, 'Three Person')
# df4 = pd.read_excel(xls, 'Four Person')
# df5 = pd.read_excel(xls, 'Five Person')
# df6 = pd.read_excel(xls, 'Six + Person')
#
#
# for i in Responses['occupants']:
#     print(i)
#     print(type(i))
#     if i == '1':
#         df1.to_csv(r'F:\PhD\Data Collection\Google Forms\OnePerson.csv', index=False)
#         print('This person selected 1 Person')
#         print('Saved OnePerson')
#     elif i == '2':
#         df2.to_csv(r'F:\PhD\Data Collection\Google Forms\TwoPerson.csv', index=False)
#     elif i == '3':
#         df3.to_csv(r'F:\PhD\Data Collection\Google Forms\ThreePerson.csv', index=False)
#     elif i == '4':
#         df4.to_csv(r'F:\PhD\Data Collection\Google Forms\FourPerson.csv', index=False)
#     elif i == '5':
#         df5.to_csv(r'F:\PhD\Data Collection\Google Forms\FivePerson.csv', index=False)
#     elif i == '6':
#         df6.to_csv(r'F:\PhD\Data Collection\Google Forms\Six+Person.csv', index=False)
#
#
#
#
#




print("==================================================================================")
print("==================================================================================")

# SENDING EMAIL BACK TO PARTICIPANT

# User configuration
sender_email = username
sender_name = 'Thomas McKinney'

receiver_emails = Responses['email']
receiver_names = Responses['name']
occupants = Responses['occupants']

# Email body
email_html = open('email.html')
email_body = email_html.read()

#files = ['responses.csv', 'responses.csv', 'responses.csv']

for receiver_email, receiver_name, occupant in zip(receiver_emails, receiver_names, occupants):
        print("Sending the email...")
        # Configurating user's info
        msg = MIMEMultipart()
        msg['To'] = formataddr((receiver_name, receiver_email))
        msg['From'] = formataddr((sender_name, sender_email))
        msg['Subject'] = 'Your Google Form Feedback'

        msg.attach(MIMEText(email_body, 'html'))

        try:
            pt = house_elec_file_matcher(occupant)
            pt.to_csv(r'F:\PhD\Data Collection\Google Forms\attachment.csv', index=False)
            filename = 'attachment.csv'
            # Open PDF file in binary mode
            with open(filename, "rb") as attachment:
                            part = MIMEBase("application", "octet-stream")
                            #print(filename)
                            part.set_payload(attachment.read())

            # Encode file in ASCII characters to send by email
            encoders.encode_base64(part)
            #print(part)

            # Add header as key/value pair to attachment part
            part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {filename}",
            )

            msg.attach(part)
        except Exception as e:
                print(f'Oh no! We didnt found the attachment!n{e}')
                break

        try:
                # Creating a SMTP session | use 587 with TLS, 465 SSL and 25
                server = smtplib.SMTP('smtp.gmail.com', 587)
                # Encrypts the email
                context = ssl.create_default_context()
                server.starttls(context=context)
                # We log in into our Google account
                server.login(sender_email, password)
                # Sending email from sender, to receiver with the email body
                server.sendmail(sender_email, receiver_email, msg.as_string())
                print('Email sent!')
        except Exception as e:
                print(f'Oh no! Something bad happened!n{e}')
                break
        finally:
                print('Closing the server...')
                server.quit()
