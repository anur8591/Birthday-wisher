import pandas as pd 
import datetime
import smtplib as smp
import os 

current_path = os.getcwd()
print(current_path)

os.chdir(current_path)

email_ID = input("enter the E-mail ID: ")
email_Pass = input("enter the password of E-mail: ")


def send_Email (to, sub, msg):
    print(f"Email to: {to} sent: \nSubject: {sub}, \nMassage: {msg}")
    s = smp.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(email_ID, email_Pass)
    s.sendmail(email_ID, to, f"subject: {sub} \n\n {msg}")
    s.quit()

if __name__ == "__main__":
    df = pd.read_excel('C:\Anurag_WorkSpace\study\python\Birthday wisher\Book1.xlsx')
    today = datetime.datetime.now().strftime("%d-%m")  
    yearNow = datetime.datetime.now().strftime("%Y")

    writeInd = []
    for index, item in df.iterrows():
        bday = item['Birthday']
        bday_str = bday.strftime("%d-%m")
        if(today == bday_str) and yearNow not in str(item['LastWishedYear']):
            send_Email(item['Email'], "Happy Birthday", item['Dialogue'])
            writeInd.append(index)

    if writeInd != None:
        for i in writeInd:
            oldYear = df.loc[i, 'LastWishedYear']
            df.loc[i, 'LastWishedYear'] = str(oldYear) + ", " + str(yearNow)

    df.to_excel('C:\\Anurag_WorkSpace\\study\\python\\Birthday wisher\\Book1.xlsx', index = False)    






# The error means that Gmail needs a special password called an "App Password" to let your program send emails. 
# This happens if you have extra security (like 2-step verification) turned on for your Gmail account. 
# You can't use your regular Gmail password in the program. 
# Instead, you need to create an App Password in your Google account settings and use that in your program to 
# log in and send emails successfully.