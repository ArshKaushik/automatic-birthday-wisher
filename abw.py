import pandas as pd 
import datetime
import smtplib # it is a library which will send email with the help of SMTP
import os

os.chdir(r"C:\Users\kaush\Documents\VSC Python\automaticBirthdayWisher") # this code will help you to run the program after getting into this directory
# os.mkdir("testing") # just to test where program is running or not via windows task scheduler

# Enter your authentication details to send the emails
GMAIL_ID = ''
GMAIL_PSWD = ''

# We need to turn ON less secure apps option for our gmail account to run this program
# In future, for serious security concerns, we can also buy gmail API (gsuite) or any other mail APIs (Mail Chimp)

def sendWish(name, toEmail, sub, message):   
    # SMTP session below:
    # we can know more about this by googling 'smtp for gmail docs'
    s = smtplib.SMTP('smtp.gmail.com', 587) # '587' is a port
    s.starttls() # starting the session
    s.login(GMAIL_ID, GMAIL_PSWD)

    s.sendmail(GMAIL_ID, toEmail, f"Subject: {sub}\n\n{message}")
    print(f"Email has been sent to {name} at {toEmail} with subject as {sub} & the message is {message}")
    s.quit()

# We can also send an SMS by using SMS APIs    

if __name__ == "__main__":
    dataFrame = pd.read_excel("data.xlsx")
    # print(dataFrame)
    today = datetime.datetime.now().strftime("%d-%m") # strftime() will give us our desired date format. To include year, use '%Y'
    yearNow = datetime.datetime.now().strftime("%Y")
    # print(today) # 'today' will be in string format

    writeIndex = [] # for updation of year
    # To iterate through dataFrame (excel sheet)
    for index, item in dataFrame.iterrows(): # iterrows() will give the index and item of dataFrame
        # print(index, item["Birthday"]) # will show indexes and their birthday dates
        bday = item["Birthday"].strftime("%d-%m") # because we don't require the year
        # print(bday)
        subject = "Happy Birthday"
        if today == bday and yearNow not in str(item["Year"]):
            sendWish(item["Name"], item["Email"], subject, item["Dialogue"]) # we can create subject column in our data sheet & pull that here
            writeIndex.append(index) # updating the year to stop sending multiple wishes, only send once


# TODO1: Optimization: if writeIndex is empty, then we don't need to run further steps written below

    # print(writeIndex) # will print the index        
    for i in writeIndex:
        year = dataFrame.loc[i, "Year"] # loc[] will be using the i & Year where 'i' = row indexer & "Year" = column
        # print(year)
        dataFrame.loc[i, "Year"] = str(year) + ', ' + str(yearNow) # will add current year to the previous year in dataFrame
        # print(dataFrame.loc[i, "Year"])

    # print(dataFrame)    
    dataFrame.to_excel('data.xlsx', index = False) # to save all the changes made via codes into the excel file
                                                   # & index is set to False so that no additional indexing coloum should form in excel sheet



