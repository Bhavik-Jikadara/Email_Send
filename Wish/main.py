import pandas as pd
import datetime
import smtplib

GMAIL_ID = "____"  # Enter your email 
GMAIL_PASSWORD = "___"  # Enter your email password Before this run Program 
def sendEmail(to,sub,msg):
    print ("Successfully Send Email !!!!!!!!!!")
    s= smtplib.SMTP("smtp.gmail.com", 587)
    s.starttls()
    s.login(GMAIL_ID,GMAIL_PASSWORD)

    s.sendmail(GMAIL_ID,to, f"Subject :{sub}\n\n{msg}")
    s.quit()
    
if __name__ == "__main__":
    sendEmail(GMAIL_ID, "subject","test message")
    df = pd.read_excel("Data.xlsx")
    today = datetime.datetime.now().strftime("%d-%m")
    yearnow = datetime.datetime.now().strftime("%Y")
    
    writels = []
    for index,item in df.iterrows():
        bday = item["Brithday"].strftime("%d-%m")

        if (today == bday) and yearnow not in  str(item["years"]):
            sendEmail(item["email"],"Happy Brithday",item["Message"])
            writels.append(index)

    # print(writels)
    for i in writels:
        yr = df.loc[i,"years"]
        df.loc[i,"years"] = str(yr) + "," + str(yearnow)
        # print(df.loc[i,"years"])

    df.to_excel("Data.xlsx", index = False)

