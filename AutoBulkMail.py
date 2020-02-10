import pandas as pd
import smtplib
from email.message import EmailMessage
import imghdr


file = pd.ExcelFile("YourExcelFile.xlsx")  # ADD your Excel file ******************************

count = 1  # SrNO

# Email Setup
s = smtplib.SMTP("smtp.gmail.com", 587)
s.starttls()  # Traffic encryption
s.login("YourEmail@gmail.com", "YourPassword")  # ADD your login Credentials ******************************

for sheet in file.sheet_names:
    print("\n\n New Sheet...\n")
    count = 2
    df1 = file.parse(sheet)
    for i in range(len(df1['EMAIL'])):

        count += 1

        msg = EmailMessage()
        msg['Subject'] = 'Invitation from Team Abhigyan.'
        msg['From'] = "YourEmail@gmail.com"
        msg['To'] = df1['EMAIL'][i]
        msg.set_content("""
Dear """ + str(df1['NAME'][i]) + " :" + """

   Hello and Namaste from Team Walk With World!
   The Team Walk With World is back
   with the bang on event of Abhigyan 2020. 
   We are ready to host you in this 
   enthusiastic arena.

   Then, what are you waiting for?
   With a huge response, 
   we have already closed student entries	
   and now, few professional entries are left.

   Please find below information about 
   Abhigyan 2020
   Guests: 
          \t Ranveer Allahbadia(Beer Biceps),
	      \t Nikhil Sharma(MumbikarNikhil), 
		  Dr. Shital Amte-karajagi, 
		  Achyut Godbole, 
		  Indraneel Chitale 
		  and a surprise guest!

   Venue : Victoria Banquet, Hotel Sayaji, Kolhapur.
   Date : Sunday, 1st March, 2020
   Entry Fee : Professional - 600/-

   Also, find the attached poster of Abhigyan 2020
   for more details.
	
   See you at Abhigyan 2020!
	 
        """)

        #Attaching the Poster
        f = open("images/Abgyn2020.jpg", 'rb')  # Attached my Event's Poster for the mail **********
        file_data = f.read()
        file_name = "ABHIGYAN2020 : Poster"
        file_type = imghdr.what(f.name)
        msg.add_attachment(file_data, maintype='image', subtype=file_type, filename=file_name)

        s.send_message(msg)
        print("--> ", count, ": ", df1['EMAIL'][i], " : Sent")
s.quit()
print("\n Emails sent...")
