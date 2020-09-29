"""
------- Auto mail sender of Reports from Gestion Temps -------
    
    This program send mails to a list of Gestion Temps 
    project managers with reports generated from Gestion
    Temps.The list of Users and the associated Services
    are taken from the Reports directory.
    
"""

### Imports ###
import win32com.client as client
import pathlib
import pandas as pd
import os.path

file = 'Reports/Services.xlsx'  # Files path for the list of users and Services
outlook = client.Dispatch('Outlook.Application')  # Open Outlook

### Read XLS file line by line ###
with pd.ExcelFile(file) as xls:
    for sheet_name in xls.sheet_names:
        # select the sheet from which read the mails and Services
        df = pd.read_excel(xls, sheet_name=sheet_name)
        for x in range(0, len(df)):                    # reading the lines
            print("----------New Mail----------")
            message = outlook.CreateItem(0)            # creating a new email
            # the reciver email is set
            message.To = str(df.iloc[x, 2])
            message.Subject = 'Reports from Gestion Temps'

### Email message in html format ###
            html_body = """
    <div>
        <h1 style="font-family: 'Lucida Sans'; font-size: 35; font-weight: bold; color: #9eac9c;"> Reports from Gestion Temps! </h1>
        <span style="font-family: 'Lucida Sans'; font-size: 28; color: #8d395c;"> You can find the reports attached to this mail! </span>
    </div><br>
    <p>Best regards,</p><br>
    <p>Imreh Csaba│Junior Software Test and Qualification Engineer </p>
    <p>Leroy Somer - Nidec Oradea SRL </p>
    <p>20 Petre P. Carp street | Oradea | 410603 | Romania</p>
    <p>Csaba.Imreh@mail.nidec.com  | www.nidecautomation.com</p>
    """
            message.HTMLBody = html_body
            # message.Display()
            attachedFileCounter = 0

### Attach Files and Sending Email ###
            for col in range(3, len(df.columns)):

                if str(df.iloc[x, col]) != "nan":
                    file_path = pathlib.Path(str(df.iloc[x, col]) + '.xlsx')
                    file_absolute = str(file_path.absolute())

                    if os.path.exists(file_absolute):
                        attachedFileCounter = attachedFileCounter + 1
                        message.Attachments.Add(file_absolute)
                        print(str(file_path) + "  Attached Successful")

                    else:
                        print(str(file_path) + "  Dose not exist!!")

                elif str(df.iloc[x, col]) == "nan":
                    break

            if attachedFileCounter > 0:
                print("Email sent to {}".format(str(df.iloc[x, 2])))
                message.Send()
                print("----------End Mail----------\n\n")

            else:
                print("No files attached mail not sent to {}".format(
                    str(df.iloc[x, 2])))
                print("----------End Mail----------\n\n")
