import win32com.client as client
import pathlib
import pandas as pd
import os.path

file = 'Reports/Services.xlsx'
outlook = client.Dispatch('Outlook.Application')
with pd.ExcelFile(file) as xls:
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        for x in range(0,len(df)):
            message = outlook.CreateItem(0)
            message.To = str(df.iloc[x,2])
            message.Subject = 'Reports from Gestion Temps'
           # message.Body = "Hello, The reports for last mounth are attached in this mail."
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
            message.Display()
            for col in range(3,len(df.columns)):
                if str(df.iloc[x,col]) != "nan":
                    file_path = pathlib.Path(str(df.iloc[x,col]) + '.xlsx')
                    file_absolute = str(file_path.absolute())
                    if os.path.exists(file_absolute):
                        message.Attachments.Add(file_absolute)
                        print(df.iloc[x,col])
                        print(col)
                    else:
                        print(str(file_path) + "  Dose not exist!!")
                elif str(df.iloc[x,col]) == "nan":
                    print("################")
                    print(col)
                   # message.Save()
                   # message.Send()
                    break

                



#len(df.columns)
#outlook = client.Dispatch('Outlook.Application')
#message = outlook.CreateItem(0)

#message.To = 'Csaba.Imreh@mail.nidec.com'

#message.Subject = 'Test Auto mail'
#message.Body = "Acest mail este un test."

#file_path = pathlib.Path('TaskCoach User Update LOG  All Locations.xlsx')
#file_absolute = str(file_path.absolute())
#message.Attachments.Add(file_absolute)
#message.SendOnBehalfOfName = "SVP_TaskCoach_Improvments.IALS@mail.nidec.com"
#message.Save()
#message.Send()
