import os

import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import messagebox

import smtplib
from email import encoders
from email.message import Message
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from smtplib import SMTPException

root = tk.Tk()
root.withdraw()
cwd = os.getcwd()
folder = os.path.join(cwd, "Windows10")

if not os.path.exists(folder):
    messagebox.showerror("Warning", '"Windows10" folder does not exist!')

else:
    global directory
    unsupported_path = os.chdir(folder)
    new_path = os.getcwd()
    directory = new_path


def removefiles(dir_path):
    """Function that will delete all the unnecessary files inside the Windows10 directory"""
    versions = ['_f01_', '_g01_', '_0.csv',
                '_LTSB_', '_LTSC_', '_Other_Versions_']

    for filename in os.listdir(dir_path):
        for i in versions:
            try:
                if filename .__contains__(i):
                    os.remove(os.path.join(dir_path, filename))
                else:
                    continue
            except FileNotFoundError:
                pass

###Function to Count Final Number of Rows###


def countItems(dir_path):

    count_windows = os.listdir(dir_path)
    new_count_windows = []

    for file_win in count_windows:
        try:
            df2 = pd.read_csv(file_win, encoding='utf-8')
            # Final count of Total Number of Rows
            n3_rows = str(len(df2.axes[0]))

            ### Final count Total Number of Rows ###
            #Creating new filename
            x = file_win.split("_")
            y1 = x.pop()
            y2 = n3_rows+".csv"
            y3 = x.append(y2)
            z = "_".join(x)
            csv2_filename = z

            #Saving the Filtered Data Frame into the New CSV file
            csv2_data = df2.to_csv(
                csv2_filename, index=False, header=True, encoding='utf-8-sig')
            new_count_windows.append(csv2_filename)

            #Delete the Original CSV file
            if file_win not in new_count_windows:
                os.remove(file_win)

        except KeyError:
            pass
        except ValueError:
            pass
        except FileNotFoundError:
            pass

###Function that will create G09 Email template with Attached Files###


def send_email_g09(msg_from, msg_to):

    ###Create message container - the correct MIME type is multipart/alternative###
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "[Request] Upgrading Unsupported Windows 10 (Data as of MONTH DATE YEAR)"
    msg['From'] = msg_from
    msg['To'] = msg_to

    ###Create the body of the message -- HTML###

    #************G09 Admin ************#
    html = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
    table {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    border-collapse: collapse;
    text-align: middle;
    width: 55%;
    style: "text-align:center"
    }

    td {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    text-align: middle;
    style: "text-align:center";
    }

    th {
    border: 1px solid black;
    text-align: middle;
    padding: 1px;
    background-color:#ccccff;
    }

    </style>
    </head>

    <body>
    <p style="font-family:calibri;font-size:13px">
    Dear G09 Admins,

        <br><br>This is <strong>(Your name here)</strong> on behalf of CIT SOC Tier 1 and CIT SOC Tier 2.</br>
        --------------------------------------------------------------------------------------------------------------------------</br>
        
        We are sending you this notification according to our schedule of semimonthly updates with the latest data generated on 
        <strong>(MONTH DATE YEAR)</strong>.</br>
        <br>If you’ve already been contacting us to provide with your upgrading schedule relating to this issue, please just keep the file(s) 
        for your reference. Otherwise, please read the following and make corresponding action at your earliest convenience.</br>

        --------------------------------------------------------------------------------------------------------------------------</br>
        Please see the attached file(s) of computer(s) in your domain using Windows 10 Versions of which its support had already been expired.
        The dates of support termination for each Windows 10 version are listed in the following table.</br></p>

        <p style="font-family:calibri;font-size:13px;color:blue">
        *For more information about Windows 10 lifecycles, please refer to below URL:
        <br> <a href="https://docs.microsoft.com/en-us/lifecycle/faq/windows">
        https://support.microsoft.com/en-us/help/13853/windows-lifecycle-fact-sheet</a></p>

        <p style="font-family:calibri;font-size:13px">
        Due to security concerns, we recommend you need to upgrade them to higher version as soon as possible.

    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        <table>
    <tr>
        <th>Windows 10 version</th>
        <th>End of Support</th>
        <th># of Objects</th>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1607 (14393) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">April 9, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1703 (15063) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">October 8, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1703 (15063) <br>
        Pro Edition </br></td>
        <td style="text-align:center">October 9, 2018</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1709 (16299) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">October 13, 2018</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1709 (16299) <br>
        Pro Edition </br></td>
        <td style="text-align:center">April 9, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1803 (17134) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1809 (17763) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1909 (18363) <br>
        Pro Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>  
    <tr>
        <td style="text-align:center">Windows 10 Pro 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>    
    <tr>
        <td style="text-align:center">Windows 10 Enterprise 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>   
        </table>


    <p style="font-family:calibri;font-size:13px">
        If you have any questions or concerns regarding this issue, please do not hesitate to contact us.
    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        
    <br><strong>(Your Signature Here)</strong></br>   


    </body>
    </html>

    """
    ###--------End of HTML body message--------###

    ###Attaching the CSV files to the Email###

    filelist = os.listdir(directory)
    g09 = []
    for csv_file in filelist:
        if csv_file .__contains__("_g09_"):
            g09.append(csv_file)
        else:
            continue

    for file_name in g09:
        with open(file_name, 'rb') as filetosend:
            record = MIMEBase('application', 'octet-stream')
            record.set_payload(filetosend.read())
            encoders.encode_base64(record)
            record.add_header('Content-Disposition', 'attachment',
                              filename=os.path.basename(file_name))
            msg.attach(record)

            part = MIMEText(html, 'html')
            msg.attach(part)

    if g09 != []:
        try:
            server = smtplib.SMTP('webmail.g07.fujitsu.local')
            server.sendmail(msg_from, msg_to, msg.as_string())
        except SMTPException:
            messagebox.showerror("Error", "Error: unable to send G09 email")
        server.quit()
    else:
        pass

###Function that will create G08 Email template with Attached Files###


def send_email_g08(msg_from, msg_to):

    ###Create message container - the correct MIME type is multipart/alternative###
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "[Request] Upgrading Unsupported Windows 10 (Data as of MONTH DATE YEAR)"
    msg['From'] = msg_from
    msg['To'] = msg_to

    ###Create the body of the message -- HTML###

    #************G08 Admin ************#
    html = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
    table {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    border-collapse: collapse;
    text-align: middle;
    width: 55%;
    style: "text-align:center"
    }

    td {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    text-align: middle;
    style: "text-align:center";
    }

    th {
    border: 1px solid black;
    text-align: middle;
    padding: 1px;
    background-color:#ccccff;
    }

    </style>
    </head>

    <body>
    <p style="font-family:calibri;font-size:13px">
    Dear G08 Admins,

        <br><br>This is <strong>(Your name here)</strong> on behalf of CIT SOC Tier 1 and CIT SOC Tier 2.</br>
        --------------------------------------------------------------------------------------------------------------------------</br>
        
        We are sending you this notification according to our schedule of semimonthly updates with the latest data generated on 
        <strong>(MONTH DATE YEAR)</strong>.</br>
        <br>If you’ve already been contacting us to provide with your upgrading schedule relating to this issue, please just keep the file(s) 
        for your reference. Otherwise, please read the following and make corresponding action at your earliest convenience.</br>

        --------------------------------------------------------------------------------------------------------------------------</br>
        Please see the attached file(s) of computer(s) in your domain using Windows 10 Versions of which its support had already been expired.
        The dates of support termination for each Windows 10 version are listed in the following table.</br></p>

        <p style="font-family:calibri;font-size:13px;color:blue">
        *For more information about Windows 10 lifecycles, please refer to below URL:
        <br> <a href="https://docs.microsoft.com/en-us/lifecycle/faq/windows">
        https://support.microsoft.com/en-us/help/13853/windows-lifecycle-fact-sheet</a></p>

        <p style="font-family:calibri;font-size:13px">
        Due to security concerns, we recommend you need to upgrade them to higher version as soon as possible.

    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        <table>
    <tr>
        <th>Windows 10 version</th>
        <th>End of Support</th>
        <th># of Objects</th>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1709 (16299) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">October 13, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1709 (16299) <br>
        Pro Edition </br></td>
        <td style="text-align:center">April 9, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr> 
    <tr>
        <td style="text-align:center">Windows 10  1803 (17134) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1803 (17134) <br>
        Pro Edition </br></td>
        <td style="text-align:center">November 12, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr>  
    <tr>
        <td style="text-align:center">Windows 10  1809 (17763) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1809 (17763) <br>
        Pro Edition </br></td>
        <td style="text-align:center">May 12, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1909 (18363) <br>
        Pro Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>  
    <tr>
        <td style="text-align:center">Windows 10 Pro 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>    
    <tr>
        <td style="text-align:center">Windows 10 Enterprise 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
        </table>


    <p style="font-family:calibri;font-size:13px">
        If you have any questions or concerns regarding this issue, please do not hesitate to contact us.
    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
     
    <br><strong>(Your Signature Here)</strong></br>   

    </body>
    </html>

    """

    ###--------End of HTML body message--------###

    ###Attaching the CSV files to the Email###

    filelist = os.listdir(directory)
    g08 = []
    for csv_file in filelist:
        if csv_file .__contains__("_g08_"):
            g08.append(csv_file)
        else:
            continue

    for file_name in g08:
        with open(file_name, 'rb') as filetosend:
            record = MIMEBase('application', 'octet-stream')
            record.set_payload(filetosend.read())
            encoders.encode_base64(record)
            record.add_header('Content-Disposition', 'attachment',
                              filename=os.path.basename(file_name))
            msg.attach(record)

            part = MIMEText(html, 'html')
            msg.attach(part)

    if g08 != []:
        try:
            server = smtplib.SMTP('webmail.g07.fujitsu.local')
            server.sendmail(msg_from, msg_to, msg.as_string())
        except SMTPException:
            messagebox.showerror("Error", "Error: unable to send G08 email")
        server.quit()
    else:
        pass

###Function that will create G07 Email template with Attached Files###


def send_email_g07(msg_from, msg_to):

    ###Create message container - the correct MIME type is multipart/alternative###
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "[Request] Upgrading Unsupported Windows 10 (Data as of MONTH DATE YEAR)"
    msg['From'] = msg_from
    msg['To'] = msg_to

    ###Create the body of the message -- HTML###

    #************G07 Admin ************#
    html = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
    table {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    border-collapse: collapse;
    text-align: middle;
    width: 55%;
    style: "text-align:center"
    }

    td {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    text-align: middle;
    style: "text-align:center";
    }

    th {
    border: 1px solid black;
    text-align: middle;
    padding: 1px;
    background-color:#ccccff;
    }

    </style>
    </head>

    <body>
    <p style="font-family:calibri;font-size:13px">
    Dear G07 Admins,

        <br><br>This is <strong>(Your name here)</strong> on behalf of CIT SOC Tier 1 and CIT SOC Tier 2.</br>
        --------------------------------------------------------------------------------------------------------------------------</br>
        
        We are sending you this notification according to our schedule of semimonthly updates with the latest data generated on 
        <strong>(MONTH DATE YEAR)</strong>.</br>
        <br>If you’ve already been contacting us to provide with your upgrading schedule relating to this issue, please just keep the file(s) 
        for your reference. Otherwise, please read the following and make corresponding action at your earliest convenience.</br>

        --------------------------------------------------------------------------------------------------------------------------</br>
        Please see the attached file(s) of computer(s) in your domain using Windows 10 Versions of which its support had already been expired.
        The dates of support termination for each Windows 10 version are listed in the following table.</br></p>

        <p style="font-family:calibri;font-size:13px;color:blue">
        *For more information about Windows 10 lifecycles, please refer to below URL:
        <br> <a href="https://docs.microsoft.com/en-us/lifecycle/faq/windows">
        https://support.microsoft.com/en-us/help/13853/windows-lifecycle-fact-sheet</a></p>

        <p style="font-family:calibri;font-size:13px">
        Due to security concerns, we recommend you need to upgrade them to higher version as soon as possible.

    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        <table>
    <tr>
        <th>Windows 10 version</th>
        <th>End of Support</th>
        <th># of Objects</th>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1607 Enterprise</br></td>
        <td style="text-align:center">April 9, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1607 (14393) <br>
        Pro Edition </br></td>
        <td style="text-align:center">April 10, 2018</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1703 (15063) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">October 8, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1703 (15063) <br>
        Pro Edition </br></td>
        <td style="text-align:center">October 9, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1709 (16299) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">October 13, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1709 (16299) <br>
        Pro Edition </br></td>
        <td style="text-align:center">April 9, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1803 (17134) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1803 (17134) <br>
        Pro Edition </br></td>
        <td style="text-align:center">November 12, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1809 (17763) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1809 (17763) <br>
        Pro Edition </br></td>
        <td style="text-align:center">May 12, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1909 (18363) <br>
        Pro Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>  
    <tr>
        <td style="text-align:center">Windows 10 Pro 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>    
    <tr>
        <td style="text-align:center">Windows 10 Enterprise 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
        </table>


    <p style="font-family:calibri;font-size:13px">
        If you have any questions or concerns regarding this issue, please do not hesitate to contact us.
    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        
    <br><strong>(Your Signature Here)</strong></br>   


    </body>
    </html>

    """

    ###--------End of HTML body message--------###

    ###Attaching the CSV files to the Email###

    filelist = os.listdir(directory)
    g07 = []
    for csv_file in filelist:
        if csv_file .__contains__("_g07_"):
            g07.append(csv_file)
        else:
            continue

    for file_name in g07:
        with open(file_name, 'rb') as filetosend:
            record = MIMEBase('application', 'octet-stream')
            record.set_payload(filetosend.read())
            encoders.encode_base64(record)
            record.add_header('Content-Disposition', 'attachment',
                              filename=os.path.basename(file_name))
            msg.attach(record)

            part = MIMEText(html, 'html')
            msg.attach(part)

    if g07 != []:
        try:
            server = smtplib.SMTP('webmail.g07.fujitsu.local')
            server.sendmail(msg_from, msg_to, msg.as_string())
        except SMTPException:
            messagebox.showerror("Error", "Error: unable to send G07 email")
        server.quit()
    else:
        pass

###Function that will create G06 Email template with Attached Files###


def send_email_g06(msg_from, msg_to):

    ###Create message container - the correct MIME type is multipart/alternative###
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "[Request] Upgrading Unsupported Windows 10 (Data as of MONTH DATE YEAR)"
    msg['From'] = msg_from
    msg['To'] = msg_to

    ###Create the body of the message -- HTML###

    #************G06 Admin ************#
    html = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
    table {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    border-collapse: collapse;
    text-align: middle;
    width: 55%;
    style: "text-align:center"
    }

    td {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    text-align: middle;
    style: "text-align:center";
    }

    th {
    border: 1px solid black;
    text-align: middle;
    padding: 1px;
    background-color:#ccccff;
    }

    </style>
    </head>

    <body>
    <p style="font-family:calibri;font-size:13px">
    Dear G06 Admins,

        <br><br>This is <strong>(Your name here)</strong> on behalf of CIT SOC Tier 1 and CIT SOC Tier 2.</br>
        --------------------------------------------------------------------------------------------------------------------------</br>
        
        We are sending you this notification according to our schedule of semimonthly updates with the latest data generated on 
        <strong>(MONTH DATE YEAR)</strong>.</br>
        <br>If you’ve already been contacting us to provide with your upgrading schedule relating to this issue, please just keep the file(s) 
        for your reference. Otherwise, please read the following and make corresponding action at your earliest convenience.</br>

        --------------------------------------------------------------------------------------------------------------------------</br>
        Please see the attached file(s) of computer(s) in your domain using Windows 10 Versions of which its support had already been expired.
        The dates of support termination for each Windows 10 version are listed in the following table.</br></p>

        <p style="font-family:calibri;font-size:13px;color:blue">
        *For more information about Windows 10 lifecycles, please refer to below URL:
        <br> <a href="https://docs.microsoft.com/en-us/lifecycle/faq/windows">
        https://support.microsoft.com/en-us/help/13853/windows-lifecycle-fact-sheet</a></p>

        <p style="font-family:calibri;font-size:13px">
        Due to security concerns, we recommend you need to upgrade them to higher version as soon as possible.

    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        <table>
    <tr>
        <th>Windows 10 version</th>
        <th>End of Support</th>
        <th># of Objects</th>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1709 (16299) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">October 13, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1803 (17134) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1809 (17763) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1909 (18363) <br>
        Pro Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>  
    <tr>
        <td style="text-align:center">Windows 10 Pro 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>    
    <tr>
        <td style="text-align:center">Windows 10 Enterprise 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
        </table>


    <p style="font-family:calibri;font-size:13px">
        If you have any questions or concerns regarding this issue, please do not hesitate to contact us.
    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        
    <br><strong>(Your Signature Here)</strong></br>   


    </body>
    </html>

    """

    ###--------End of HTML body message--------###

    ###Attaching the CSV files to the Email###

    filelist = os.listdir(directory)
    g06 = []
    for csv_file in filelist:
        if csv_file .__contains__("_g06_"):
            g06.append(csv_file)
        else:
            continue

    for file_name in g06:
        with open(file_name, 'rb') as filetosend:
            record = MIMEBase('application', 'octet-stream')
            record.set_payload(filetosend.read())
            encoders.encode_base64(record)
            record.add_header('Content-Disposition', 'attachment',
                              filename=os.path.basename(file_name))
            msg.attach(record)

            part = MIMEText(html, 'html')
            msg.attach(part)

    if g06 != []:
        try:
            server = smtplib.SMTP('webmail.g07.fujitsu.local')
            server.sendmail(msg_from, msg_to, msg.as_string())
        except SMTPException:
            messagebox.showerror("Error", "Error: unable to send G06 email")
        server.quit()

    else:
        pass
###Function that will create G05 Email template with Attached Files###


def send_email_g05(msg_from, msg_to):

    ###Create message container - the correct MIME type is multipart/alternative###
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "[Request] Upgrading Unsupported Windows 10 (Data as of MONTH DATE YEAR)"
    msg['From'] = msg_from
    msg['To'] = msg_to

    ###Create the body of the message -- HTML###

    #************G05 Admin ************#
    html = """
    <!DOCTYPE html> 
    <html>
    <head>
    <style>
    table {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    border-collapse: collapse;
    text-align: middle;
    width: 55%;
    style: "text-align:center"
    }

    td {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    text-align: middle;
    style: "text-align:center";
    }

    th {
    border: 1px solid black;
    text-align: middle;
    padding: 1px;
    background-color:#ccccff;
    }

    </style>
    </head>

    <body>
    <p style="font-family:calibri;font-size:13px">
    Dear G05 Admins,

        <br><br>This is <strong>(Your name here)</strong> on behalf of CIT SOC Tier 1 and CIT SOC Tier 2.</br>
        --------------------------------------------------------------------------------------------------------------------------</br>
        
        We are sending you this notification according to our schedule of semimonthly updates with the latest data generated on 
        <strong>(MONTH DATE YEAR)</strong>.</br>
        <br>If you’ve already been contacting us to provide with your upgrading schedule relating to this issue, please just keep the file(s) 
        for your reference. Otherwise, please read the following and make corresponding action at your earliest convenience.</br>

        --------------------------------------------------------------------------------------------------------------------------</br>
        Please see the attached file(s) of computer(s) in your domain using Windows 10 Versions of which its support had already been expired.
        The dates of support termination for each Windows 10 version are listed in the following table.</br></p>

        <p style="font-family:calibri;font-size:13px;color:blue">
        *For more information about Windows 10 lifecycles, please refer to below URL:
        <br> <a href="https://docs.microsoft.com/en-us/lifecycle/faq/windows">
        https://support.microsoft.com/en-us/help/13853/windows-lifecycle-fact-sheet</a></p>

        <p style="font-family:calibri;font-size:13px">
        Due to security concerns, we recommend you need to upgrade them to higher version as soon as possible.

    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        <table>
    <tr>
        <th>Windows 10 version</th>
        <th>End of Support</th>
        <th># of Objects</th>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1511 (10586)</td>
        <td style="text-align:center">October 10, 2017</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1607 (14393) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">April 9, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1703 (15063) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">October 8, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1709 (16299) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">October 13, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1803 (17134) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1803 (17134) <br>
        Pro Edition </br></td>
        <td style="text-align:center">November 12, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1809 (17763) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1809 (17763) <br>
        Pro Edition </br></td>
        <td style="text-align:center">May 12, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1909 (18363) <br>
        Pro Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>  
    <tr>
        <td style="text-align:center">Windows 10 Pro 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>    
    <tr>
        <td style="text-align:center">Windows 10 Enterprise 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
        </table>


    <p style="font-family:calibri;font-size:13px">
        If you have any questions or concerns regarding this issue, please do not hesitate to contact us.
    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        
    <br><strong>(Your Signature Here)</strong></br>   


    </body>
    </html>

    """

    ###--------End of HTML body message--------###

    ###Attaching the CSV files to the Email###

    filelist = os.listdir(directory)
    g05 = []
    for csv_file in filelist:
        if csv_file .__contains__("_g05_"):
            g05.append(csv_file)
        else:
            continue

    for file_name in g05:
        with open(file_name, 'rb') as filetosend:
            record = MIMEBase('application', 'octet-stream')
            record.set_payload(filetosend.read())
            encoders.encode_base64(record)
            record.add_header('Content-Disposition', 'attachment',
                              filename=os.path.basename(file_name))
            msg.attach(record)

            part = MIMEText(html, 'html')
            msg.attach(part)

    if g05 != []:
        try:
            server = smtplib.SMTP('webmail.g07.fujitsu.local')
            server.sendmail(msg_from, msg_to, msg.as_string())
        except SMTPException:
            messagebox.showerror("Error", "Error: unable to send G05 email")
        server.quit()
    else:
        pass

###Function that will create G02 Email template with Attached Files###


def send_email_g02(msg_from, msg_to):

    ###Create message container - the correct MIME type is multipart/alternative###
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "[Request] Upgrading Unsupported Windows 10 (Data as of MONTH DATE YEAR)"
    msg['From'] = msg_from
    msg['To'] = msg_to

    ###Create the body of the message -- HTML###

    #************G02 Admin ************#
    html = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
    table {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    border-collapse: collapse;
    text-align: middle;
    width: 55%;
    style: "text-align:center"
    }

    td {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    text-align: middle;
    style: "text-align:center";
    }

    th {
    border: 1px solid black;
    text-align: middle;
    padding: 1px;
    background-color:#ccccff;
    }

    </style>
    </head>

    <body>
    <p style="font-family:calibri;font-size:13px">
    Dear G02 Admins,

        <br><br>This is <strong>(Your name here)</strong> on behalf of CIT SOC Tier 1 and CIT SOC Tier 2.</br>
        --------------------------------------------------------------------------------------------------------------------------</br>
        
        We are sending you this notification according to our schedule of semimonthly updates with the latest data generated on 
        <strong>(MONTH DATE YEAR)</strong>.</br>
        <br>If you’ve already been contacting us to provide with your upgrading schedule relating to this issue, please just keep the file(s) 
        for your reference. Otherwise, please read the following and make corresponding action at your earliest convenience.</br>

        --------------------------------------------------------------------------------------------------------------------------</br>
        Please see the attached file(s) of computer(s) in your domain using Windows 10 Versions of which its support had already been expired.
        The dates of support termination for each Windows 10 version are listed in the following table.</br></p>

        <p style="font-family:calibri;font-size:13px;color:blue">
        *For more information about Windows 10 lifecycles, please refer to below URL:
        <br> <a href="https://docs.microsoft.com/en-us/lifecycle/faq/windows">
        https://support.microsoft.com/en-us/help/13853/windows-lifecycle-fact-sheet</a></p>

        <p style="font-family:calibri;font-size:13px">
        Due to security concerns, we recommend you need to upgrade them to higher version as soon as possible.

    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        <table>
    <tr>
        <th>Windows 10 version</th>
        <th>End of Support</th>
        <th># of Objects</th>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1703 (15063) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">October 8, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1803 (17134) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 1803 (17134) <br>
        Pro Edition </br></td>
        <td style="text-align:center">November 12, 2019</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1809 (17763) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1909 (18363) <br>
        Pro Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>  
    <tr>
        <td style="text-align:center">Windows 10 Pro 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>    
    <tr>
        <td style="text-align:center">Windows 10 Enterprise 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
        </table>

    <p style="font-family:calibri;font-size:13px">
        If you have any questions or concerns regarding this issue, please do not hesitate to contact us.
    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        
    <br><strong>(Your Signature Here)</strong></br>   


    </body>
    </html>

    """

    ###--------End of HTML body message--------###

    ###Attaching the CSV files to the Email###

    filelist = os.listdir(directory)
    g02 = []
    for csv_file in filelist:
        if csv_file .__contains__("_g02_"):
            g02.append(csv_file)
        else:
            continue

    for file_name in g02:
        with open(file_name, 'rb') as filetosend:
            record = MIMEBase('application', 'octet-stream')
            record.set_payload(filetosend.read())
            encoders.encode_base64(record)
            record.add_header('Content-Disposition', 'attachment',
                              filename=os.path.basename(file_name))
            msg.attach(record)

            part = MIMEText(html, 'html')
            msg.attach(part)

    if g02 != []:
        try:
            server = smtplib.SMTP('webmail.g07.fujitsu.local')
            server.sendmail(msg_from, msg_to, msg.as_string())
        except SMTPException:
            messagebox.showerror("Error", "Error: unable to send G02 email")
        server.quit()
    else:
        pass

##------------------------ Domains with almost NO REPORT ------------------------##

###Function that will create G03 Email template with Attached Files###


def send_email_g03(msg_from, msg_to):

    ###Create message container - the correct MIME type is multipart/alternative###
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "[Request] Upgrading Unsupported Windows 10 (Data as of MONTH DATE YEAR)"
    msg['From'] = msg_from
    msg['To'] = msg_to

    ###Create the body of the message -- HTML###

    #************G03 Admin ************#
    html = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
    table {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    border-collapse: collapse;
    text-align: middle;
    width: 55%;
    style: "text-align:center"
    }

    td {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    text-align: middle;
    style: "text-align:center";
    }

    th {
    border: 1px solid black;
    text-align: middle;
    padding: 1px;
    background-color:#ccccff;
    }

    </style>
    </head>

    <body>
    <p style="font-family:calibri;font-size:13px">
    Dear G03 Admins,

        <br><br>This is <strong>(Your name here)</strong> on behalf of CIT SOC Tier 1 and CIT SOC Tier 2.</br>
        --------------------------------------------------------------------------------------------------------------------------</br>
        
        We are sending you this notification according to our schedule of semimonthly updates with the latest data generated on 
        <strong>(MONTH DATE YEAR)</strong>.</br>
        <br>If you’ve already been contacting us to provide with your upgrading schedule relating to this issue, please just keep the file(s) 
        for your reference. Otherwise, please read the following and make corresponding action at your earliest convenience.</br>

        --------------------------------------------------------------------------------------------------------------------------</br>
        Please see the attached file(s) of computer(s) in your domain using Windows 10 Versions of which its support had already been expired.
        The dates of support termination for each Windows 10 version are listed in the following table.</br></p>

        <p style="font-family:calibri;font-size:13px;color:blue">
        *For more information about Windows 10 lifecycles, please refer to below URL:
        <br> <a href="https://docs.microsoft.com/en-us/lifecycle/faq/windows">
        https://support.microsoft.com/en-us/help/13853/windows-lifecycle-fact-sheet</a></p>

        <p style="font-family:calibri;font-size:13px">
        Due to security concerns, we recommend you need to upgrade them to higher version as soon as possible.

    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        <table>
    <tr>
        <th>Windows 10 version</th>
        <th>End of Support</th>
        <th># of Objects</th>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 ? (?) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">Date ? </td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 ? (?) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">Date ? </td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 ? (?) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">Date ? </td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1909 (18363) <br>
        Pro Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>  
    <tr>
        <td style="text-align:center">Windows 10 Pro 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>    
    <tr>
        <td style="text-align:center">Windows 10 Enterprise 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
        </table>

    <p style="font-family:calibri;font-size:13px">
        If you have any questions or concerns regarding this issue, please do not hesitate to contact us.
    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        
    <br><strong>(Your Signature Here)</strong></br>   


    </body>
    </html>

    """

    ###--------End of HTML body message--------###

    ###Attaching the CSV files to the Email###

    filelist = os.listdir(directory)
    g03 = []
    for csv_file in filelist:
        if csv_file .__contains__("_g03_"):
            g03.append(csv_file)
        else:
            continue

    for file_name in g03:
        with open(file_name, 'rb') as filetosend:
            record = MIMEBase('application', 'octet-stream')
            record.set_payload(filetosend.read())
            encoders.encode_base64(record)
            record.add_header('Content-Disposition', 'attachment',
                              filename=os.path.basename(file_name))
            msg.attach(record)

            part = MIMEText(html, 'html')
            msg.attach(part)

    if g03 != []:
        try:
            server = smtplib.SMTP('webmail.g07.fujitsu.local')
            server.sendmail(msg_from, msg_to, msg.as_string())
        except SMTPException:
            messagebox.showerror("Error", "Error: unable to send G03 email")
        server.quit()
    else:
        pass

###Function that will create G04 Email template with Attached Files###


def send_email_g04(msg_from, msg_to):

    ###Create message container - the correct MIME type is multipart/alternative###
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "[Request] Upgrading Unsupported Windows 10 (Data as of MONTH DATE YEAR)"
    msg['From'] = msg_from
    msg['To'] = msg_to

    ###Create the body of the message -- HTML###

    #************G04 Admin ************#
    html = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
    table {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    border-collapse: collapse;
    text-align: middle;
    width: 55%;
    style: "text-align:center"
    }

    td {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    text-align: middle;
    style: "text-align:center";
    }

    th {
    border: 1px solid black;
    text-align: middle;
    padding: 1px;
    background-color:#ccccff;
    }

    </style>
    </head>

    <body>
    <p style="font-family:calibri;font-size:13px">
    Dear G04 Admins,

        <br><br>This is <strong>(Your name here)</strong> on behalf of CIT SOC Tier 1 and CIT SOC Tier 2.</br>
        --------------------------------------------------------------------------------------------------------------------------</br>
        
        We are sending you this notification according to our schedule of semimonthly updates with the latest data generated on 
        <strong>(MONTH DATE YEAR)</strong>.</br>
        <br>If you’ve already been contacting us to provide with your upgrading schedule relating to this issue, please just keep the file(s) 
        for your reference. Otherwise, please read the following and make corresponding action at your earliest convenience.</br>

        --------------------------------------------------------------------------------------------------------------------------</br>
        Please see the attached file(s) of computer(s) in your domain using Windows 10 Versions of which its support had already been expired.
        The dates of support termination for each Windows 10 version are listed in the following table.</br></p>

        <p style="font-family:calibri;font-size:13px;color:blue">
        *For more information about Windows 10 lifecycles, please refer to below URL:
        <br> <a href="https://docs.microsoft.com/en-us/lifecycle/faq/windows">
        https://support.microsoft.com/en-us/help/13853/windows-lifecycle-fact-sheet</a></p>

        <p style="font-family:calibri;font-size:13px">
        Due to security concerns, we recommend you need to upgrade them to higher version as soon as possible.

    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        <table>
    <tr>
        <th>Windows 10 version</th>
        <th>End of Support</th>
        <th># of Objects</th>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 ? (?) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">Date ? </td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 ? (?) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">Date ? </td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 ? (?) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">Date ? </td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1909 (18363) <br>
        Pro Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>  
    <tr>
        <td style="text-align:center">Windows 10 Pro 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>    
    <tr>
        <td style="text-align:center">Windows 10 Enterprise 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
        </table>

    <p style="font-family:calibri;font-size:13px">
        If you have any questions or concerns regarding this issue, please do not hesitate to contact us.
    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        
    <br><strong>(Your Signature Here)</strong></br>   


    </body>
    </html>

    """

    ###--------End of HTML body message--------###

    ###Attaching the CSV files to the Email###

    filelist = os.listdir(directory)
    g04 = []
    for csv_file in filelist:
        if csv_file .__contains__("_g04_"):
            g04.append(csv_file)
        else:
            continue

    for file_name in g04:
        with open(file_name, 'rb') as filetosend:
            record = MIMEBase('application', 'octet-stream')
            record.set_payload(filetosend.read())
            encoders.encode_base64(record)
            record.add_header('Content-Disposition', 'attachment',
                              filename=os.path.basename(file_name))
            msg.attach(record)

            part = MIMEText(html, 'html')
            msg.attach(part)

    if g04 != []:
        try:
            server = smtplib.SMTP('webmail.g07.fujitsu.local')
            server.sendmail(msg_from, msg_to, msg.as_string())
        except SMTPException:
            messagebox.showerror("Error", "Error: unable to send G04 email")
        server.quit()
    else:
        pass

###Function that will create R01 Email template with Attached Files###


def send_email_r01(msg_from, msg_to):

    ###Create message container - the correct MIME type is multipart/alternative###
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "[Request] Upgrading Unsupported Windows 10 (Data as of MONTH DATE YEAR)"
    msg['From'] = msg_from
    msg['To'] = msg_to

    ###Create the body of the message -- HTML###

    #************R01 Admin ************#
    html = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
    table {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    border-collapse: collapse;
    text-align: middle;
    width: 55%;
    style: "text-align:center"
    }

    td {
    font-family: Fujitsu Sans;
    font-size: 13px;
    border: 1px solid black;
    text-align: middle;
    style: "text-align:center";
    }

    th {
    border: 1px solid black;
    text-align: middle;
    padding: 1px;
    background-color:#ccccff;
    }

    </style>
    </head>

    <body>
    <p style="font-family:calibri;font-size:13px">
    Dear R01 Admins,

        <br><br>This is <strong>(Your name here)</strong> on behalf of CIT SOC Tier 1 and CIT SOC Tier 2.</br>
        --------------------------------------------------------------------------------------------------------------------------</br>
        
        We are sending you this notification according to our schedule of semimonthly updates with the latest data generated on 
        <strong>(MONTH DATE YEAR)</strong>.</br>
        <br>If you’ve already been contacting us to provide with your upgrading schedule relating to this issue, please just keep the file(s) 
        for your reference. Otherwise, please read the following and make corresponding action at your earliest convenience.</br>

        --------------------------------------------------------------------------------------------------------------------------</br>
        Please see the attached file(s) of computer(s) in your domain using Windows 10 Versions of which its support had already been expired.
        The dates of support termination for each Windows 10 version are listed in the following table.</br></p>

        <p style="font-family:calibri;font-size:13px;color:blue">
        *For more information about Windows 10 lifecycles, please refer to below URL:
        <br> <a href="https://docs.microsoft.com/en-us/lifecycle/faq/windows">
        https://support.microsoft.com/en-us/help/13853/windows-lifecycle-fact-sheet</a></p>

        <p style="font-family:calibri;font-size:13px">
        Due to security concerns, we recommend you need to upgrade them to higher version as soon as possible.

    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        <table>
    <tr>
        <th>Windows 10 version</th>
        <th>End of Support</th>
        <th># of Objects</th>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 ? (?) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">Date ? </td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 ? (?) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">Date ? </td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10 ? (?) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">Date ? </td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Enterprise Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1903 (18362) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 8, 2020</td>
        <td style="text-align:center"> ?  </td>
    </tr>
    <tr>
        <td style="text-align:center">Windows 10  1909 (18363) <br>
        Pro Edition </br></td>
        <td style="text-align:center">May 11, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>  
    <tr>
        <td style="text-align:center">Windows 10 Pro 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>    
    <tr>
        <td style="text-align:center">Windows 10 Enterprise 2004 (20H1) (19041) <br>
        Pro Edition </br></td>
        <td style="text-align:center">December 14, 2021</td>
        <td style="text-align:center"> ?  </td>
    </tr>
        </table>

    <p style="font-family:calibri;font-size:13px">
        If you have any questions or concerns regarding this issue, please do not hesitate to contact us.
    <br>--------------------------------------------------------------------------------------------------------------------------</br></p>
        
        
    <br><strong>(Your Signature Here)</strong></br>   


    </body>
    </html>

    """

    ###--------End of HTML body message--------###

    ###Attaching the CSV files to the Email###

    filelist = os.listdir(directory)
    r01 = []
    for csv_file in filelist:
        if csv_file .__contains__("_r01_"):
            r01.append(csv_file)
        else:
            continue

    for file_name in r01:
        with open(file_name, 'rb') as filetosend:
            record = MIMEBase('application', 'octet-stream')
            record.set_payload(filetosend.read())
            encoders.encode_base64(record)
            record.add_header('Content-Disposition', 'attachment',
                              filename=os.path.basename(file_name))
            msg.attach(record)

            part = MIMEText(html, 'html')
            msg.attach(part)

    if r01 != []:
        try:
            server = smtplib.SMTP('webmail.g07.fujitsu.local')
            server.sendmail(msg_from, msg_to, msg.as_string())
        except SMTPException:
            messagebox.showerror("Error", "Error: unable to send R01 email")
        server.quit()
    else:
        pass


def main():
    removefiles(directory)
    countItems(directory)
    ##first email add is "msg_from", second email add is "msg_to"##
    msg_is_from = "teamdistrolist@xyzcompany.com"
    msg_is_to = (
        "user01@xyzcompany.com", 
        "user02@xyzcompany.com",
        "user03@xyzcompany.com",
        "user04@xyzcompany.com",
        "user05@xyzcompany.com",
        "user06@xyzcompany.com",
        "user07@xyzcompany.com",
        "user08@xyzcompany.com",
        "user09@xyzcompany.com")

    for x in msg_is_to:
            send_email_g0x(msg_is_from,x)
            send_email_g0x(msg_is_from,x)
            send_email_g0x(msg_is_from,x)
            send_email_g0x(msg_is_from,x)
            send_email_g0x(msg_is_from,x)
            send_email_g0x(msg_is_from,x)


    messagebox.showinfo("Info", "Processing Complete")


if __name__ == "__main__":
    main()

