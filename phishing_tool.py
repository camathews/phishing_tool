# Phishing Tool
# Integrates with Outlook
# requires a Virus Total API Key and a urlscan.io API Key


import urllib.parse
import urllib.request
import requests
import webbrowser
import json
import time
import win32com.client as win32

def inputURL():
    url = input("Enter wrapped URL: ") # wrapped by Office 365 ATP Safe Links
    return(url)

# unwraps ATP Safe Links. 
def unwrap(wrappedURL):
    return(urllib.parse.unquote(wrappedURL.split('=')[1]))

# calls the Virus Total API
# insert your own key for "your_key_here"
def vtScan(unwrappedURL):
    url = 'https://www.virustotal.com/vtapi/v2/url/scan'
    params = {'apikey': 'your_key_here', 'url': unwrappedURL}
    response = requests.post(url, data=params)
    webbrowser.open(response.json()["permalink"])

# calls the urlscan.io API
# insert your own key for "your_key_here"
def urlScan(unwrappedURL):
    url = 'https://urlscan.io/api/v1/scan/'
    params = {'Content-Type': 'json', 'url': unwrappedURL, 'public': 'on'}
    headers = {'API-Key': 'your_key_here'}
    response = requests.post(url, data=params, headers=headers)
    uuid = response.json()["uuid"]
    webbrowser.open("https://urlscan.io/screenshots/" + uuid + ".png")

# creates an email in Outlook
#insert the TO address in mail.to and CC address in mail.cc
def emailer(subject, ticketNum, url):
    emailBody = "In reference to: <br> Ticket: " + ticketNum + "<br><br>Please block the following URL: " + url
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = "insert_email_address" # seperate multiple emails with semi colons - "email@mail.com; email2@mail.com"
    mail.CC = "insert_email_address" 
    mail.Subject = "PHISH: " + subject
    mail.HtmlBody = emailBody
    mail.Display(True)

def main():
    url = unwrap(inputURL())
    vtScan(url)
    urlScan(url)
    if (input("Enter y to draft email or n to exit: ") == 'y'):
        sub = input("Enter email subject: ")
        ticketNum = input("Enter VU Ticker number: ")
        emailer(sub, ticketNum, url)
    else:
        print("Exit")

main()
