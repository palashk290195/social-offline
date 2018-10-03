from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import re
import math
import csv
from collections import OrderedDict
import smtplib  
from email import encoders
from email.message import Message
from email.mime.text import MIMEText
import base64
import urllib2
import sys
# from google import searchGoogle

# If modifying these scopes, delete the file token.json.
SCOPES = ('https://www.googleapis.com/auth/gmail.send ' + 'https://www.googleapis.com/auth/spreadsheets')

# The ID and range of a sample spreadsheet.
#RESPONSE_SPREADSHEET_ID = '1cUJQKFFP-2A3Cd2RTsdHjEJ86FjCNBhkBEoa4qLQP94'
RESPONSE_SPREADSHEET_ID = ''
RESPONSE_RANGE_NAME = 'A:E'

COMBINED_MATCHING_SPREADSHEET_ID = '1QOk2kXpiHQ3_xMwFmcoirFpVXY0RupJsoUcoP7qaqj0'
COMBINED_MATCHING_RANGE_NAME = 'A:B'

value_render_option = 'FORMATTED_VALUE'

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    store = file.Storage('token.json')
    creds = store.get()
    #flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
    #creds = tools.run_flow(flow, store)
     
    #creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))

    # Call the Sheets API
    RESPONSE_SPREADSHEET_ID = sys.argv[1]
    response_result = service.spreadsheets().values().get(spreadsheetId=RESPONSE_SPREADSHEET_ID,
                                                range=RESPONSE_RANGE_NAME, valueRenderOption=value_render_option).execute()
    response_values = response_result.get('values', [])
    # response_values = [[google.searchGoogle(y).encode("utf-8") for y in x] for x in response_values]

    if not response_values:
        print('No data found.')
    #else:
    #    for row in response_values:
    #        print(row)

    combined_matching_result = service.spreadsheets().values().get(spreadsheetId=COMBINED_MATCHING_SPREADSHEET_ID,
                                                range=COMBINED_MATCHING_RANGE_NAME, valueRenderOption=value_render_option).execute()
    combined_matching_values = combined_matching_result.get('values', [])
    # combined_matching_values = [[google.searchGoogle(y).encode("utf-8") for y in x] for x in combined_matching_values]

    if not combined_matching_values:
        print('No data found.')
    # else:
    #     for row in combined_matching_values:
    #         print(row)
    matchingList = matching(response_values, combined_matching_values)
    print(matchingList)
    if sys.argv[2] is "2":
        print("is 2")
        #TODO: Send emails
        gmailService = build('gmail', 'v1', http=creds.authorize(Http()))
        for row in matchingList:
            message = createMessage(row)
            sendMessage(gmailService, message)

        #TODO: Add to combined matching
        appendToCombinedMatching = []
        for row in matchingList:
            if row[1] is not "":
                appendToCombinedMatching.append([row[0],row[1],"social-offline"])
        body = {
        'values': appendToCombinedMatching
        }
        # result = service.spreadsheets().values().update(
        # spreadsheetId=SENDING_EMAILS_SPREADSHEET_ID, range=SENDING_EMAILS_RANGE_NAME, 
        # valueInputOption = 'RAW', body=body).execute()

        
        response = service.spreadsheets().values().append(spreadsheetId=COMBINED_MATCHING_SPREADSHEET_ID, range=COMBINED_MATCHING_RANGE_NAME, valueInputOption='RAW', body=body).execute()
        # sendingEmailsSheet = service.open(SENDING_EMAILS_SPREADSHEET_ID).sheet1
        # updateMatching(matchingList, sendingEmailsSheet)

def sendMessage(gmailService, message):
    print("sending")
    try:
        message = (gmailService.users().messages().send(userId="me", body=message).execute())
        print('Message Id: %s' % message['id'])
        return message
    except urllib2.HTTPError as error:
        print('An error occurred: %s' % error)

def createMessage(row):
    """Create a message for an email.

    Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

    Returns:
    An object containing a base64url encoded email object.
    """
    emailAddress1 = row[0] # First column
    emailAddress2 = row[1] # Second column
    timeSlot = row[2] # Third column
    name1 = row[3]
    name2 = row[4]
    contact1 = row[5]
    contact2 = row[6]
    cc1 = "palashrajendra.kala_yif19@ashoka.edu.in," + "sahil.mahajan_yif19@ashoka.edu.in," + "amruta.pagariya_yif19@ashoka.edu.in," + "komal.manglani_yif19@ashoka.edu.in," + "neelansha.trivedi_yif19@ashoka.edu.in"
    subject1 = 'Your Social Offline Match for today!'
    message_text = "Hi Guys,\n\n" + "The Social Offline initiative received amazing responses " + "and we have tried to match you guys as per the time slots you guys are free in.\n" +name1 + " and " + name2 + " your slot is " + timeSlot + ".\n"+ "Contact details are " + contact1 + " and " + contact2 + " respectively.\n\n"+"So go ahead, ping the person you have been matched with." + "In case you do not have the contact details of the other person, " + "please get in touch with the Social Offline team members, " + "whose details have been mentioned below. \n\nThe default venue for the meeting is Outside the RH2 area " + "(for those of you who don't want to go through the hassle and formality of deciding on where to meet up)." + "For others, the complete campus is your playground! " + "Please coordinate and fix any venue you guys would love to meet at and have an amazing conversation.\n\n" + "For the sake of respecting other person's time, please do coordinate before hand with the person you are matched with " + "and communicate in case you won't be able to make it on designated time and need a change. " + "You can communicate directly and have a word with your partner." + "And last but not the least, we would love to hear from you and get your feedback about how did your meetup go." + "So please ping us on mail/call/whatsapp/catch hold of us in the campus or fill this form: https://goo.gl/forms/mxFQgs9VN8gA18vg1." +"\n\nHave a great time guys! Hopefully this would bring us one step closer to getting to know each other. Cheers! \n\nThanks and Love\nSocial Offline Team\n \nAmruta Pagariya (7588350072)\nNeelansha Trivedi (9650420000)\nKomal Manglani (9149238422) \nPalash Kala (9665543333) \nSahil Mahajan (9818512958)"
    message = MIMEText(message_text)
    message['to'] = emailAddress1 + "," + emailAddress2
    #message['from'] = "me"
    message['subject'] = subject1
    message['cc'] = cc1
    return {'raw': base64.urlsafe_b64encode(message.as_string())}

#timeslots is a string as follows: "06:00 PM, 07:00 PM, 08:00 PM, 09:00 PM, 10:00 PM, 11:00 PM, 12:00 AM, 01:00 AM, 02:00 AM, 03:00 AM"
#Or timeslots can be a single element as 08:00 PM
def formSet(timeslots):
    timeSet = timeslots.split(",")
    timeSet = map(unicode.strip, timeSet)
    #print(timeSet)
    return set(timeSet)

def matching(responses, combinedMatchings):
    #for each line in responses
    #convert time slots in inverted commas to a set or array may be randomized
    #add to responsedictionary:  email id: Set of timeslots
    #responseDictionary = OrderedDict({})
    responseDictionary = {}
    namesDictionary = {}
    contactsDictionary = {}
    iterResponses = iter(responses)
    next(iterResponses)
    for row in iterResponses:
        #print(row)
        email = row[1]
        name = row[2]
        timeslots = row[3]
        if len(row) > 4:
            contact = row[4]
        else:
            contact = "Contact not provided"
        responseDictionary[email] = formSet(timeslots)
        namesDictionary[email] = name
        contactsDictionary[email] = contact
        #print(responseDictionary[email])

    #for each line in matching
    #add to nomatchingdictionary: email id1: Set1 + email id2
    #email id2 = Set2 + email id 1
    noMatchingDictionary = {}
    iterCombinedMatchings = iter(combinedMatchings)
    for row in iterCombinedMatchings:
        email1 = row[0].strip()
        email2 = row[1].strip()
        if email1 in noMatchingDictionary:
            noMatchingDictionary[email1].add(email2)
        else:
            noMatchingDictionary[email1] = set([email2])
        if email2 in noMatchingDictionary:
            noMatchingDictionary[email2].add(email1)
        else:
            noMatchingDictionary[email2] = set([email1])

    #print(noMatchingDictionary)

    #for each element in responsedictionary
    #for all other elements in response dictionary
    #If not in nomatchingdictionary
    #If they have same time slots
    #print matched and then print time slot and two email ids
    #remove both of them from response directory

    matchingList = []

    for email1 in responseDictionary:
        for email2 in responseDictionary:
            if email1 is not email2:
                if (email1 not in noMatchingDictionary) or (email2 not in noMatchingDictionary[email1]):
                    if responseDictionary[email1] & responseDictionary[email2]:
                        matchingList.append([email1,email2,
                            str(next(iter(responseDictionary[email1] & responseDictionary[email2]))),
                            namesDictionary[email1], namesDictionary[email2], contactsDictionary[email1], contactsDictionary[email2]])
                        responseDictionary[email1] = set([])
                        responseDictionary[email2] = set([])

    for email1 in responseDictionary:
        for email2 in responseDictionary:
            if email1 is not email2:
                if (email1 not in noMatchingDictionary) or (email2 not in noMatchingDictionary[email1]):
                    if bool(responseDictionary[email1]) and bool(responseDictionary[email2]):
                        matchingList.append([email1,email2,"Please coordinate your time",
                            namesDictionary[email1], namesDictionary[email2], contactsDictionary[email1], contactsDictionary[email2]])
                        responseDictionary[email1] = set([])
                        responseDictionary[email2] = set([])

    #print(matchingList)

    #TODO: Add elements to the matching file which have been matched

    #print(" ")

    for email in responseDictionary:
        if bool(responseDictionary[email]):
            matchingList.append([email,"","Could not find a partner",
                namesDictionary[email], "", contactsDictionary[email], ""])

    return matchingList
    #print(matchingList)

if __name__ == '__main__':
    main()
