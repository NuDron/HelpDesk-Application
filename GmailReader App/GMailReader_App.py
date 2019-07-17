"""
CREATOR:    Krzysztof Szumko
LINKED-IN:  https://www.linkedin.com/in/krzysztofszumko1989
GITHUB:     https://github.com/NuDron
Version:    0.07
"""

import datetime
import email
import imaplib
import os
import base64
import fernet
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from cryptography.fernet import Fernet

# Currently unused.
def savePass():
    """
    The method reads a key for encryption and then writes the password(Gmail account) in config.txt file.
    :return: None
    """
    file1 = open("key.key", "rb")
    key = file1.read()  # The key will be in bytes
    file1.close()
    f = Fernet(key)
    encoded = password.encode()
    encrypted = f.encrypt(encoded)
    file2 = open("config.txt", "wb")
    file2.write(encrypted)
    file2.close()


def getKey():
    """
    Returns a key used for encryption/decryption (cryptography fernet) from file present in the same directory as
    executable (file named key.key).
    :return: encryption key
    """
    file1 = open("key.key", "rb")
    key = file1.read()  # The key will be in bytes
    file1.close()
    return key


def getSavedPass():
    """
    Uses a previously stored key(encrypted using cryptography fernet) for decrypting the password for particular Gmail account that serves as
    way of sending tickets between two clients.
    :return: decrypted password
    """
    file = open('key.key', 'rb')
    key2 = file.read()
    file.close()
    f2 = Fernet(key2)
    file2 = open("config.txt", "rb")
    haslo = file2.read()
    decrypted = f2.decrypt(haslo)
    return decrypted


def getPasswordAsString():
    """
    Stringifies the password returned by getSavedPass method.
    :return: String
    """
    return str(getSavedPass().decode())


def setAttachDir():
    """
    The method is invoked to set a path to directory meant for storing the images (attached to tickets).
    :return: None
    """
    currdir = os.getcwd()
    tempdir = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')
    if len(tempdir) > 0:
        print("You chose: %s" % tempdir)
    attachment_dir = tempdir


def get_body(msg):
    """
    Extracts the body of the email.
    :return: message payload in MIME format
    """
    # Check if email contains any attachments. Loops through all of the payloads and return the value when 1st is found.
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        # In case there's no multiple payloads this is returned.
        return msg.get_payload(None, True)


def search(key, value, con):
    """
    The method searches the emails in designated Gmail account (value) based on the chosen variable (key).
    Con is used to determine the port of connection to Gmail account.
    :param key: use TO or FROM
    :param value: EMAIL adress here - the value want to check for.
    :param con: connection used

    EXAMPLE:
         print(search('FROM', 'kszumko@gmail.com', connection))
         gives:
                [b'1 3'] #List of index value of eMails that satisfies the value input.
    :return: List of numbers: in [b'1 2 ... n'] format
    """
    result, data = con.search(None, key, '"{}"'.format(value))  # Value will be inserted into {}.
    return data


def get_emails(result_bytes):
    """
    Parse the found emails (result from search).
    Initialise and empty list and appends emails to list which is returned.
    :param result_bytes: search function results = uses the list of mails index values to fetch.
    :return: List of messages(msgs): Returns a list containing all found messages that
    were fetched using list of index(from search function).
    """
    msgs = []
    for num in result_bytes[0].split():
        typ, data = connection.fetch(num, '(RFC822)')
        msgs.append(data)
    return msgs


def get_attachments(msg):
    """
    This function is checking email(msg) for attachments and storing it in previously specified folder.
    :param msg: The message to search through
    INVARIANT: attachment_dir: Needs to be specified
    :return: The attachment file is stored in previously specified folder.
    """
    # Walk() allows to go through content of message.
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        # Filename is obtained at this point.
        fileName = part.get_filename()

        # False means no attachment.
        if bool(fileName):
            filePath = os.path.join(attachment_dir,
                                    fileName)  # <--Joining of the filename and file path (previously specified).
            # With open automatically close the file after reading.
            with open(filePath, 'wb') as file:
                file.write(part.get_payload(decode=True))


def get_unseen_mails():
    """
    The method outputs a list of eMail IDs, which have status 'UNSEEN'.
    EXAMPLE: [b'1', b'2']
    :return: List of IDs (emails)
    """
    result, data = connection.search(None, '(UNSEEN)')
    mail_ids = data[0].split()
    return mail_ids


def get_archived_mails():
    """
    The method outputs a list of Ids of e-mails that are in mailbox 'Important', but are not in mailbox 'Inbox'.
    :return: List of Ids
    """
    result = []
    # Select Gmail subfolder "IMPORTANT"
    connection.select('[Gmail]/Important')
    status, data = connection.search(None, 'ALL')
    mail_ids_Important = data[0].split()
    connection.select('INBOX')
    status, data = connection.search(None, 'ALL')
    mail_ids_inbox = data[0].split()
    # Compare emails in INBOX and IMPORTANT - leave the one that are in IMPORTANT subfolder only.
    for eachId in mail_ids_Important:
        if eachId not in mail_ids_inbox:
            result.append(eachId)
    return result


def get_starred_emails():
    """
    The method outputs a list of eMail IDs, which have active star (are starred).
    :return: List of starred mail IDs
    """
    result, data = connection.search(None, 'FLAGGED')
    mail_ids = data[0].split()
    return mail_ids


def get_unflagged_emails():
    """
    The method outputs a list of email IDs, which has not been read yet(have 'UNFLAGGED' status).
    :return: List of mail IDs(unflagged/ not starred)
    """
    result, data = connection.search(None, 'UNFLAGGED')
    mail_ids = data[0].split()
    return mail_ids


def onlyIntegersFromString(aString):
    """
    Method for parsing integers values from String variable.
    :param aString: String value to parse integers from.
    :return: Integer value(s) parsed from aString (supplied argument)
    """
    result = int(re.search(r'\d+', aString).group())
    return result


def getSpecificLine(rawMail, strFind, splitStr):
    """
    The method splits the raw email(rawMail) by utilizing string to Find (strFind) and splitting string (splitStr).
    :param rawMail: the content of email parsed by IMAPlib(list or String).
    :param strFind: a String variable to find in rawMail.
    :param splitStr: a value that will be determine "string.replace()".
    :return: LIST value that starts from 'strFind' value and is split by 'splitStr'.
    """
    temp = str(rawMail)
    tempInt = temp.rfind(strFind)
    temp = (temp[tempInt:len(temp)])
    temp = temp.split(splitStr)
    return temp


# Helping method for printing in more distinctive way (by adding long series of "-" before and after given argument)
def printCom(aString):
    c = ""
    for x in range(60):  # Amount of "-" sign printed
        c += "-"
    print(c + "\n")
    print(aString)
    print("\n" + c)


def getSimpleEmails():
    """
    Method obtains eMails from selected Gmail account of MIME-type/regular and parse it for data.
    :return: listOfMails where KEY is mail ID (e.g. b'1') and value is parsed mail content (FROM, SUBJECT, DATE, TEXT).
    """
    titleTag = "<titolo>"
    fromTag = "<sendinto>"
    toTag = "<sendu_al>"
    dateTag = "<nuna_dato>"
    timeTag = "<nuna_tempo>"
    textTag = "<interpretus>"
    # Step 1: Get usable List of eMails in Inbox.
    connection.select('INBOX')
    mail_text, data = connection.search(None, 'ALL')
    mail_ids = data[0].split()  # <-- Outputs a list of [b'1', b'2', b'3']
    """Step 2: Creation of main frame 
    (dictionary => key is mail ID, value => List - based on email -(FROM, SUBJECT, DATE, TEXT)) """
    listOfMails = dict.fromkeys(mail_ids)  # <-- Outputs {b'1': None, b'2': None, b'3': None}
    for eachKey in listOfMails:
        mail_text, data = connection.fetch(eachKey, '(RFC822)')
        rawMail = email.message_from_bytes(data[0][1])
        haveTime = False
        for eachPart in rawMail.walk():
            if eachPart.get_content_maintype() == 'multipart':
                parts = len(eachPart.get_payload())
                part = str(eachPart)
                try:
                    # Step 3: Obtaining Date from emails.
                    mail_data = part.split(dateTag)
                    mail_data = str(mail_data[1])
                except:
                    print("error")
                    mail_data = "--/--/--"
                # Step 4: Obtaining From(sender).
                try:
                    mail_From = part.split(fromTag)
                    mail_From = str(mail_From[1])
                    if not mail_From.rfind("gmail.com"):
                        mail_From = "Junk Mail"
                except:
                    print("Error with mail_from")
                    mail_From = "--"

                # Step 5: Obtaining Subject(error code).
                try:
                    mail_Subject = part.split(titleTag)
                    mail_Subject = str(mail_Subject[1])
                except:
                    print("error with subject")
                    mail_Subject = "--"
                # Step 6: Obtaining Text from email and parsing it.
                try:
                    mail_text = part.split(textTag)
                    mail_text = str(mail_text[1])
                except:
                    mail_text = "---"
                # Step 7: Obtaining status of e-mail.
                notStarredM = get_unflagged_emails()
                mail_status = ""
                if eachKey not in notStarredM:
                    mail_status = "SOLVING"
                if part.rfind(timeTag):
                    haveTime = True
                    try:
                        mail_time = part.split(timeTag)
                        mail_time = str(mail_time[1])
                    except:
                        mail_time = "--:--:--"
            if parts > 1:
                try:
                    pLoad = eachPart.get_payload()
                    pLoad = pLoad[1]
                    get_attachments(pLoad)
                except:
                    print("No image found!")
            # Step 8: Constructing the List which will be inserted into listOfMails as value for eachKey.
            list_of_values = []
            list_of_values.append(shorterString(mail_Subject, 30))
            list_of_values.append(shorterString(mail_data, 30))
            list_of_values.append(shorterString(mail_From, 30))
            list_of_values.append(mail_text)
            list_of_values.append(eachKey)
            list_of_values.append(mail_status)
            if haveTime:
                list_of_values.append(mail_time)
            else:
                list_of_values.append("-")
            listOfMails[eachKey] = list_of_values

    if not listOfMails == {}:
        return listOfMails
    else:
        messagebox.showwarning('Warning', 'Mailbox is empty.')


def refreshSimple():
    """
    Method used for simple refreshing MultiListBox table of mails. Method used when button [Get Emails] is clicked.
    """
    root.winfo_toplevel().wm_geometry("")
    eMails = getSimpleEmails()
    # Clear the MultiListBox table from previous results
    mlb.delete(0, END)
    for eachMail in eMails:
        result = eMails[eachMail]
        mlb.insert(END,
                   (result[5], result[0], result[4], result[2], result[1], result[6], shorterString(result[3], 130)))
    connection.close()
    textWidget_insert("")


def getCurrentDate():
    """
    The method obtains a current date and changes its format to "DD MM YY".
    Helper methods: strLenAlign(), dateAlign()
    :return: String formatted "DD MM YY".
    """
    now = datetime.datetime.now()
    day = str(now.day)
    day = dateAlign(day)
    month = str(now.month)
    month = dateAlign(month)
    year = str(now.year)
    currentDate = day + "-" + month + "-" + year
    return currentDate


def strLenAlign(aString, anInt, aChar):
    """
    Helping method for getCurrentDate(). Works together with dateAlign() to deal with one digit numbers in data.
    :param aString:
    :param anInt:
    :param aChar:
    :return: String
    """
    if len(aString) == anInt:
        aString = aChar + aString
    return aString


def dateAlign(aString):
    """
    Helping method for getCurrentDate()
    The method uses the strLenAlign(aString, 1, '0') method to deal with the case when there is a single digit in
    getcurrentDate() method, so every time the format is the same - DD MM YY.
    :param aString:
    :return:
    """
    result = strLenAlign(aString, 1, '0')
    return result


def flagYellow():
    """
    Method changes the selected MultiListbox email from unFlagged (not starred) to Flagged (starred).
    INVARIANT: Must have item selected in multibox + press [Yellow Flag] button.
    Output: Selected Email is Flagged (starred).
    """
    try:
        selectedEmail = mlb.get(mlb.curselection(), last=None)
        allEmails = getSimpleEmails()
        selectedEmail = selectedEmail[2]  # <-- This is just ID like: b'3'
        result = allEmails[selectedEmail]
        final_result = result[4]
        connection.select('INBOX')
        status, data = connection.search(None, 'UNFLAGGED')
        unflaggedIds = data[0].split()
        for eachId in unflaggedIds:
            if final_result == eachId:
                connection.store(final_result, '+FLAGS', '\\Flagged')
        refreshSimple()

    except:
        messagebox.showwarning('Warning', 'Email not selected!' + '\n' + 'Please select an email you want to FLAG.')


def unFlagYellow():
    """
    Methods unFlags(not starred) the selected email(ticket) and refreshes the list of emails(tickets) in multilistbox.
    Output: Selected Email is Unflagged (not starred).
    :return: None
    """
    try:
        selectedEmail = mlb.get(mlb.curselection(), last=None)
        allEmails = getSimpleEmails()
        selectedEmail = selectedEmail[2]  # <-- This is just ID like: b'3'
        result = allEmails[selectedEmail]
        final_result = result[4]
        connection.select('INBOX')
        status, data = connection.search(None, 'FLAGGED')
        flaggedIds = data[0].split()
        for eachId in flaggedIds:
            if final_result == eachId:
                connection.store(final_result, '-FLAGS', '\\FLAGGED')
        refreshSimple()
    except:
        messagebox.showwarning('Warning', 'Email not selected!' + '\n' + 'Please select an email you want to UN-FLAG!')


def shorterString(aString, anInt):
    """
    Method makes passed String(aString) limited to specified number (anInt) of characters.
    :param aString: String to limit
    :param anInt: limit value for string length
    :return: String
    """
    temp = str(aString)
    result = ""
    if len(temp) >= anInt:
        for eachChar in range(anInt):
            result = result + temp[eachChar]
        result = result + "..."
        return result
    return temp


def getCurrentTime():
    """
    The method obtains the current time and then reformats it to HH:MM:SS.
    :return: String in format HH:MM:SS
    """
    result = str(datetime.datetime.now())
    result = result.split(" ")
    result = result[1]
    result = result.split(".")
    result = result[0]
    return result

def textWidget_insert(aString):
    """
    The method invoked when emails content(text) is displayed + info of refresh time.
    :param aString: email content
    :return: None
    """
    textWidget.config(state=NORMAL)
    textWidget.delete(1.0, END)
    lastAction = "\n" + "Last refresh: " + getCurrentTime()
    result = aString + lastAction
    textWidget.insert(END, result)
    textWidget.config(state=DISABLED)


def readMailText():
    """
    The method when executed displays in feedback window currently selected email(ticket) full description.
    INVARIANT: Email must be selecte din MulitListBox.
    :return: None
    """
    try:
        textWidget.delete(1.0, END)
        selected = mlb.get(mlb.curselection(), last=None)
        # refreshTable()
        mails = getSimpleEmails()
        text = mails[selected[2]]
        result = "Text from e-Mail titled " + str(text[0]) + " from " + str(text[2]) + ":" + "\n" + str(text[3])
        textWidget_insert(result)
    except:
        messagebox.showwarning('Warning',
                               'NO EMAIL SELECTED.' + '\n' + 'Please select an email you want to read text from.')


if __name__ == '__main__':
    user = 'hayaicode@gmail.com'
    password = getPasswordAsString()
    imap_url = 'imap.gmail.com'
    connection = imaplib.IMAP4_SSL(imap_url)
    connection.login(user, password)
    # Inbox is selected as a folder to search through at this point.
    connection.select('INBOX')
    result, data = connection.search(None, 'ALL')
    mail_ids = data[0].split()
    # Used for storing the attachments from emails.
    attachment_dir = os.getcwd()

"""
---------------> TKINTER INTERFACE <------------------------------------------------------------------------------------
"""


class MultiListbox(Frame):
    def __init__(self, master, lists):
        Frame.__init__(self, master)
        self.lists = []
        for l, w in lists:
            frame = Frame(self);
            frame.pack(side=LEFT, expand=YES, fill=BOTH)
            Label(frame, text=l, borderwidth=1, relief=RAISED).pack(fill=X)
            lb = Listbox(frame, width=w, borderwidth=0, selectborderwidth=0,
                         relief=FLAT, exportselection=FALSE)
            lb.pack(expand=YES, fill=BOTH)
            self.lists.append(lb)
            lb.bind('<B1-Motion>', lambda e, s=self: s._select(e.y))
            lb.bind('<Button-1>', lambda e, s=self: s._select(e.y))
            lb.bind('<Leave>', lambda e: 'break')
            lb.bind('<B2-Motion>', lambda e, s=self: s._b2motion(e.x, e.y))
            lb.bind('<Button-2>', lambda e, s=self: s._button2(e.x, e.y))
        frame = Frame(self);
        frame.pack(side=LEFT, fill=Y)
        Label(frame, borderwidth=1, relief=RAISED).pack(fill=BOTH)
        sb = Scrollbar(frame, orient=VERTICAL, command=self._scroll)
        sb.pack(expand=YES, fill=Y)
        self.lists[0]['yscrollcommand'] = sb.set

    def _select(self, y):
        row = self.lists[0].nearest(y)
        self.selection_clear(0, END)
        self.selection_set(row)
        return 'break'

    def _button2(self, x, y):
        for l in self.lists: l.scan_mark(x, y)
        return 'break'

    def _b2motion(self, x, y):
        for l in self.lists: l.scan_dragto(x, y)
        return 'break'

    def _scroll(self, *args):
        for l in self.lists:
            l.yview(*args)

    def curselection(self):
        return self.lists[0].curselection()

    def delete(self, first, last=None):
        for l in self.lists:
            l.delete(first, last)

    def get(self, first, last=None):
        result = []
        for l in self.lists:
            result.append(l.get(first, last))
        if last:
            return l.yview(map, [None] + result)
        return result

    def index(self, index):
        self.lists[0].index(index)

    def insert(self, index, *elements):
        for e in elements:
            i = 0
            for l in self.lists:
                l.insert(index, e[i])
                i = i + 1

    def size(self):
        return self.lists[0].size()

    def see(self, index):
        for l in self.lists:
            l.see(index)

    def selection_anchor(self, index):
        for l in self.lists:
            l.selection_anchor(index)

    def selection_clear(self, first, last=None):
        for l in self.lists:
            l.selection_clear(first, last)

    def selection_includes(self, index):
        return self.lists[0].selection_includes(index)

    def selection_set(self, first, last=None):
        for l in self.lists:
            l.selection_set(first, last)


root = Tk()
root.title("HelpDesk App 2019")
root.geometry('800x800')

# Creation of Frames inside Tkinter Root object, which will be populated with buttons.
topFrame = Frame(root)
middleFrame = Frame(root)
topBottomFrame = Frame(root)
bottomFrame = Frame(root)

topFrame.pack(side=TOP, anchor=N)
middleFrame.pack(fill=BOTH, expand=YES)
topBottomFrame.pack(fill=BOTH, anchor=N)
bottomFrame.pack(fill=BOTH, anchor=S)
refreshSimpleButton = Button(topFrame, text="Get Emails", command=refreshSimple)
refreshSimpleButton.pack(side=LEFT)

changeDirButton = Button(topFrame, text="Set Download folder", command=lambda: setAttachDir())
changeDirButton.pack(side=RIGHT)
yFlagButton = Button(topFrame, text="FLAG MAIL", command=flagYellow)
yFlagButton.pack(side=RIGHT)
unFlagButton = Button(topFrame, text="UN-FLAG", command=unFlagYellow)
unFlagButton.pack(side=RIGHT)
readButton = Button(topFrame, text="Read Email Text", command=readMailText)
readButton.pack()

"""
ListBox Here
"""
Label(middleFrame, text='MultiListbox').pack()
mlb = MultiListbox(middleFrame,
                   (('Status', 0), ('Subject', 0), ('ID', 0), ('From', 0), ('Date', 0), ('Time', 0),
                    ('Text preview', 0)))
mlb.pack(expand=YES, fill=BOTH)

textWidget = Text(topBottomFrame)
textWidget.insert(END, "Text will be displayed here")
textWidget.pack(expand=YES, fill=BOTH)
textWidget.config(state=DISABLED)

"""
Responsible for visibility of Tkinter GUI window.
"""
root.mainloop()

imap_url = 'imap.gmail.com'
connection = imaplib.IMAP4_SSL(imap_url)
connection.login(user, password)
# Inbox is selected as a folder to search through at this point.
connection.select('INBOX')
result, data = connection.search(None, 'ALL')
mail_ids = data[0].split()
# Used for storing the attachments from emails.
attachment_dir = os.getcwd()

"""

"""
def setAttachDir():
    currdir = os.getcwd()
    tempdir = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')
    if len(tempdir) > 0:
        print("You chose: %s" % tempdir)
    attachment_dir = tempdir

"""
Extracts the body of the email.
"""


def get_body(msg):
    # Check if email contains any attachments. Loops through all of the payloads and return the value when 1st is found.
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        # In case there's no multiple payloads this is returned.
        return msg.get_payload(None, True)


"""
INPUT:
Key: use TO or FROM
value: EMAIL adress here - the value want to check for.
con: connection used

OUTPUT:
List of numbers: in [b'1 2 ... n'] format
    EXAMPLE:
         print(search('FROM', 'kszumko@gmail.com', connection))
         gives:
                [b'1 3'] #List of index value of eMails that satisfies the value input.
"""


def search(key, value, con):
    result, data = con.search(None, key, '"{}"'.format(value))  # Value will be inserted into {}.
    return data


"""
Parse the found emails (result from search).
Initialise and empty list and appends emails to list which is returned.
INPUT:
search function results = uses the list of mails index values to fetch.

OUTPUT:
List of messages(msgs): Returns a list containing all found messages that 
                        were fetched using list of index(from search function).
EXAMPLE:
print(get_emails(search('FROM', 'kszumko@gmail.com', connection))) - 
"""


def get_emails(result_bytes):
    msgs = []
    for num in result_bytes[0].split():
        typ, data = connection.fetch(num, '(RFC822)')
        msgs.append(data)
    return msgs


"""
This function is checking email(msg) for attachments and storing it in previously specified folder.
INPUT:
msg: The message to search through
attachment_dir: Needs to be specified
OUTPUT:
The attachment file is stored in previously specified folder.
"""


def get_attachments(msg):
    # Walk() allows to go through content of message.
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        # Filename is obtained at this point.
        fileName = part.get_filename()

        # False means no attachment.
        if bool(fileName):
            filePath = os.path.join(attachment_dir,
                                    fileName)  # <--Joining of the filename and filepath previously specified.
            # With open automatically close the file after reading.
            with open(filePath, 'wb') as file:
                file.write(part.get_payload(decode=True))
"""
INPUT: None
Output: List of eMail IDs, which have status 'UNSEEN'
EXAMPLE: [b'1', b'2']
"""


def get_unseen_mails():
    result, data = connection.search(None, '(UNSEEN)')
    mail_ids = data[0].split()
    return mail_ids


"""
INPUT: None
Output: List of Ids of e-mails that are in mailbox 'Important', but are not in mailbox 'Inbox'.
"""


def get_archived_mails():
    result = []
    connection.select('[Gmail]/Important')
    status, data = connection.search(None, 'ALL')
    mail_ids_Important = data[0].split()
    connection.select('INBOX')
    status, data = connection.search(None, 'ALL')
    mail_ids_inbox = data[0].split()
    for eachId in mail_ids_Important:
        if eachId not in mail_ids_inbox:
            result.append(eachId)
    return result


"""
INPUT: None
OUTPUT: List of eMail IDs, which have active star (are starred).
"""
def get_starred_emails():
    result, data = connection.search(None, 'FLAGGED')
    mail_ids = data[0].split()
    return mail_ids


"""
INPUT: None
OUTPUT: List of email IDs, which has not been read yet(have 'UNFLAGGED' status).
"""


def get_unflagged_emails():
    result, data = connection.search(None, 'UNFLAGGED')
    mail_ids = data[0].split()
    return mail_ids


"""
Method for parsing integers values from String variable.
INPUT: aString
OUTPUT: Integer value parsed from aString (supplied argument)
"""


def onlyIntegersFromString(aString):
    result = int(re.search(r'\d+', aString).group())
    return result


"""
INPUT:
rawMail - the content of email parsed by IMAPlib(list or String).
strFind - a String variable to find in rawMail.
splitStr - a value that will be determine "string.replace()".
OUTPUT: Return - a LIST value that starts from 'strFind' value and is split by 'splitStr'.
"""


def getSpecificLine(rawMail, strFind, splitStr):
    temp = str(rawMail)
    tempInt = temp.rfind(strFind)
    temp = (temp[tempInt:len(temp)])
    temp = temp.split(splitStr)
    return temp

# Helping method for printing in more distinctive way (by adding long series of "-" before and after given argument)
def printCom(aString):
    c = ""
    for x in range(60): # Amount of "-" sign printed
        c += "-"
    print(c + "\n")
    print(aString)
    print("\n" + c)


"""
Method obtains eMails from selected Gmail account of MIME-type/regular and parse it for data.
OUTPUT: Return - listOfMails where KEY is mail ID (e.g. b'1') and 
value is parsed mail content (FROM, SUBJECT, DATE, TEXT).
"""


def getSimpleEmails():
    titleTag = "<titolo>"
    fromTag = "<sendinto>"
    toTag = "<sendu_al>"
    dateTag = "<nuna_dato>"
    timeTag = "<nuna_tempo>"
    textTag = "<interpretus>"
    # Step 1: Get usable List of eMails in Inbox.
    connection.select('INBOX')
    mail_text, data = connection.search(None, 'ALL')
    mail_ids = data[0].split()  # <-- Outputs a list of [b'1', b'2', b'3']
    """Step 2: Creation of main frame 
    (dictionary => key is mail ID, value => List - based on email -(FROM, SUBJECT, DATE, TEXT)) """
    listOfMails = dict.fromkeys(mail_ids)  # <-- Outputs {b'1': None, b'2': None, b'3': None}
    for eachKey in listOfMails:
        mail_text, data = connection.fetch(eachKey, '(RFC822)')
        rawMail = email.message_from_bytes(data[0][1])
        haveTime = False
        for eachPart in rawMail.walk():
            if eachPart.get_content_maintype() == 'multipart':
                parts = len(eachPart.get_payload())
                part = str(eachPart)
                try:
                    # Step 3: Obtaining Date from emails.
                    mail_data = part.split(dateTag)
                    mail_data = str(mail_data[1])
                except:
                    print("error")
                    mail_data = "--/--/--"
                # Step 4: Obtaining From(sender).
                try:
                    mail_From = part.split(fromTag)
                    mail_From = str(mail_From[1])
                    if not mail_From.rfind("gmail.com"):
                        mail_From = "Junk Mail"
                except:
                    print("Error with mail_from")
                    mail_From = "--"

                # Step 5: Obtaining Subject(error code).
                try:
                    mail_Subject = part.split(titleTag)
                    mail_Subject = str(mail_Subject[1])
                except:
                    print("error with subject")
                    mail_Subject = "--"
                # Step 6: Obtaining Text from email and parsing it.
                try:
                    mail_text = part.split(textTag)
                    mail_text = str(mail_text[1])
                except:
                    mail_text = "---"
                # Step 7: Obtaining status of e-mail.
                notStarredM = get_unflagged_emails()
                mail_status = ""
                if eachKey not in notStarredM:
                    mail_status = "SOLVING"
                if part.rfind(timeTag):
                    haveTime = True
                    try:
                        mail_time = part.split(timeTag)
                        mail_time = str(mail_time[1])
                    except:
                        mail_time = "--:--:--"
            if parts > 1:
                try:
                    pLoad = eachPart.get_payload()
                    pLoad = pLoad[1]
                    get_attachments(pLoad)
                except:
                    print("No image found!")
            # Step 8: Constructing the List which will be inserted into listOfMails as value for eachKey.
            list_of_values = []
            list_of_values.append(shorterString(mail_Subject, 30))
            list_of_values.append(shorterString(mail_data, 30))
            list_of_values.append(shorterString(mail_From, 30))
            list_of_values.append(mail_text)
            list_of_values.append(eachKey)
            list_of_values.append(mail_status)
            if haveTime:
                list_of_values.append(mail_time)
            else:
                list_of_values.append("-")
            printCom(str(list_of_values))
            listOfMails[eachKey] = list_of_values

    if not listOfMails == {}:
        return listOfMails
    else:
        messagebox.showwarning('Warning', 'Mailbox is empty.')


def refreshSimple():
    """
    Method used for simple refreshing MultiListBox table of mails.
    """
    root.winfo_toplevel().wm_geometry("")
    eMails = getSimpleEmails()
    # Clear the MultiListBox table from previous results
    mlb.delete(0, END)
    for eachMail in eMails:
        result = eMails[eachMail]
        mlb.insert(END, (result[5], result[0], result[4], result[2], result[1], result[6], shorterString(result[3], 130)))
    connection.close()
    textWidget_insert("")


def getCurrentDate():
    """
    INPUT: None
    OUTPUT: Return - single String variable in format "dd-mm-yy".
    """
    now = datetime.datetime.now()
    day = str(now.day)
    day = dateAlign(day)
    month = str(now.month)
    month= dateAlign(month)
    year = str(now.year)
    currentDate = day + "-" + month + "-" + year
    return currentDate

# Look UP
def strLenAlign(aString, anInt, aChar):
    if len(aString) == anInt:
        aString = aChar + aString
    return aString

# Look UP
def dateAlign(aString):
    result = strLenAlign(aString, 1, '0')
    return result

def flagYellow():
    """
    Method changes the selected MultiListbox email from unFlagged (not starred) to Flagged (starred).
    Input: Must have item selected in multibox + press [Yellow Flag] button.
    Output: Selected Email is Flagged (starred).
    """
    try:
        selectedEmail = mlb.get(mlb.curselection(), last=None)
        allEmails = getSimpleEmails()
        selectedEmail = selectedEmail[2]  # <-- This is just ID like: b'3'
        result = allEmails[selectedEmail]
        final_result = result[4]
        connection.select('INBOX')
        status, data = connection.search(None, 'UNFLAGGED')
        unflaggedIds = data[0].split()
        for eachId in unflaggedIds:
            if final_result == eachId:
                connection.store(final_result, '+FLAGS', '\\Flagged')
        refreshSimple()

    except:
        messagebox.showwarning('Warning', 'Email not selected!' + '\n' +'Please select an email you want to FLAG.')


def unFlagYellow():
    """
    Methods unFlags(not starred) the selected email(ticket) and refreshes the list of emails(tickets) in multilistbox.
    :return:
    """
    try:
        selectedEmail = mlb.get(mlb.curselection(), last=None)
        allEmails = getSimpleEmails()
        selectedEmail = selectedEmail[2]  # <-- This is just ID like: b'3'
        result = allEmails[selectedEmail]
        final_result = result[4]
        connection.select('INBOX')
        status, data = connection.search(None, 'FLAGGED')
        flaggedIds = data[0].split()
        for eachId in flaggedIds:
            if final_result == eachId:
                connection.store(final_result, '-FLAGS', '\\FLAGGED')
        refreshSimple()
    except:
        messagebox.showwarning('Warning', 'Email not selected!' + '\n' +'Please select an email you want to UN-FLAG!')

def shorterString(aString, anInt):
    """
    Method makes passed String value shorter than 230 characters.
    Input: Item from List, String
    Output: String
    """
    temp = str(aString)
    result = ""
    if len(temp) >= anInt:
        for eachChar in range(anInt):
            result = result + temp[eachChar]
        result = result + "..."
        return result
    return temp


def getCurrentTime():
    """
    Method reads a full text of email selected in MultiListBox frame. The text is
    displayed in textWidget (tkinter text widget in topBottomFrame).
    """
    result = str(datetime.datetime.now())
    result = result.split(" ")
    result = result[1]
    result = result.split(".")
    result = result[0]
    return result

def textWidget_insert(aString):
    textWidget.config(state=NORMAL)
    textWidget.delete(1.0, END)
    lastAction = "Last refresh:" + "\n" + getCurrentTime()
    result = aString + lastAction
    textWidget.insert(END, result)
    textWidget.config(state=DISABLED)

def readMailText():
    try:
        textWidget.delete(1.0, END)
        selected = mlb.get(mlb.curselection(), last=None)
        # refreshTable()
        mails = getSimpleEmails()
        text = mails[selected[2]]
        result = "Text from e-Mail titled " + str(text[0]) + " from " + str(text[2]) + ":" + "\n" + str(text[3])
        textWidget_insert(result)
    except:
        messagebox.showwarning('Warning',
                               'NO EMAIL SELECTED.' + '\n' + 'Please select an email you want to read text from.')


"""
---------------> TKINTER INTERFACE <------------------------------------------------------------------------------------
"""


class MultiListbox(Frame):
    def __init__(self, master, lists):
        Frame.__init__(self, master)
        self.lists = []
        for l, w in lists:
            frame = Frame(self);
            frame.pack(side=LEFT, expand=YES, fill=BOTH)
            Label(frame, text=l, borderwidth=1, relief=RAISED).pack(fill=X)
            lb = Listbox(frame, width=w, borderwidth=0, selectborderwidth=0,
                         relief=FLAT, exportselection=FALSE)
            lb.pack(expand=YES, fill=BOTH)
            self.lists.append(lb)
            lb.bind('<B1-Motion>', lambda e, s=self: s._select(e.y))
            lb.bind('<Button-1>', lambda e, s=self: s._select(e.y))
            lb.bind('<Leave>', lambda e: 'break')
            lb.bind('<B2-Motion>', lambda e, s=self: s._b2motion(e.x, e.y))
            lb.bind('<Button-2>', lambda e, s=self: s._button2(e.x, e.y))
        frame = Frame(self);
        frame.pack(side=LEFT, fill=Y)
        Label(frame, borderwidth=1, relief=RAISED).pack(fill=BOTH)
        sb = Scrollbar(frame, orient=VERTICAL, command=self._scroll)
        sb.pack(expand=YES, fill=Y)
        self.lists[0]['yscrollcommand'] = sb.set

    def _select(self, y):
        row = self.lists[0].nearest(y)
        self.selection_clear(0, END)
        self.selection_set(row)
        return 'break'

    def _button2(self, x, y):
        for l in self.lists: l.scan_mark(x, y)
        return 'break'

    def _b2motion(self, x, y):
        for l in self.lists: l.scan_dragto(x, y)
        return 'break'

    def _scroll(self, *args):
        for l in self.lists:
            l.yview(*args)

    def curselection(self):
        return self.lists[0].curselection()

    def delete(self, first, last=None):
        for l in self.lists:
            l.delete(first, last)

    def get(self, first, last=None):
        result = []
        for l in self.lists:
            result.append(l.get(first, last))
        if last:
            return l.yview(map, [None] + result)
        return result

    def index(self, index):
        self.lists[0].index(index)

    def insert(self, index, *elements):
        for e in elements:
            i = 0
            for l in self.lists:
                l.insert(index, e[i])
                i = i + 1

    def size(self):
        return self.lists[0].size()

    def see(self, index):
        for l in self.lists:
            l.see(index)

    def selection_anchor(self, index):
        for l in self.lists:
            l.selection_anchor(index)

    def selection_clear(self, first, last=None):
        for l in self.lists:
            l.selection_clear(first, last)

    def selection_includes(self, index):
        return self.lists[0].selection_includes(index)

    def selection_set(self, first, last=None):
        for l in self.lists:
            l.selection_set(first, last)


root = Tk()
root.title("HelpDesk App 2019")
root.geometry('800x800')

# Creation of Frames inside Tkinter Root object, which will be populated with buttons.
topFrame = Frame(root)
middleFrame = Frame(root)
topBottomFrame = Frame(root)

topFrame = Frame(root)
middleFrame = Frame(root)
topBottomFrame = Frame(root)
bottomFrame = Frame(root)

topFrame.pack(side=TOP, anchor=N)
middleFrame.pack(fill=BOTH, expand=YES)
topBottomFrame.pack(fill=BOTH, anchor=N)
bottomFrame.pack(fill=BOTH, anchor=S)
refreshSimpleButton = Button(topFrame, text="Get Emails", command=refreshSimple)
refreshSimpleButton.pack(side=LEFT)

changeDirButton = Button(topFrame, text="Set Download folder", command=lambda: setAttachDir())
changeDirButton.pack(side=RIGHT)
yFlagButton = Button(topFrame, text="FLAG MAIL", command=flagYellow)
yFlagButton.pack(side=RIGHT)
unFlagButton = Button(topFrame, text="UN-FLAG", command=unFlagYellow)
unFlagButton.pack(side=RIGHT)
readButton = Button(topFrame, text="Read Email Text", command=readMailText)
readButton.pack()

"""
ListBox Here
"""
Label(middleFrame, text='MultiListbox').pack()
mlb = MultiListbox(middleFrame,
                   (('Status', 0), ('Subject', 0), ('ID', 0), ('From', 0), ('Date', 0), ('Time', 0), ('Text preview', 0)))
mlb.pack(expand=YES, fill=BOTH)

textWidget = Text(topBottomFrame)
textWidget.insert(END, "Text will be displayed here")
textWidget.pack(expand=YES, fill=BOTH)
textWidget.config(state=DISABLED)

"""
Responsible for visibility of Tkinter GUI window.
"""
root.mainloop()
