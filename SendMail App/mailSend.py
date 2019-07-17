"""
CREATOR:    Krzysztof Szumko
LINKED-IN:  https://www.linkedin.com/in/krzysztofszumko1989
GITHUB:     https://github.com/NuDron
Version:    0.09
"""

import smtplib, ssl, datetime, trans, os, time
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from tkinter import *
from tkinter import messagebox
import tkinter.filedialog as filedialog
import unicodedata
from cryptography.fernet import Fernet


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


def getPassAsString():
    """
    Stringifies the password returned by getSavedPass method.
    :return: String
    """
    return str(getSavedPass().decode())


def getKey():
    """
    Returns a key used for encryption/decryption (cryptography fernet) from file present in the same directory as
    executable (file named key.key).
    :return: encryption key
    """
    file = open('key.key', 'rb')
    key = file.read()
    file.close()
    result = Fernet(key)
    return result


def getProblemID():
    """
    Based on the value stored in file IDs.txt, this method obtains currently stored ID int value.
    :return: Stringified version of ID number
    """
    # Open the file from within the directory named IDs in text format.
    f = open("IDs.txt", "r")
    id = f.readlines()
    f.close()
    # If there is no ID (it is the first run) then create a text file and initialise the value to 1.
    if id == None:
        f = open("IDs.txt", "w")
        newID = 1
        f.write(str(newID))
        f.close()
        # Run method again to read the newly created value.
        getProblemID()
    else:
        return str(id[0])


def upProblemID():
    """
    Method changes the value of ID stored in IDs.txt file by currentValue + 1.
    :return: None
    """
    fileName = "IDs.txt"
    # Only reading the ID value from IDs.txt file.
    f = open(fileName, "r")
    id = f.readlines()
    f.close()

    # Creating a new ID value based on the previous ID.
    newID = int(id[0]) + 1
    # Writing new ID value into IDs.txt file.
    f = open(fileName, "w")
    f.write(str(newID))
    f.close()


def setProblemID(anInt):
    """
    The method is used to set the value of ID (stored in IDs.txt) to a specified value - anInt.
    :param anInt:
    :return:
    """
    # Open the file from within directory.
    fileName = "IDs.txt"
    f = open(fileName, "w")
    # Set the ID value to provided anInt.
    f.write(str(anInt))
    f.close()

def normalize_char(c):
    """
    HELPER METHOD - see normalize(aString) method.
    :param c:
    :return:
    """
    try:
        cname = unicodedata.name(c)
        cname = cname[:cname.index(' WITH')]
        return unicodedata.lookup(cname)
    except (ValueError, KeyError):
        return c


def normalize(aString):
    """
    The method gets rid of unregular characters from aString.
    :param aString:
    :return: String
    """
    return ''.join(normalize_char(character) for character in aString)


def getCurrentTime():
    """
    Method outputs current time in "13:46:58" format.
    :return: current time formatted HH:MM:SS
    """
    result = str(datetime.datetime.now())
    result = result.split(" ")
    result = result[1]
    result = result.split(".")
    result = result[0]
    return result


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
    currentDate = day + " " + month + " " + year
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


def getImage():
    """
    The method for TKinter prompt to obtain image path of file selected by user.
    :return: a file path to image
    """
    Tk().withdraw()
    filename = filedialog.askopenfilename(initialdir="/", title="Select file", filetypes=(
    ('Image files', '*.jpg;*.jpeg;*.jpe;*.jfif;*.png;*.bmp'), ("all files", "*.*")))
    return filename


class Email_to_send:
    """
    Class of object used for storing and formatting information used to compose a ticket(email).
    """
    def __init__(self):
        self.uberTag = "<interpretus>"
        self.mainTitle = ""
        self.message = MIMEMultipart("alternative")

    def set_title(self, aString):
        """
        Standard setter method for class.
        :param aString:
        :return: None
        """
        self.mainTitle = aString

    def get_title(self):
        """
        Standard getter method for mainTitle variable.
        :return:
        """
        return self.mainTitle

    def get_text_tag(self):
        """
        Standard getter method for obtaining value of tag called uberTag.
        :return: String
        """
        return self.uberTag

    def set_message_text(self):
        """
        Setter method for the ticket(email) message so it is in the correct format
        for it to be parsed by the other end properly.

        Example:
        <interpretus>"retrievedInputString"<interpretus>
        :return: None
        """
        inputText = self.uberTag + retrieve_input() + self.uberTag
        self.message = MIMEText(inputText)

    def get_message_text(self):
        """
        Getter method for object message (email or problem description content).
        :return: String
        """
        return self.message


def attachOrNot():
    """
    The method is invoked to decide whether the user wants to attach the
    selected (by user) image (returns TRUE) or not (returns FALSE).
    :return: Boolean
    """
    result = messagebox.askquestion("Question", "Do you want to add image attachment?", icon='info')
    if result == 'yes':
        return True
    else:
        return False


def tagAString(aTag, aStr):
    """
    The method adds a HTML like tag (aTag) to a String value (aStr).

    Example:
    tagAString("myTag", "This is normal text") -- produces --> "<myTag>This is a normal text<myTag>"
    :param aTag:
    :param aStr:
    :return: String
    """
    fullTag = "<" + aTag + ">"
    result = fullTag + aStr + fullTag
    return result


def getFileName(aPath):
    """
    TODO
    :param aPath:
    :return:
    """
    aPath = aPath.split("/")
    aPath = aPath[-1]
    return aPath


def sendMail(aTitle, aSender, aReceiver):
    """
    The method gathers the information provided by a user like "category of issue", feedback
    (from problem description window) and optionally image.
    Pieces of information gathered are reformatted with HTML like tags into one eMail and sent
    if all requirements are satisfied.
    Appropriate feedback is given to the user upon "ticket" being sent.

    Notes:
    - Tags are from the Esperanto language created by Ludwik Zamenhof.
    - Tags are used for selective slicing of text and as a simple security check.
    :param aTitle:
    :param aSender:
    :param aReceiver:
    :return:
    """
    message = MIMEMultipart()
    message["Subject"] = tagAString("titolo", aTitle)               # <-- title from text widget.
    message["From"] = tagAString("sendinto", aSender)               # <-- email address of sender.
    message["To"] = tagAString("sendu_al", aReceiver)               # <-- receiver email address.
    message["Date"] = tagAString("nuna_dato", getCurrentDate())     # <-- Date when mail is sent.

    # Turn text retrieved from "Problem description" field into plain/html MIMEText objects with proper tags.
    inputText = tagAString(uberTag, retrieve_input())
    # Creation of standardDescription for simple comparison if user provided any feedback.
    standardDescription = ""
    standardDescription = tagAString(uberTag, standardDescription)
    # Check if user provided any kind of feedback in the feedback field.
    if inputText == standardDescription:
        messagebox.showwarning('Warning', 'Need to provide feedback!')
        feedbackRefresh("Feedback not provided. Please write something about encountered issue.")
        # Method should terminate at this stage to avoid sending tickets without feedback.
        return
    inputText += tagAString("nuna_tempo", getCurrentTime())
    # First filtration for Cyrillic signs (replacing them with regular chars).
    filteredText = inputText.translate(trans)

    # Second filtration for irregular signs (removing them from message).
    filteredText = filteredText.encode("ascii", errors="ignore").decode()
    part1 = MIMEText(filteredText, 'plain')

    # The email client will try to render the last part of message first.
    message.attach(part1)
    if attachOrNot():
        path = getImage()
        filename = getFileName(path)
        attachment = open(path, "rb")

        # instance of MIMEBase and named as p.
        p = MIMEBase('application', 'octet-stream')
        # To change the payload into encoded form.
        p.set_payload((attachment).read())
        # encode into base64.
        encoders.encode_base64(p)
        p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        message.attach(p)
    # Create secure connection with server and send email.
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(aSender, password)
        server.sendmail(
            aSender, aReceiver, message.as_string()
        )
        server.quit()
        upProblemID()
    # Prepare the feedback information for user.
    newMessage = inputText.replace("<interpretus>", "\n", 2)
    # Prepping time information.
    newMessage = newMessage.replace("<nuna_tempo>", "Was sent at: ", 1)
    newMessage = newMessage.replace("<nuna_tempo>", "", 1)
    # Assemble of entire message meant for feedback window.
    answer = "Message sent successfully:" + "\n" + "Title: " + aTitle + "\n" + "Problem description: " + newMessage + "\n" + "Successfully sent email"
    # Display the feedback information for user in the feedback window.
    feedbackRefresh(answer)


def sentTicket():
    """
    Send a ticket by using a sendMail method with previously specified sender_mail.
    :return: None
    """
    sendMail(toSend.get_title(), sender_email, sender_email)


def retrieve_input():
    """
    The methods retrieves the input provided by the user into "Problem Description" field.
    :return: String
    """
    input = textWidget.get("1.0", 'end-1c')
    return input


def setTitle(aString):
    """
    The method creates a suitable title used later for tickets.
    Formatted "aString-ID".
    :param aString:
    :return: None
    """
    result = aString + "-" + getProblemID()
    toSend.set_title(result)
    getTitle()


def getTitle():
    """
    The method provides feedback for a user about current ticket's selected title (category of the problem).
    Depends on previously created toSend object of class Email_to_send.
    :return: None
    """
    answer = "Category of problem selected is:" + "\n" + toSend.get_title()
    feedbackRefresh(answer)


def addStringWithNL(targetString, addedString):
    """
    HELPER METHOD
    The method combines two strings provided by adding " " between and newline after.
    :param targetString:
    :param addedString:
    :return: String
    """
    newString = targetString + " " + addedString
    newString = newString + "\n"
    return newString


def feedbackRefresh(aString):
    """
    The method used for refreshing the text in the Tkinter "feedback widget" window.
    The method enables text input, changes the text in the feedback window for aString and
    disables editing the text window shortly after.
    :param aString:
    :return: None
    """
    feedbackWidget.config(state=NORMAL)
    feedbackWidget.delete(1.0, END)
    feedbackWidget.insert(END, aString)
    feedbackWidget.config(state=DISABLED)

# Initialisation of essential variables. Executed only when run as main(program).
if __name__ == '__main__':
    # Gmail account used for sending tickets.
    sender_email = "hayaicode@gmail.com"
    # The initialisation of the mainTitle variable. Used for checking if the user has clicked any category.
    mainTitle = ""
    uberTag = "interpretus"
    # Obtaining the Gmail password from txt file.
    password = getPassAsString()
    port = 587
    # Dictionary of letters used for the string normalisation process.
    letters = {'ł': 'l', 'ą': 'a', 'ń': 'n', 'ć': 'c', 'ó': 'o', 'ę': 'e', 'ś': 's', 'ź': 'z', 'ż': 'z'}
    trans = str.maketrans(letters)
    # Used for lookup of preferred format of images.
    ftypes = [
        ('Image files', '*.jpg;*.jpeg;*.jpe;*.jfif;*.png;*.bmp'),
        ('PNG', '*.png'),
        ('JPEG', '*.jpg;*.jpeg;*.jpe;*.jfif'),  # semicolon trick
    ]
    should_attach = False
    # Initialisation of object of class Email_to_send, which will be used in other methods
    toSend = Email_to_send()

"""
-----------TKINTER INTERFACE -------------------------------------------------------------------------------------------
"""
# Initialising main Tkinter window.
root = Tk()
# Setting the title of the interface window.
root.title("HelpDesk App 2019")
# Setting specific window size.
root.geometry('800x800')

# Creating distinct subframes hooked to the root window.
first_top_frame = Frame(root)
first_middle_frame = Frame(root)
first_bottom_frame = Frame(root)
second_top_frame = Frame(root)

# Making frames visible to user. Setting how they fill the space.
first_top_frame.pack(side=TOP, anchor=N)
second_top_frame.pack(side=TOP, expand=YES, fill=BOTH)
first_middle_frame.pack(fill=BOTH, expand=YES)
first_bottom_frame.pack(fill=X, expand=False)

# Categories Buttons - this will determine the issue category (used as a title in ticket).
category1Butt = Button(first_top_frame, text="Category 1 Issue", command=lambda: setTitle("Problem_1"))
category2Butt = Button(first_top_frame, text="Category 2 Issue", command=lambda: setTitle("Problem_2"))
category3Butt = Button(first_top_frame, text="Category 3 Issue", command=lambda: setTitle("Problem_3"))
category4Butt = Button(first_top_frame, text="Category 4 Issue", command=lambda: setTitle("Problem_4"))
category5Butt = Button(first_top_frame, text="Category 5 Issue", command=lambda: setTitle("Problem_5"))
category6Butt = Button(first_top_frame, text="Category 6 Issue", command=lambda: setTitle("Problem_6"))

# Buttons made visible
category1Butt.pack(side=LEFT)
category2Butt.pack(side=LEFT)
category3Butt.pack(side=LEFT)
category4Butt.pack(side=LEFT)
category5Butt.pack(side=LEFT)
category6Butt.pack(side=LEFT)

# Screenshots attaching button - located below categories.
attachButton = Button(second_top_frame, text="Attach" + "\n" + "screenshot",
                      command=lambda: getImage())  # TODO: UNUSED BUTTON BELOW PROBLEM CATEGORIES.

# Text widget for writing text.
descriptionLab = Label(first_middle_frame, text="Problem Description:")
descriptionLab.pack(side=TOP, anchor=NW)
textWidget = Text(first_middle_frame)
textWidget.pack(expand=YES, fill=BOTH)

feedbackLab = Label(first_bottom_frame, text="Action Feedback")
feedbackLab.pack(side=TOP, anchor=NW)
sendButt = Button(first_bottom_frame, text="Send Ticket", command=lambda: sentTicket())
sendButt.pack(fill=BOTH, side=RIGHT, anchor=SW)
feedbackWidget = Text(first_bottom_frame)
feedbackWidget.pack(expand=None, fill=X)
feedbackWidget.insert(END, "Feedback of actions taken will be displayed in this section")
feedbackWidget.config(state=DISABLED)

"""
Responsible for visibility of Tkinter GUI window.
"""
root.mainloop()
