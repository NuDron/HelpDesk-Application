This is ReadMe for [GmailReader App] part of [Help Desk App 2019 ver 0.07]

Table of Contents:
I)    QUICK START
II)   NECESSARY FILES TO RUN
III)  TO DO LIST
IV)   CHANGELOG

1) QUICK START:
Step1) START GMail_Reader(Main).exe
Step2) Click [Get Emails]

Overview of App buttons (starting from the top):
Top Buttons(from left):
- [Get Emails] - the main button, as it populates the Multilistbox with emails from the predefined mailbox. Will give a warning if a mailbox is empty. Populates the MultilistBox only after receiving all of the emails in the mailbox - it can take a while depending on internet connection and amount of emails.
- [Read Email Text] - Reads a full text of "selected" email - text preview is limited to 130 characters.
- [UN-FLAG] - takes off the [FLAG MAIL] from "selected" email. It provides means of recuperating from accidentally flagging some email.
- [FLAG MAIL] - flags "selected" email. Changes status to "SOLVING" - currently ticket sender cannot check mail status.
- [Set Download folder] - currently does not work properly - work in progress.

MultiListBox:
- [Status] - can be empty or "SOLVING" (toggled by [FLAG MAIL] & [UN-FLAG])
- [Subject] - category problem chosen by ticket sender with the addition of unique ID
- [ID] - ID of mail parsed from GMAIL. Shows how GMAIL registered the order of incoming emails.
- [From] - emails address of ticket sender(s).
- [Date] - Date parsed from an incoming ticket.
- [Text preview] - Preview of problem description shorten to 130 characters.

Feedback Window:
In this window feedback regarding user's action is displayed.

2) NECESSARY FILES TO RUN:
- key.key - cryptography.fernet module key used for encryption
- config.txt - file with email password stored (encrypted with key.key)
+ all of the import from the top

3) TO DO LIST:
- coming soon!!!

4) CHANGELOG:
17-07-2019 - Added better documentation for methods (of " and # type). Move necessary initialisation into __main__ clause.

( ･_･) PLEASE NO CODE STEALING (･_･ )

THIS IS ALL  ┑(￣▽￣)┍  SO GOOD DAY TO YOU!
Krzysztof Szumko









