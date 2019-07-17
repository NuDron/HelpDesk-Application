This is ReadMe for [MailSend] part of [Help Desk App 2019 ver 0.09]

Table of Contents:
I)    Project Introduction
II)   Simple project structure
III)  Essentials to run a SendMail app
IV)   To do list
V)    Changelog

1) Project Introduction:
This is a two-piece application meant for simple sending and receiving tickets related to technical issues. 
The main concept is to use Gmail for sending and storing information used by both ends (separate apps). 

2) Simple project structure
Applications relations:
SendMail app        ---->         Gmail account           ---->           GmailReader app 

Example actions:
Preparing ticket   --sending-->  ticket stored as email   --reading-->    Tech-team app responds to the ticket

3) Essentials to run a SendMail app
- key.key - cryptography.fernet module key used for encryption
- config.txt - file with email password stored (encrypted with key.key)
- IDs.txt - simple text file with only a number. Used to generate an ID for each ticket sent. 
Essential for Gmail not to couple emails accidentally into one thread.
+ all of the import included

4) To Do List:
- Add functionality: "Check ticket status" - option for the client(SendMail) to check the sent ticket status (NONE, YELLOW OR GREEN).
- Add functionality: "Accept or Reject solution" - way of rejecting or accepting a solution proposed by tech team (app on another end).
- Add functionality: "Add/change the comment to already existing ticket" - way of adding or changing the comments to an already existing ticket. 
- Make tickets as threads - so multiple messages can be included in the same thread (one ticket). 
- Add Unit testing where necessary - make docstrings unit tests for better error proofing.

5) Changelog
16-07-2019 - Added documentation to most methods(" type comments). Checked again all of the notes (# type notes).
