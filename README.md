# BotMail
Automation of email composition and mass sending from personal email - will not register as spam. Email access is TLS encrypted.

Supports attachments, dynamic names, and emphases (bold, color, etc.)

**Currently in process of developing user interface using PyQt** 
**Error codes will be documented**

Use:
**Must enable less secure app access for your email account in order to log in using BotMail (Active TLS Encryption)**
(1) Create list of recipients as Excel file with columns of [CONTACT NAME, EMAIL ADDRESS, COMPANY] **do not name columns**
(2) Excel file must be in PATH
(3) Edit attachment name in BotMail.py for all instances (see notes)
(4) Run program and enter Excel file name
(5) Check table to make sure all contact entries are accurate
(6) Log into email address and fill out email
(7) A test email will first be sent to the logged in email address. Check to make sure everything looks right. In this test email, the dynamic contact name will appear as your email address, NOT your name - in the emails that will be sent to contact list, these will be replaced with CONTACT NAME. Dynamic company names will appear as 'COMPANYNAMEWILLBEHERE' and be replaced by COMPANY in emails that will be sent out.
(8) Confirm and send emails. Wait for success message and for program to self-terminate.

Sending speeds: **Under testing. Will be updated**
For emails with attachment payloads of <1 MB: ~6 emails per second

Max emails: **Under testing. Will be updated**
Depending on the size of the email (with attachment), your email provider may place a cap on the amount of emails sent in a short time period. 
**If first time using program, do not send more than 150 emails of >1.5MB size within a 2 hour period. Google WILL DISABLE your account temporarily for suspicious activity (ranges 2-24 hr suspension). Once less-secure app access has been turned on and program has been used for at least 3 days, then this max capacity can be expanded.**

Notes:
- An 'undo' feature and 'back' button will be added in future versions. 
- An option to edit contact entries in the table directly will be added in future versions. This will also update the Excel file.
- Email text body is currently inserted paragraph by paragraph. In future versions, there will be features to (1) type whole email body in program with all typography emphases features and (2) 'browse' for word document (.doc or .docx) with prewritten email and automatically parse and convert to HTML script. In-program text editing also supported for this option.
- Currently only supports Gmail, Yahoo, Outlook, and Hotmail. Support for more domains will be added in future versions.
- Excel file (.xlsx) needs to be manually typed as user input. A 'browse' option to locate file and auto-generate directory path will be added in future versions.
- Attachment payloads currently need to be manually loaded by editing BotMail.py. A 'browse' option to open folders and auto-generate directory path will be added in future versions.  
- Custom personal email signatures currently not supported.
