# BankEmailScraper
For anyone who receives charge notifications via gmail and wants to connect that to a google sheets or other excel sheet for tracking (low risk of data leaks vs budget apps)


#Code
Code.gs goes into scripts.google.com in order to execute

It must have access to your email

Note: 
- Still works for the 03/23 format of Chase Bank payment alert emails
- Must change 'from email' and spreadsheet link in Code.gs
- Must have a google sheets with a tab called 'Transactions' in order for the script to write to that sheet
