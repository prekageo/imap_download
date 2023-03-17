# IMAP download

This tool downloads all emails from an IMAP server and stores each email in an individual file (similar to the Maildir format but not exactly the same). It sets the modification time of each file to match the email timestamp.

## Instructions

Modify the script and set the variables `host`, `user`, and `destination`:
* Host is the IMAP server, e.g. `export.imap.mail.yahoo.com` for Yahoo! Mail.
* User is the username, e.g. your email address for Yahoo! Mail.
* Destination is the local file path where to store your emails.
