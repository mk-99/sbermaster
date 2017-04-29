# sbermaster
Creates bank statements for Sberbank

Sberbank (https://www.sberbank.ru) has quite good electronic/mobile bank clients but it's bank statements are terrible.
The best way to obtain reasonable financial information is to read SMS messages sent by Sberbank on every operation
with credit card or in electronic/mobile bank.

**sbermaster** works with SMS Backup+ Android application
(https://play.google.com/store/apps/details?id=com.zegoggles.smssync&hl=en)

It reads SMS messages from IMAP server (gmail.com is used by default), parses them and generates bank statement
as Excel 2010 table (.xlsx) file.

Python 3 is required (maybe it works with Python 2, but it's not tested).

    usage: sbermaster.py [-h] [-d DATE] [-l LOGIN] [-p PASSWORD] [-s SITE]
                         [-f FOLDER] [-S SEARCH] [-w]
                         outfile
    
    Process Sberbank SMS messages backed up to imap server and generatexlsx sheet
    
    positional arguments:
      outfile               Output MS Excel file, please add .xlsx explicitly
    
    optional arguments:
      -h, --help            show this help message and exit
      -d DATE, --date DATE  Start from date (default: 1-Mar-2017)
      -l LOGIN, --login LOGIN
                            Login with this name (default: None)
      -p PASSWORD, --password PASSWORD
                            Login with this password (default: None)
      -s SITE, --site SITE  Connect to this imap server (default: imap.gmail.com)
      -f FOLDER, --folder FOLDER
                            Folder to read SMS from (default: SMS)
      -S SEARCH, --search SEARCH
                            IMAP search string (default: FROM 900)
      -w, --warn            print warnings (default: False)
