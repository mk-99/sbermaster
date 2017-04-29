# sbermaster
Creates bank statements for Sberbank

Sberbank (https://www.sberbank.ru) has quite good electronic/mobile bank clients but it's bank statements are terrible.
The best way to obtain reasonable financial information is to read SMS messages sent by Sberbank on every operation
with credit card or in electronic/mobile bank.

**sbermaster** works with SMS Backup+ Android application (https://play.google.com/store/apps/details?id=com.zegoggles.smssync&hl=en)

It reads SMS messages from IMAP server (gmail.com is used by default), parses them and generates bank statement as Excel 2010 table (.xlsx) file.

Python 3 is required (maybe it works with Python 2, but it's not tested).
