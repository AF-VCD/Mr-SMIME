# Mr-SMIME

This is a PowerShell script that can be used to manually decrypt email messages with a smartcard on Windows. It is intended for Air Force employees, but may also be usable by other people in government service.


## Instructions
In order to utilize this tool, follow the steps in the two sections below: 

### Getting S/MIME File:
1. Go to your webmail and find the offending encrypted email that you can't open. 
1. You should see that the email has an attachment called "smime.p7m"
1. Download that attachment somewhere to your computer. This is the encrypted version of the email you can't open, and includes attachments as well.

### Getting and using script:
1. Download the script from this github repository
1. Right-click on the script and choose "Run with PowerShell"
1. On the first file selection window, select the "smime.p7m" file you downloaded earlier
1. Enter your pin when prompted
1. On the second file window, choose a location to save the raw decrypted text output, such as DECRYPTED-MESSAGE.txt
1. Wait for the script to complete

The script produces two items upon completion. The first is a text file, with its name being what you specified in the second file window. This text file contains the raw text output of your email after decryption. You can glean some email message text from this file, but not attachments, because they are further encoded in there.

The second, more useful item is a folder with the same name as the raw text file (i.e. DECRYPTED-MESSAGE/). The folder will be located in the same location as the text file. Inside this folder are text files associated with the email message body, as well as any attachments included with the email.

## Notes on implementation and usage
a. This only works on Windows, tested on Windows 10. In order for this to work, your computer needs to at least be able to access webmail (have correct certificates, etc).
a. My script does not see anything related to your card, like your PIN; that part is handled by Windows and/or ActivClient. The script tells Windows to decrypt the email, and Windows figures out that it can't decrypt the email without the PIN, and asks you for it on its own.
a. No information whatsoever is transmitted from this script to anywhere on the internet.
