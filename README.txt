SMARTEDGE WHATSAPP ASSISTANT — WINDOWS 7 SETUP

IMPORTANT
This is a manual-send assistant, not an automatic bulk sender. It opens an opted-in
contact, prepares the selected image and caption, and waits for you to review and
click Send yourself. It does not use randomised anti-detection delays, hide automation,
or attempt to bypass WhatsApp safeguards. No method can guarantee an account will
never be restricted.

1) SYSTEM
- Windows 7
- Python 3.8.10 (legacy / end-of-life; upgrade Windows when possible)
- Microsoft Excel 2007
- Chrome version compatible with the ChromeDriver you install

2) FOLDER CONTENTS
- main.py: assistant
- students_contacts.xlsx: your workbook (create using the template below)
- POSTS: put JPG/PNG/WEBP posters here
- chromedriver.exe: place the driver matching your installed Chrome version here
- whatsapp_chrome_profile: created on first run

3) EXCEL COLUMNS (row 1 headers)
A: Name
B: WhatsApp_Number
C: Status
D: Last_Updated
E: Notes
F: Opt_In

Enter YES in column F only for contacts who explicitly agreed to receive these
WhatsApp messages. Use Indian 10-digit mobile numbers or 91 + 10 digits.
Close Excel before running the script.

4) INSTALL
Install Python 3.8.10 and tick "Add Python to PATH".
Open Command Prompt in this folder and run:
    python -m pip install --upgrade "pip<25"
    python -m pip install -r requirements.txt

If pip cannot install a package because Windows 7 is unsupported, do not force
random newer versions. Use a supported computer / upgraded Windows instead.

5) CHROMEDRIVER
Selenium 4.9.1 may not reliably obtain drivers automatically on old Windows.
Download a ChromeDriver that exactly matches your installed Chrome major version,
then place chromedriver.exe beside main.py. Do not use a random driver version.

6) RUN
Double-click RUN_BOT.bat or run:
    python main.py

Choose a poster from POSTS. Scan the WhatsApp Web QR code if prompted. Review
each recipient, poster and caption. Click Send manually in WhatsApp Web, then
return to the console and enter S. Enter R if unsure, K to skip, Q to stop.

7) EXCEL STATUS
Sent = you confirmed you clicked Send; the script does not independently verify it.
Review = check WhatsApp manually before retrying.
Failed = script could not prepare the chat/media.
Skipped = you skipped that contact.
Preview OK = preview prepared in DRY_RUN mode.

Do not send unsolicited messages. Honour opt-outs immediately. Keep campaign
volumes reasonable and follow WhatsApp Business Messaging Policy and local law.
