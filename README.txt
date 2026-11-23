=============================================================================
SMARTEDGE WHATSAPP ASSISTANT (WINDOWS 7 & MS OFFICE 2007)
=============================================================================

1. SYSTEM COMPATIBILITY
-----------------------------------------------------------------------------
- OS: Windows 7 (SP1 recommended)
- Python: 3.8.10 (Add Python to PATH during installation)
- Browser: Google Chrome 109 (the final Chrome version for Windows 7)
- Driver: chromedriver.exe (Version 109 included in this folder)
- Spreadsheet: Microsoft Excel 2007 or newer (students_contacts.xlsx)


2. FOLDER OVERVIEW
-----------------------------------------------------------------------------
- main.py                  : The core Python script
- message.txt              : Message template (open and edit in Notepad!)
- POSTS/                   : Folder containing your promotional images / posters
- students_contacts.xlsx   : Contact list workbook (Excel 2007 format)
- config.json              : Automation, delay, and batching settings
- chromedriver.exe         : Matching ChromeDriver for Windows 7 Chrome 109
- RUN_BOT.bat              : One-click launcher to run the assistant
- whatsapp_chrome_profile/ : Saved login session (QR code scan only needed once)
- whatsapp_debug/          : Screenshots saved automatically if any error occurs


3. HOW TO USE (EASY 3-STEP WORKFLOW)
-----------------------------------------------------------------------------
Step 1: Edit your Message
- Double-click "message.txt" in Notepad.
- Type your message text.
- Use {Name} where you want each student's name to appear (e.g. "Hi {Name}!").
- Save and close Notepad.

Step 2: Add Poster Image
- Put your image (JPG, PNG, or WEBP) into the "POSTS" folder.
- The script automatically selects your image and attaches it with the caption.

Step 3: Update Contacts in Excel
- Open "students_contacts.xlsx" in Microsoft Excel 2007.
- Columns:
    Column A: Name (e.g. Sahil)
    Column B: WhatsApp_Number (10 digits or 91 + 10 digits)
    Column C: Status (Leave blank or Pending; script sets Sent/Failed)
    Column D: Last_Updated (Auto-filled by script)
    Column E: Notes (Auto-filled by script)
    Column F: Opt_In (Set to YES for contacts who agreed to be messaged)
- IMPORTANT: Save and close Excel before running the script!

Step 4: Launch
- Double-click "RUN_BOT.bat" or open cmd and run: python main.py
- On the first run, scan the WhatsApp Web QR code using your phone.
- The script will process queued contacts, attach the poster, insert the personalized caption, send, and pause between messages.


4. CONFIGURATION (config.json)
-----------------------------------------------------------------------------
You can open config.json in Notepad to adjust:
- "auto_send": true           -> Automatically clicks send
- "min_delay_seconds": 25     -> Minimum pause between contacts (seconds)
- "max_delay_seconds": 45     -> Maximum pause between contacts (seconds)
- "batch_size": 15            -> How many messages to send before a cooling rest
- "batch_pause_minutes": 5    -> Minutes to pause after each batch
- "attach_post": true         -> Set to false if you only want to send text
- "require_opt_in": true      -> Only sends to contacts where Opt_In is YES


5. WHATSAPP ACCOUNT SAFETY & POLICY NOTICE
-----------------------------------------------------------------------------
- WhatsApp strictly prohibits unauthorized bulk messaging and spam bots.
- Client-side tricks (such as spinning words or random delays) DO NOT guarantee
  an account will not be flagged by WhatsApp's machine learning systems.
- WhatsApp detects bulk automation primarily through:
  * Recipient reporting and blocking (if users hit "Report Spam", bans occur quickly)
  * Sudden spikes in messages sent to people who do not have your number saved
  * Outbound-only traffic without normal two-way conversation history
- BEST PRACTICES:
  1. Only message contacts who explicitly opted in and know your academy.
  2. Keep volumes small and spread across the day.
  3. Never spam cold/unverified numbers.
  4. For official, high-volume automated marketing campaigns, use the official
     WhatsApp Business Platform / Cloud API via Meta for Developers.
