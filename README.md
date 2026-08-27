# Expense Claim Form Generator

Scan receipts and turn them into expense claim forms. Powered by LlamaParse.

## Installation Instructions

### For regular users: Download release

> [!NOTE]
> Please find Merton for tech support if anything doesn't work, or if I am not here, the IT department can probably help you

1. Download the corresponding ZIP file from the releases page. Depending on your operating system, you will need to download a different file.

2. Extract the file. A new folder named `ExpenseClaimForm-[OS_TYPE]` should appear.

3. Open that folder and double-click the executable named `ExpenseClaimForm`. Depending on your operating system, it may have a different file extension.

> [!IMPORTANT]
> Windows Defender or another antivirus program may mark this application as suspicious or malicious. This is because the application is packaged with PyInstaller, which is also commonly used to package malware. Please be assured that we are not trying to hack you. If necessary, configure your antivirus program to allow the application.

### For advanced users: Run with Python

1. Clone the repository:
   ```
   git clone https://github.com/adrienpartier-cmd/VSceptre-Expense-Claim-Form.git
   cd VSceptre-Expense-Claim-Form
   ```

2. Create a virtual environment. This step is technically optional, but it helps avoid conflicts with packages installed for other purposes, so it is strongly recommended.

   Windows:
   ```
   python3 -m venv venv
   .\venv\Scripts\activate
   ```

   macOS/Linux:
   ```
   python3 -m venv venv
   source venv/bin/activate
   ```

> [!NOTE]
> 
> We tested on Python 3.14 and 3.13, older versions may or may not work.

3. Install the required packages:
   ```
   pip install -r requirements.txt
   ```
4. Run the program:
   ```
   python llamacode.py
   ```
5. Optional: Package for distribution with PyInstaller:
   ```
   pip install pyinstaller
   pyinstaller -F -n ExpenseClaimForm llamacode.py
   ```

## Legal notices

All original code in this application is licensed under the MIT License. Some third-party libraries may be licensed differently.

See [LEGAL_NOTICES.md](LEGAL_NOTICES.md) for included library licenses and source links.
