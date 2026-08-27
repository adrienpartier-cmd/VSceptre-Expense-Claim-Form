# Expense Claim Form Generator

Scan receipts and turn them into expense claim forms

## Install instructions

### For regular users: Download release

download the corresponding zip file from the releases page

foo bar

not completed

> [!IMPORTANT]
> Windows Defender or whatever antivirus you use will probably mark this as suspicious/malware. This is because the way we used to package the application (pyinstaller) is also commonly used to package malicious programs. Please be assured that we are not interested in hacking anyone. Just tell Windows Defender to ignore it.

### For advanced users: Run with Python

1. Git clone our repo
   ```
   git clone https://github.com/adrienpartier-cmd/VSceptre-Expense-Claim-Form.git
   cd VSceptre-Expense-Claim-Form
   ```

2. Make a virtual environment. This step is technically optional, but it avoids causing conflicts with other packages that you may install/have installed for other purposes, so it is strongly recommended.

   Windows:
   ```
   python3 -m venv venv
   .\venv\Scripts\activate
   ```

   MacOS/Linux:
   ```
   python3 -m venv venv
   source venv/bin.activate
   ```
   
> [!NOTE]
> 
> We tested on Python 3.14 and 3.13, older versions may or may not work.

3. Install the required packages
   ```
   pip install -r requirements.txt
   ```
4. Run program
   ```
   python3 llamacode.py
   ```
5. Optional: package for distribution using pyinstaller
   ```
   pip install pyinstaller
   pyinstaller -Fn ExpenseClaimForm llamacode.py
   ```

**the above is not yet finalized please fix before present