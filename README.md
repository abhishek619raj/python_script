What Is This?
This is a simple Python/automation application.

Install Python for prefered Windows/ Linux / Mac
https://www.python.org/ftp/python/3.11.1/Python-3.11.1-amd64.exe 
# Check on ADD to path if asked by Installation Manual

Take the project from git (*url to be provided) or else download from this link
# https://agreeyasolutions-my.sharepoint.com/personal/abhishek_raj_agreeya_net/_layouts/15/onedrive.aspx?login_hint=abhishek%2Eraj%40agreeya%2Enet&id=%2Fpersonal%2Fabhishek%5Fraj%5Fagreeya%5Fnet%2FDocuments%2FAutomation%2FPython%5FScript

Once the project Folder Downloaded open the folder into CMD (command line)
# for example "C:\Users\xyz\Download\Python_Script>"

RUN "pip install virtualenv"

RUN "virtualenv venv"

RUN "cd venv\Scripts\"

RUN "activate"

RUN "cd .."

RUN "cd .."

RUN "pip install -r requirment.txt"  
# Make sure the location looks like C:\Users\xyz\Download\Python_Script>

RUN "python GenericPythonCode.py"

After Code RUN Sucessfully please check the File inside all_policy Folder and dowloads folder.