# MeroShare
MeroShare Automation

This is a simple Python script to automate certain tasks in MeroShare (meroshare.cdsc.com.np)


Installation:
------------
1. Download Python and install with 'admin privileges' and 'add to path' checked
2. Install Requirements



Files List:
-----------

1. Menu.py
    - Will be your default interface to access the program

2. MeroShare Login Details.xlsx
    - Contains Login Details for clients
    - Fields for 'Name'(optional), 'Client ID', 'DP ID', and 'Password' are required to login into the website. Therefore should not be blank.
    - 'CRN', 'Pin', and 'Bank Name' are required to apply for IPOs
    - The 'Active' field indicates if the client is currently ACTIVE. Fill in 'NO' if you don't want the program to phase through the entry. (Used in 'Check Acc Status', 'Get Share List', 'Get Applicable Issue')
    - The 'Apply IPO' field indicates if you want to apply for an IPO for the client. Fill in 'NO' if you don't want to apply for IPO for the entry. (Used in 'Get Applicable Issue', 'Apply for IPO')
    - It is recommended to fill in 'NO' for the 'Apply IPO' field if 'Active' is filled in as 'NO'
    - Both fields should be blank if the entry is to be passed to the program
    - 'Bank Name' should be the exact copy of what appears in the 'Bank' dropdown (not 'Branch' but 'Bank') when applying for IPO down to CAPITALIZATION and the dot(.) that may or may not be after the LTD in banks name (Eg, GLOBAL IME BANK LTD.). Be sure there are no extra spaces before or after the full name if there are none in the website dropdown. If the field is any different the website will reject the application request.

3. The 'files' folder contains the core program modules and other req files.
    - All files with .py extensions are the Python script (program) that are accessed through 'Menu.py'
    - 'capital.json' contains the list of capitals from the login dropdown (if deleted the program will recreate the file during the next run). If the website changes the order for the list, maybe the file will need to be deleted so that it can be updated. I have not had any problems till now though.
    - 'cdsc-com-np-chain.pem' contains the SSL certificate to verify the request sent to the server. It's a recent change that I had to add and maybe it should be updated after a few months. The process to update will be somewhere below.



What the program can do:
------------------------
All the options will output an Excel file after it is done running.

1. Check Acc Status
    - Checks if the entry can log in
    - Use to check if other options are available for the entry
    - No other options will be able to be used if you can't log in
    - Status and their meaning are pretty self-explanatory
    - 'Login Failed! 401' means that the password is wrong

2. Get My Share List
    - Will output an Excel file with the items from the "My Share" tab from the website (Current and Free Balance)
    - 'Login Failed! 401' means that the password is wrong
    - 'Share list Connection Refused!' is when the program can't access the tab due to some reason.

3. Check Applicable Issue
    - Will return the 'Apply for Issue' tab from 'My ASBA'
    - Useful to check if the clients are eligible to apply for an active Issue (IPO, Right or Foreign Employment). Mainly use it to see if someone is eligible for Rights.

4. Apply for Issue
    - Use to apply IPO and Debentures
    - Does not work for Rights
    - 'Apply Failed' may be due to different reasons ranging from network error to password error
    - 'Couldn't Apply - edit' means it has already been applied and is available to edit from the website
    - 'Couldn't Apply - in-process' means that it has applied already and is not available for editing
    - 'Couldn't Apply - reapply' means for some reason it needs to be reapplied. Check through the website and Reapply. This Program cant Reapply
    - There are other outputs too but these are the common one

5. Check Application Status
    - Use to check IPO or other application status
    - Application Status is from the MeroShare website and not the IPO results one.
    - It usually takes a while for the IPO results to be updated on the MeroShare website.




Updating 'cdsc-com-np-chain.pem':
--------------------------------
1. Download Firefox
2. Open MeroShare website
3. Click on the lock icon in the address bar
4. Connection Secure -> More Info
5. A popup tab will open. Click on 'View certificate'
6. Scroll a little bit and download the 'PEM (chain)'
7. Replace the old file with the newly downloaded file
