randomPrint
===========

This program accepts a set of documents with the number of copies desired, and prints them in random order.

Usage:
Download the randomPrint.ps1 file to a folder on your harddrive. You then need to start powershell (should be installed by default on win7 or newer (search for powershell using start-menu, and click it). 

Navigate to the folder where you saved the script, and run it by typing "`.\randomPrint.ps1`". 

If you get an execution policy error you need to allow execution of local unsigned scripts on your pc by running powershell in admin mode (left click the icon on your taskbar and select "Run as administrator") and then type "`Set-ExecutionPolicy RemoteSigned`" followed by enter. You should then be able to run the script as described above.

The script will prompt for input parameters: supply the full path of the documents followed by the number of copies desired. Printing will start after you answer no to a "More documents?" question.
