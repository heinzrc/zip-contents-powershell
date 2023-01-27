# zip-contents-powershell
Parses nested zip and 7z files for their contents. 

Outputs results to a CSV on the desktop

This script has been setup to run via drag and drop for accessability.

# Drag and Drop Setup:

First, create a new shortcut on the desktop:

![image](https://user-images.githubusercontent.com/65210276/215060697-49bc4467-872f-403e-84a0-6b00a69be6eb.png)

In the resulting window, enter "C:\Windows\explorer.exe" and name it whatever you'd like:

![image](https://user-images.githubusercontent.com/65210276/215061224-acb256c0-6053-4f52-ba39-fd6976354e44.png)

Next, right click the shortcut and go to properties:

![image](https://user-images.githubusercontent.com/65210276/215061531-efff4bf9-4451-41eb-ba85-c6741f206620.png)

In the properties window, change the target to the following:
(Make sure to replace {Path-to-script} with the path to where you downloaded the ps1 file)

C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit -noprofile -file "{Path-to-script}\Get-DirectoryContents.ps1"

![image](https://user-images.githubusercontent.com/65210276/215062244-bc6e525b-502e-415c-91bf-0549f21019a1.png)

After clicking apply and ok your shortcut should change to resemble the powershell icon:

![image](https://user-images.githubusercontent.com/65210276/215062798-a0d05e76-b6d9-4af5-b3b7-afe7216deea6.png)

Now, dragging any folder, zip, or 7z should open a csv with all the contents inside them.
