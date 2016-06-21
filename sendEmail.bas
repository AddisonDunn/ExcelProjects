Attribute VB_Name = "Module2"
'Method that sends desired text to an email address through Powershell
Sub sendEmail()
    
    Const fileToWrite As String = "informationToBeMailed.txt"  ' the path of a new file
    Dim FSO As Object
    Dim writeFile As Object 'the file you will CREATE
    Dim repLine As Variant   'the array of lines you will WRITE
    Dim ln As Variant
    Dim l As Long
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set writeFile = FSO.CreateTextFile(fileToWrite, True, False)
    
    'Whatever text you want to be put into your txt file
    Set desiredText = "i love my new job."
    
    'Write the desired text to the file
    writeFile.Write (desiredText)
    
    'clean up
    writeFile.Close
    Set readFile = Nothing
    Set writeFile = Nothing
    Set FSO = Nothing

    'Make sure to use the file location of your Powershell script that will send an email using the same txt filename.
    'Such a .ps1 file can be found at github.com/AddisonDunn/PowershellScripts
    strCommand = "powershell.exe -ExecutionPolicy Unrestricted -File ""C:\FILE_LOCATION"""
    'Sends the email through the shell
    Shell (strCommand)
End Sub
