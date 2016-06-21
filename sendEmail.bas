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
    
    'whatever text you want to be put into your txt file
    Set desiredText = "i love my new job."
    
    'write the desired text to the file
    writeFile.Write (desiredText)
    
    'clean up
    writeFile.Close
    Set readFile = Nothing
    Set writeFile = Nothing
    Set FSO = Nothing

    'make sure to use the correct file location. It will be in the same location as your Excel file
    strCommand = "powershell.exe -ExecutionPolicy Unrestricted -File ""C:\FILE_LOCATION"""
    'sends the email through the shell
    Shell (strCommand)
End Sub
