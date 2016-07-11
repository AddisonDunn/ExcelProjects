Sub AddItemsToCart()
    
    Const ForReading = 1    '
    'Const fileToRead As String = "C:\Users\david_zemens\Desktop\test.txt"  ' the path of the file to read
    Const fileToWrite As String = "orderingLinks.txt"  ' the path of a new file
    Dim FSO As Object
    Dim writeFile As Object 'the file you will CREATE
    Dim repLine As Variant   'the array of lines you will WRITE
    Dim ln As Variant
    Dim l As Long
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set writeFile = FSO.CreateTextFile(fileToWrite, True, False)

    Range("C2").Activate
    Dim firstLine As Integer
    firstLine = 1

    Do Until (IsEmpty(ActiveCell.Value))
        'check to see if the corresponding link-cell for an item is empty
        If Not IsEmpty(ActiveCell.Offset(0, 5).Value) And Not ActiveCell.Offset(0, 5).Value = 0 Then
            'start the next line
            If firstLine = 0 Then
                writeFile.Write (Chr(13) & Chr(10))
            End If
            'quantity
            writeFile.Write (ActiveCell.Offset(0, 1).Value)
            'if not, then write the link to the text file
            writeFile.Write (ActiveCell.Offset(0, 5).Value)
            
            firstLine = 0
        
        Else
            MsgBox ActiveCell.Value & " does not have a link to order"
        End If
        
        ActiveCell.Offset(1, 0).Select
    Loop
 
    writeFile.Close
    
    '# clean up
    Set readFile = Nothing
    Set writeFile = Nothing
    Set FSO = Nothing
    
    'location of the AutoIT file
    Const fileLocation As String = "C:\Users\addison.dunn\Documents\AutoIT_Purchasing.exe"
    'Shell (fileLocation)
    
    
End Sub
