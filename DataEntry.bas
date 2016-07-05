Sub SortOrders()
    'just use the active sheet reference this time--if we call this method, we will already be on the sheet we want to be on
    ActiveSheet.Unprotect Password:="ir"

    Set originalCell = ActiveCell
    
    'sorts items by the first column (in this case, the date)
    Set rangeToBeSorted = Range("A2", "J1000")
    rangeToBeSorted.Sort Key1:=Range("A2"), Order1:=xlDescending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    
    'sorts items by second column, in this case, the name
    Set rangeToBeSorted = Range("A2", "J1000")
    rangeToBeSorted.Sort Key1:=Range("B2"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    
    originalCell.Activate
    ActiveSheet.Protect Password:="ir", UserInterFaceOnly:=True
End Sub

'moves items from Order Form to Orders In Progress
Sub DataEntry()
    'keeps sheet from running methods that run whenever the a sheet is changed. This way we can edit
    'the sheet without disturbance from other methods trying to run at the same time.
    On Error GoTo ErrorHandler
     
    Application.EnableEvents = False
    
    Set ToWS = Worksheets("Orders In Progress")
    ToWS.Unprotect Password:="ir"

    'Move the active cell if it is in a selection that is about to be edited (so that no errors occur)
    If Not Application.Intersect(ActiveCell, Range("A2:J1000")) Is Nothing Then
        Range("K2").Activate
    End If
    Set originalCell = ActiveCell
    
    AddItemsToCart
    
    Sheets("Order Form").Activate
    Range("C2").Activate
    
    Set FromWS = Worksheets("Order Form")
    
    
    'loops through all entries
    If (Not IsEmpty(ActiveCell.Value)) Then
        
        
        'finds the bottom right cell of the range we want to copy
        Set tempRange1 = Range("J" & (Range("C" & Rows.Count).End(xlUp).Row))
        'sourceRange will be the range we copy
        Set sourceRange = Range(Range("C2"), tempRange1)
        
        'check to see if item already exists in Orders in Progress
        With ToWS
            Dim name As String
            Dim numRows As Integer
            numRows = sourceRange.Rows.Count
            'iterate through orders in progress to see if item already exists
            For i = 1 To numRows Step 1
                name = sourceRange.Cells(i, 1).Value
                For j = 2 To .Range("C" & Rows.Count).End(xlUp).Row + 1 Step 1
                    If (.Cells(j, 3).Value = name) Then
                        .Cells(j, 4).Value = .Cells(j, 4).Value + sourceRange.Cells(i, 2).Value
                        .Cells(j, 7).Value = .Cells(j, 7).Value + sourceRange.Cells(i, 5).Value
                        sourceRange.Rows(i).Delete Shift:=xlUp
                        numRows = numRows - 1
                    End If
                Next
            Next
        End With
        
        'check if sourceRange exists
        If numRows > 0 Then
            'copy entries
            sourceRange.Copy
            
            'move all items over
            With ToWS
                Dim targetRange As Range
                
                'finds the top left cell of the range we want to paste to
                Set targetRange = .Range("C" & (.Range("C" & Rows.Count).End(xlUp).Row + 1))
                
                'use special paste to only paste the values(so formulas do not get carried over
                targetRange.PasteSpecial xlPasteValues
                'set the date to today's date and the status to the default value
                Range(targetRange.Offset(0, -2), targetRange.Offset(sourceRange.Rows.Count - 1, -2)).Value = Date
                Range(targetRange.Offset(0, -1), targetRange.Offset(sourceRange.Rows.Count - 1, -1)).Value = "Requested"
                
                
            End With
            
            'clear contents instead of delete so that references to the cells are not confused
            sourceRange.ClearContents
        End If
        
        
        
    End If
    
    'sort orders in progress after data is moved there
    ToWS.Activate
    SortOrders
    FromWS.Activate
    
    Range("D2").Activate
    'set formulas for the index/match lookup
    ActiveCell.Value = "0"
    ActiveCell.Offset(0, 1).Interior.ColorIndex = 0
    ActiveCell.Offset(0, 2).Value = "=INDEX(Products!C2:G5000,MATCH(C2&E2, Products!A2:A5000, 0), 3)"
    ActiveCell.Offset(0, 3).Value = "=D2*F2"
    ActiveCell.Offset(0, 4).Value = "=INDEX(Products!C2:G5000,MATCH(C2&E2, Products!A2:A5000, 0), 4)"
    ActiveCell.Offset(0, 5).Value = "=INDEX(Products!C2:G5000,MATCH(C2&E2, Products!A2:A5000, 0), 5)"
    
    
ErrorHandler:
    Application.EnableEvents = True
        
    originalCell.Activate
    
    ToWS.Protect Password:="ir", UserInterFaceOnly:=True
End Sub
