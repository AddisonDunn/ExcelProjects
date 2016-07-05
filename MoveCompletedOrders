'archives "picked up" items
Sub MoveCompletedOrders()
    'keeps sheet from running methods that run whenever the a sheet is changed. This way we can edit
    'the sheet without disturbance from other methods trying to run at the same time.
    On Error GoTo ErrorHandler
    Application.EnableEvents = False
    
    Set FromWS = Worksheets("Orders In Progress") 'where the majority of work is done
    FromWS.Activate
    Set ToWS = Worksheets("Order History")
    
    'Unprotect the worksheets we want to edit to avoid errors
    FromWS.Unprotect Password:="ir"
    ToWS.Unprotect Password:="ir"

    If Not Application.Intersect(ActiveCell, Range("A2:J1000")) Is Nothing Then
        Range("C1").Activate
    End If
    
    'save location of the user-selected cell so it can be re-selected after method is finished
    Set OIPoriginalCell = ActiveCell
    
    'Changes the location of the ActiveCell so that actions can be performed at that range
    Range("B2").Activate
    
    Dim status As String
    Dim movingCell As Range

    'loops through statuses
    i = 2
    
    Do Until i > Range("B" & Rows.Count).End(xlUp).Row
        
        status = LCase(Cells(i, 2).Value)
        
        'move an order if its status is "picked up"
        If (status = "picked up") Then
            
            'the order's location in "orders in progress"
            Set sourceRange = Range(Cells(i, 1), Cells(i, 1).Offset(0, 8))
            
            'tempRange is set to be the first empty cell in column A
            Set tempRange = ToWS.Range("A" & (ToWS.Range("A" & ToWS.Rows.Count).End(xlUp).Row + 1))
            'targetRange is location in Order History where the data will be copied. It is the row that tempRange is located at
            Set targetRange = Range(tempRange, tempRange.Offset(0, 8))
            
            'copy sourceRange data to targetRange
            targetRange.Value = sourceRange.Value
            
            'this WITH is to calculate the money saved on an order
            With ToWS
                Dim objCell As Range
                Dim intRow As Integer
                
                'add in Amazon prices for money-saved calculation
                    
                ' Get the actual row index (in the worksheet).
                intActualRow = targetRange.Row
                
                'if the item is already from amazon then there is no money saved compared to the amazon price
                If .Range("E" & intActualRow).Value = "Amazon" Then
                    .Range("J" & intActualRow).Value = 0
                    
                ElseIf Not IsEmpty(.Range("E" & intActualRow).Value) Then 'item is from Microcenter or Newegg
                    Set DataWS = Worksheets("Products")
                    Dim amazonPrice, lowestPrice, difference, quantity As Double
                    
                    'temp is used as a temporary variable to be the middle man between a range value and a double-
                    'without it, there are many issues with casting the values returned by the index/range.value functions
                    'Finds the Amazon price on the products sheet using concatenations
                     Set temp = Application.Index(DataWS.Range("$C$2:$G$1000"), Application.Match((.Range("C" & intActualRow).Value) & "Amazon", DataWS.Range("$A$2:$A$1000"), 0), 3)
                     
                     
                     amazonPrice = temp.Cells(1, 1).Value
                     lowestPrice = .Range("F" & intActualRow).Value 'price of the current item
                     
                     difference = amazonPrice - lowestPrice
                     quantity = .Range("D" & intActualRow).Value
                     'number format has to be set or excel will round it to nearest dollar
                     .Range("J" & intActualRow).NumberFormat = "0.00"
                     'set the J column value to be the money saved
                     .Range("J" & intActualRow).Value = difference * quantity
                End If
            End With
            
            'clear the data off of the orders in progress sheet
            sourceRange.Delete Shift:=xlUp
            Cells(i, 1).Offset(0, 9).Delete Shift:=xlUp
            
        Else
            i = i + 1
        End If
    
    Loop
    
    'sort the order history after we edited it
    ToWS.Activate
    SortOrders
    FromWS.Activate
    SortOrders
    
ErrorHandler:
    Application.EnableEvents = True
    
    'set active cell back to the orginial
    'OIPoriginalCell.Activate

    ToWS.Protect Password:="ir", UserInterFaceOnly:=True
    FromWS.Protect Password:="ir", UserInterFaceOnly:=True
End Sub

'sorts Orders In Progress and Order History
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
