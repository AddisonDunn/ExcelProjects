Sub RemoveItem2()
    'keeps sheet from running methods that run whenever the a sheet is changed. This way we can edit
    'the sheet without disturbance from other methods trying to run at the same time.
    On Error GoTo ErrorHandler
     
    Application.EnableEvents = False
    
    Set FromWS = Worksheets("Orders In Progress")
    FromWS.Unprotect Password:="ir"
    
    Set originalCell = ActiveCell
    
    'If the first cell in the selection is over the products list
    If (ActiveCell.Column < 11) Then
        Dim objSelection As Range
        Dim objSelectionArea As Range
        Dim objCell As Range
        Dim intRow As Integer
        
        Set objSelection = Application.Selection
        ' Get the current selection
        Set objSelection = Application.Selection
    
        ' Walk through the areas
        For Each objSelectionArea In objSelection.Areas
    
            ' Walk through the rows
            For intRow = 1 To objSelectionArea.Rows.Count Step 1
                
                ' Get the row reference
                Set objCell = objSelectionArea.Rows(intRow)
                
                ' Get the actual row index (in the worksheet).
                intActualRow = objCell.Row

                Set sourceRange = Range(Range("A" & intActualRow), Range("J" & intActualRow))
                sourceRange.ClearContents
                               
            Next
        Next
    End If
        
ErrorHandler:
    Application.EnableEvents = True
    
    SortOrders
    
    originalCell.Activate
    FromWS.Protect Password:="ir", UserInterFaceOnly:=True
End Sub
