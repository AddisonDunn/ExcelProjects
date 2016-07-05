Sub UpdateProductsList()
    Set originalCell = ActiveCell

    Dim sourceRange As Range
    Dim targetRange As Range
    
    Set FromWS = Worksheets("Products") ' this is where the majority of work is done
    Set ToWS = Worksheets("Order Form")
    
    'finds the last empty cell in column C
    Set sourceRange = Range("C" & (Range("C" & Rows.Count).End(xlUp).Row + 1))
        
    With ToWS
    
        'targetRange will be the data in column C of the Order Form
        Set targetRange = .Range("$C$2:C" & (.Range("C" & .Rows.Count).End(xlUp).Row + 1))
        'data validation dropdown is added to the Name and Website columns on the Order Form
        With targetRange.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=("=Products!" & "$C$2:" & sourceRange.Address)
        End With
        
        Set targetRange2 = .Range("$E$2:E" & (.Range("E" & .Rows.Count).End(xlUp).Row + 1))
        With targetRange2.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=("=Products!N8:N10")
        End With
        
    End With
    
    'bottom cell of the list of categories
    Set sourceRange2 = Range("N" & (Range("N" & Rows.Count).End(xlUp).Row + 1))
    'where we want to put data validation for self-updating list of categories
    Set targetRange3 = Range("$H$2:H" & (Range("C" & Rows.Count).End(xlUp).Row))
    With targetRange3.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=("=$N$18:" & sourceRange2.Address)
    End With
    
    
    'Dates
    Range("C2").Select

    Do Until IsEmpty(ActiveCell)
        
        'current date is added if there is none
        If (IsEmpty(ActiveCell.Offset(0, -1))) Then
            ActiveCell.Offset(0, -1).Value = Date
        'if date is >30 days old, add an orange highlight
        ElseIf DateDiff("d", ActiveCell.Offset(0, -1).Value, Date) > 30 Then
            ActiveCell.Offset(0, -1).Interior.ColorIndex = 46
        Else
            ActiveCell.Offset(0, -1).Interior.ColorIndex = 0
        End If
        

        ActiveCell.Offset(1, 0).Select
    Loop
    
    'Sort products list by category (alphabetical)
    Set rangeToBeSorted = Range("B2", "H1000")
    rangeToBeSorted.Sort Key1:=Range("H2"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    
    
    originalCell.Activate
End Sub
