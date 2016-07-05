Sub SortOrderForm()
    Set originalCell = ActiveCell
    
    Set rangeToBeSorted = Range("C2", "J1000")
    rangeToBeSorted.Sort Key1:=Range("C2"), Order1:=xlDescending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    
    originalCell.Activate
End Sub
