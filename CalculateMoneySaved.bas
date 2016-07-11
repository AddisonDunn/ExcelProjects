Sub CalculateMoneySaved()
    Set FromWS = Worksheets("Order History")
    FromWS.Unprotect Password:="ir"

    Dim month, year, allTime, itemValue As Double
    month = 0#
    year = 0#
    allTime = 0#
    
    Dim objCell As Range
    Dim intRow As Integer
    
    'the location of the money saved column
    Set targetRange = Range("J2:J" & (Range("J" & Rows.Count).End(xlUp).Row + 1))

    For intRow = 1 To targetRange.Rows.Count Step 1
        Set objCell = targetRange.Rows(intRow)
        
        ' Get the actual row index (in the worksheet).
        intActualRow = objCell.Row
        
        itemValue = Range("J" & intActualRow).Value
        
        'add the money saved by the item to the running total
        allTime = allTime + itemValue
        
        'if item is younger than a year
        If DateDiff("d", Range("A" & intActualRow), Date) < 366 Then
            year = year + itemValue
            
            'if item is younger than 31 days
            If DateDiff("d", Range("A" & intActualRow), Date) < 31 Then
                month = month + itemValue
            End If
        End If
        
        
    
    Next
    
    Range("M2").Value = month
    Range("M3").Value = year
    Range("M4").Value = allTime
    
    FromWS.Protect Password:="ir", UserInterFaceOnly:=True
    
End Sub
