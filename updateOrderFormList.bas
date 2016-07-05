Sub updateOrderFormList()

    Set originalCell = ActiveCell
    Set FromWS = Worksheets("Products")
    Set ToWS = Worksheets("Order Form")

    Dim arr(100) As String 'will be array of products list data
    arr(1) = "a"
    Dim priceMatchTicker(100) As String
    
    Range("C2").Activate
    
    Do Until IsEmpty(ActiveCell)
        
        If IsInArray(ActiveCell.Value, arr) Then 'is the current item we're on a duplicate?
            Dim i As Integer
            For i = 2 To UBound(arr) 'look through array to find the other item with the same name
                If (arr(i) = ActiveCell.Value) And (Not Range("D" & i).Value = ActiveCell.Offset(0, 1).Value) Then
                ' if statements to find the item with the lower value
                    If (ActiveCell.Offset(0, 2).Value < (Range("E" & i).Value - 1#)) Then
                        arr(ActiveCell.Row) = ActiveCell.Value
                        arr(i) = ""
                        If Not LCase(ActiveCell.Offset(0, 1).Value) = "microcenter" Then
                            priceMatchTicker(ActiveCell.Row) = "1"
                        End If
                    'tie goes to microcenter (within a dollar)
                    ElseIf (ActiveCell.Offset(0, 2).Value < (Range("E" & i).Value + 1#)) And (LCase(ActiveCell.Offset(0, 1).Value) = "microcenter") Then
                        arr(ActiveCell.Row) = ActiveCell.Value
                        arr(i) = ""
                    ElseIf Not LCase(Range("E" & i).Value) = "microcenter" Then
                        priceMatchTicker(i) = "1"
                    End If
                End If
            Next i
        Else
            'if there is no duplicate, add the item to the array!
            arr(ActiveCell.Row) = ActiveCell.Value
        End If
        
        ActiveCell.Offset(1, 0).Select
    Loop
    
    'moves us over to the Order Form
    'ToWS.Activate
    ToWS.Range("M2:R1000").ClearContents
    
    
    With ToWS
        Dim x As Integer
        Dim y As Integer
        
        'iterates through the array, adding the data from the Products list to the Order Form
        y = 2
        For x = 2 To UBound(arr)
            If Not (arr(x) = "") Then
                Set movingCell = .Cells(y, 14)
            
                Dim targetRange, sourceRange As Variant
                'copy data from a row in Products list and adds it to the list on the Order Form
                Set targetRange = .Range(movingCell.Offset(0, -1), movingCell.Offset(0, 5))
                Set sourceRange = FromWS.Range("B" & x & ":H" & x)
                
                sourceRange.Copy (targetRange)
                
                'if the item can be price matched, mark it
                If priceMatchTicker(x) = "1" Then
                    movingCell.Offset(0, 6).Value = "Y"
                Else
                    movingCell.Offset(0, 6).Value = "N"
                End If
                
                y = y + 1
            End If
        Next x
    End With
    
    'FromWS.Activate
    originalCell.Activate
    
End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean

    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)

End Function
