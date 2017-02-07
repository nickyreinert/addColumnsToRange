
Sub addColumnsToNamedRange()

    lastColumn = Range("A1").SpecialCells(xlCellTypeLastCell).Column


    For currentColumn = 1 To lastColumn
        
        Cells(1, currentColumn).Activate
        
        If Not IsEmpty(ActiveCell.Value) Then
        
            targetRange = ActiveCell.Value
            
            ActiveCell.EntireColumn.Select
            
            On Error Resume Next

            ActiveWorkbook.Names.Add Name:=targetRange, RefersTo:=Range(targetRange & "," & Selection.Address)

            If Err <> 0 Then
            
                Debug.Print "Identified range does not exists: " & targetRange
                
            Else
            
                Debug.Print "Identified range found, extended it with " & Selection.Address
                
            End If

        End If

    Next currentColumn
    




End Sub
