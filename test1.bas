Attribute VB_Name = "Module1"
Sub randomWeight()
Dim lastRow
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastRow
    Cells(i, 16) = Int((500 - 100 + 1) * Rnd + 100)
Next i
End Sub


Sub name_reduce_columns()
    'data manipulation => delete columns
    
    'parser would take the column name inputs and return the index
    
    'let user choose the column to have their rows counted on, default to 1
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    colIndex_of_weight = aux_find_col_index("weight")
    colIndex_of_DispositionIDDesc = aux_find_col_index("DispositionIDDesc")
    
    Range("P:P,K:K").Select
    
    
End Sub


Function aux_find_col_index(colName)
    'must have addon
    lastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        If Cells(1, i) = colName Then
            aux_find_col_index = i
            Exit For
        End If
    Next i
End Function
