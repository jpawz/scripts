Sub grupuj_bom()
'
' grupuj_bom Makro
' Grupuje BOM wyeksportowany z Windchilla (Multi-Level Bill of Materials Report)
' do excela (xlsx). Grupuje według kolumny "A", gdzie pierwszy wiersz to nagłówek,
' a kolejne wiersze tej kolumny to oznaczenia poziomów elementu w drzewie zespołu
'
'
Dim level, rowNumber, maxLevel, nRows, startRow, endRow As Integer
Dim levelFound As Boolean

nRows = Worksheets(ActiveSheet.Name).Range("A2").End(xlDown).Row

Range("A1:A" & nRows) = Range("A1:A" & nRows).Value
Range("A1:A" & nRows).NumberFormat = "0"

maxLevel = WorksheetFunction.Max(Range("A1:A" & nRows))

If maxLevel > 8 Then maxLevel = 8


For level = maxLevel To 1 Step -1
    For rowNumber = 2 To nRows - 1
        levelFound = False
        If Range("A" & rowNumber).Value >= level Then
            startRow = rowNumber
            endRow = rowNumber
            levelFound = True
        End If
        While Range("A" & rowNumber).Value >= level And Range("A" & rowNumber + 1).Value >= level
            endRow = rowNumber + 1
            rowNumber = rowNumber + 1
        Wend
        If levelFound = True Then
            Rows(startRow & ":" & endRow).Group
            levelFound = False
        End If
    Next rowNumber
Next level


'
End Sub
