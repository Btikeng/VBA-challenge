Sub analyzeData()
    Dim Sh As Worksheet
    For Each Sh In Worksheets
        n = Sh.Name
        lr = Sh.Cells(Rows.Count, 9).End(xlUp).Row
        Sh.Range("I2:L" & lr).ClearContents
        Sh.Range("P2:Q4").ClearContents
        nodupeLIST Sh:=Sh
        Dim fnds As Range, fnde As Range, Rng As Range
        With Sh
            lr = .Cells(Rows.Count, 9).End(xlUp).Row
            For i = 2 To lr
                Set fnds = .Range("A:A").Find(.Cells(i, 9).Value)
                Set fnde = .Range("A:A").Find(.Cells(i, 9).Value, , , xlWhole, xlByRows, xlPrevious, False, , False)
                .Cells(i, 10).Value = fnde.Offset(0, 5).Value - fnds.Offset(0, 2).Value
                 If fnds.Offset(0, 2).Value <> 0 Then
                    .Cells(i, 11).Value = (fnde.Offset(0, 5).Value - fnds.Offset(0, 2).Value) / fnds.Offset(0, 2).Value
                Else
                    .Cells(i, 11).Value = 0
                End If
                .Cells(i, 12).Value = WorksheetFunction.Sum(.Range(fnds.Offset(0, 6), fnde.Offset(0, 6)))
            Next i
            minVal = Application.WorksheetFunction.Min(.Range("K2:K" & lr))
            maxVal = Application.WorksheetFunction.Max(.Range("K2:K" & lr))
            maxVol = Application.WorksheetFunction.Max(.Range("L2:L" & lr))
            For Each Rng In .Range("K2:K" & lr)
                If Not IsError(Application.Match(minVal, Rng, 0)) Then
                    .Range("Q3").Value = Rng.Value
                    .Range("P3").Value = Rng.Offset(0, -2).Value
                End If
                If Not IsError(Application.Match(maxVal, Rng, 0)) Then
                    .Range("Q2").Value = Rng.Value
                    .Range("P2").Value = Rng.Offset(0, -2).Value
                End If
            Next Rng
            For Each Rng In .Range("L2:L" & lr)
                If Not IsError(Application.Match(maxVol, Rng, 0)) Then
                    .Range("Q4").Value = Rng.Value
                    .Range("P4").Value = Rng.Offset(0, -3).Value
                End If
            Next Rng
        End With
    Next Sh
End Sub
Sub nodupeLIST(Sh As Worksheet)
    Dim r1 As Range, lastrow As Long

    With Sh
        lastrow = .Cells(Rows.Count, "A").End(xlUp).Row
        Set r1 = .Range("A2:A" & lastrow)
    End With

    With Sh
        With .Range("I2").Resize(r1.Rows.Count, 1)
            .Cells = r1.Value
            .RemoveDuplicates Columns:=1, Header:=xlNo
        End With
    End With

End Sub

