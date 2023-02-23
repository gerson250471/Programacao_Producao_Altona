Private Sub RelataAtrasoSeq()
    L = 5
    While P10.Range("A" & L) <> ""
        If P10.Range("M" & L) >= P10.Range("E" & L) Then
            P10.Range("A" & L & ":M" & L).Font.Color = vbRed
            P10.Range("A" & L & ":M" & L).Font.Bold = True
            P10.Range("A" & L & ":M" & L).Font.Size = 12
        Else
            P10.Range("A" & L & ":M" & L).Font.Color = vbBlack
            P10.Range("A" & L & ":M" & L).Font.Bold = False
            P10.Range("A" & L & ":M" & L).Font.Size = 12
        End If
        L = L + 1
    Wend
End Sub