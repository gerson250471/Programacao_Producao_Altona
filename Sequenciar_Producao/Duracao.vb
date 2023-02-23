Private Sub Duracao()
    Tp = Hour(P10.Range("k" & L) + P10.Range("l" & L))
    If Tp < 13 Then
        Refeicao = 0.0208
    ElseIf Tp < 21 Then
        Refeicao = 0.0416
    Else
        Tp = P10.Range("k" & L) + P10.Range("l" & L)
        Tp = WorksheetFunction.RoundUp(P10.Range("k" & L) + P10.Range("l" & L), 0) - Tp
        Refeicao = Tp
    End If
End Sub
