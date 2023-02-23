Private Sub PrimeiraSelecao()
    
    Vdt1 = 0
    Vdt2 = 0
    '---------------------------------------------      Ajuste Verificar Se está disponível para Produção
    L = 2
    P05.Range("CD1") = "PrimeiraSeleção"
    P05.Range("CE1") = "SemanaDeEntrega"
    
    Vdt2 = InputBox("Quantos Meses devo Considerar para Programação da Produção?", "Programação")
    Vdt2 = Month(Date) + Vdt2 - 1
    While P05.Range("A" & L) <> ""
        Parada = P05.Range("BY" & L)
        If Parada = "MAR" Then
            P05.Range("BY" & L).Select
        End If
        VMes (Parada)
        Vdt1 = Parada
        
        If Vdt1 > Vdt2 Then
            P05.Range("CD" & L) = "Não Programar para Produção"
        ElseIf P05.Range("BG" & L) = "OK" Or P05.Range("BH" & L) = "OK" Or P05.Range("BG" & L) = "" Then
            P05.Range("CD" & L) = "Não Programar para Produção"
        Else
            If P05.Range("BE" & L) = "OK" Then
                P05.Range("CD" & L) = "Não Programar para Produção"
            Else
                If P05.Range("BE" & L) = "OK" Then
                    P05.Range("CD" & L) = "Não Programar para Produção"
                ElseIf P05.Range("U" & L) <> "" Then
                    Verif = Left(P05.Range("U" & L), 4)
                    Cont = WorksheetFunction.CountIf(P02.Range("A:A"), P05.Range("W" & L))
                Else
                    P05.Range("CD" & L) = "Não Programar para Produção"
                End If
                
                If P05.Range("U" & L) = "" Then
                    P05.Range("CK1") = P05.Range("BG" & L)
                    P05.Range("CE" & L) = P05.Range("CO1")
                    P05.Range("CD" & L) = "Indisponível para Produção"
                ElseIf Verif > 10 And Cont > 0 Then
                    P05.Range("CK1") = P05.Range("BG" & L)
                    P05.Range("CE" & L) = P05.Range("CO1")
                    P05.Range("CD" & L) = "Disponível para Produção"
                ElseIf Verif > 10 And Cont = 0 Then
                    P05.Range("CK1") = P05.Range("BG" & L)
                    P05.Range("CE" & L) = P05.Range("CO1")
                    P05.Range("CD" & L) = "Indisponível para Produção"
                ElseIf Verif < 10 Then
                    If P05.Range("BG" & L) > Date Then
                        P05.Range("CK1") = P05.Range("BG" & L)
                        P05.Range("CE" & L) = P05.Range("CO1")
                        P05.Range("CD" & L) = "Pendente no Prazo"
                    Else
                        P05.Range("CK1") = P05.Range("BG" & L)
                        P05.Range("CE" & L) = P05.Range("CO1")
                        P05.Range("CD" & L) = "Indisponível para Produção"
                    End If
                ElseIf P05.Range("BD" & L) <> "OK" Then
                    P05.Range("CK1") = P05.Range("BG" & L)
                    P05.Range("CE" & L) = P05.Range("CO1")
                    P05.Range("CD" & L) = "Atraso Acabamento"
                End If
            End If
        End If
        If P05.Range("CD" & L) = "" Then
            P05.Range("CD" & L) = "Entender o que está acontecendo"
        Else
            L = L + 1
        End If
    Wend
    '---------------------------------------------      Ajustar Nome Cliente
    L = 2
    While P05.Range("D" & L) <> ""
        Cod = P05.Range("D" & L)
        Cont = WorksheetFunction.CountIf(P02.Range("N:N"), Cod)
        If Cont > 0 Then
            Lt = 2
            While P02.Range("N" & Lt) <> Cod
                Lt = Lt + 1
            Wend
            P05.Range("D" & L) = P02.Range("O" & Lt)
        Else
            Stop
        End If
        L = L + 1
    Wend
End Sub