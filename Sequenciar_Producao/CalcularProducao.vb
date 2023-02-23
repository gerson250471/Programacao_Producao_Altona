Private Sub CalcularProducao()
Dim Verif As Double
    If IsNull(LstProd(A + 1, 6)) Then
        LstProd(A + 1, 6) = 0
    End If
    
    If Cont > 1 Then
        Verif = WorksheetFunction.RoundDown(LstProd(A, 8), 0)
        Verif = LstProd(A, 8) - Verif
        If Verif = 0 Then
            CapProd = 0.9375                                                'Capacidade de Produção 22:30
        Else
            CapProd = 0.9375 - Verif
        End If
    
        If LstProd(A, 11) <> "-" Then
            '----------------------     Tempo de Produção e tempo de espera pelo segundo Modelto
            Prod = (LstProd(A, 6) + LstProd(A + 1, 6)) * #12:01:00 AM#
            '----------------------     Verificando se existe parada de fim de semana
        ElseIf LstProd(A, 11) = "-" Then
            If B = 0 Then
                '----------------------     Tempo de Produção e tempo de espera pelo segundo Modelto
                Prod = (LstProd(A, 6) + LstProd(A + 1, 6)) * #12:01:00 AM#
                '----------------------     Verificando se existe parada de fim de semana
            End If
        End If
        If LstProd(A, 3) >= 2 Then
            Prod = LstProd(A, 6) * #12:01:00 AM#
        End If
        Dt = Weekday(LstProd(A, 8))
    Else
        Verif = WorksheetFunction.RoundDown(LstProd(0, 8), 0)
        Verif = LstProd(0, 8) - Verif
        If Verif = 0 Then
            CapProd = 0.9375                                                'Capacidade de Produção 22:30
        Else
            CapProd = 0.9375 - Verif
        End If
    
        '----------------------     Tempo de Produção e tempo de espera pelo segundo Modelto
        Prod = LstProd(0, 6) * #12:01:00 AM#
        '----------------------     Verificando se existe parada de fim de semana
        Dt = Weekday(LstProd(A, 8))
    End If
End Sub