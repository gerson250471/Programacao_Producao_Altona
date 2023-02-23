Sub RelatorioProducao()
    'Obtendo Informações do Modelo
    '====================================
    L = 4
    With P04.PivotTables("tblPCP")
        .PivotCache.Refresh                                                                             'Atualizar tabela
    End With
    
    While P04.Range("A" & L) <> ""
        If P04.Range("A" & L) = "Não Programar para Produção" Then
            L = L + 1
        Else
            Erase LstProd
            Verif = WorksheetFunction.CountIf(P01.Range("A:A"), P04.Range("C" & L))
            If Verif > 0 Then
                LstProd(0, 0) = P04.Range("C" & L)                                                      'Modelo
                LstProd(0, 1) = P04.Range("H" & L)                                                      'Quantidade
                Lproc = P01.Columns("A:A").Find(What:=LstProd(0, 0), _
                LookAt:=xlPart, SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, MatchCase:=False, _
                SearchFormat:=False).Row
                LstProd(0, 2) = P01.Range("E" & Lproc)                                                  'Peso
                LstProd(0, 3) = P01.Range("T" & Lproc)                                                  'Qt Mesa
                LstProd(0, 4) = P01.Range("Y" & Lproc) & "-" & P01.Range("X" & Lproc)                   'Maquina
                LstProd(0, 20) = P01.Range("C" & Lproc).Row                                             'Endereço do modelo
                LstProd(0, 5) = P04.Range("B" & L)                                                      'Cliente
                LstProd(0, 6) = P01.Range("K" & Lproc)                                                  'tempo(Min)
                LstProd(0, 7) = P01.Range("R" & Lproc)                                                  'Setup
                LstProd(0, 8) = P00.Range("J9")                                                         'Hora Inicio
'-----------------------        REALIZAR CONTA PARA SABER PREVISÃO DE CONCLUSÃO
                HRetorno = LstProd(0, 8) + (LstProd(0, 6) * LstProd(0, 1) * #12:01:00 AM#)
                LstProd(0, 9) = HRetorno                                                                'Hora Fim
                LstProd(0, 10) = CInt(P01.Range("L" & Lproc))                                           'Lote min
                LstProd(0, 11) = P01.Range("U" & Lproc)                                                 'Peça Irmã
                LstProd(0, 12) = P04.Range("E" & L)                                                     'Dt Cart
                LstProd(0, 13) = P04.Range("F" & L)                                                     'Dt Dep
                LstProd(0, 14) = P04.Range("G" & L)                                                     'Prev Usinagem
                LstProd(0, 15) = P04.Range("A" & L)                                                     'PRIMEIRA SELEÇÃO
                GerarDemanda
            Else
                LstProd(0, 0) = P04.Range("B" & L)                      'Modelo
                LstProd(0, 1) = P04.Range("H" & L)                      'Quantidade
                Lt = P06.Range("A1048576").End(xlUp).Offset(1, 0).Row   'Posição do Relatório
                P06.Range("A" & Lt) = "Falta Cadastro"
                P06.Range("B" & Lt) = LstProd(0, 0)
                P06.Range("C" & Lt) = LstProd(0, 1)
            End If
        End If
        L = L + 1
    Wend
End Sub