Option Explicit
'Módulo Programar Produção

Sub Main()

'On Error GoTo Erro
'-------------------------  Perguntar ----------------------------------------------
    Resp = MsgBox("Realizar Relatório Produção", vbQuestion + vbYesNo, "Relatório Produção")
    If Resp = 6 Then
    '-------------------------  LOG DE INICIO DE OPERAÇÃO   ------------------------------------
        Llog = P11.Range("A1048576").End(xlUp).Offset(1, 0).Row
        P11.Range("A" & Llog) = "Relatório de Produção"
        P11.Range("B" & Llog) = Now
        Application.DisplayAlerts = False
        Home = ActiveWorkbook.Name
        LimpezaInicial
    '-------------------------  Bucar a informação ----------------------------------------------
        Job = Application.GetOpenFilename(Title:="Escolha o arquivo Inicari Programação do PCP")
        If Job = "Falso" Then
            Exit Sub
        End If
            Workbooks.Open Filename:=Job
            Job = ActiveWorkbook.Name
            Sheets("MPS-UPR").Select
            A = Range("A1048576").End(xlUp).Row
            B = Range("XFD2").End(xlToLeft).Column
            Range(Cells(2, 1), Cells(A, B)).Copy
            Windows(Home).Activate
            'P05.Select
            P05.Range("A1").PasteSpecial xlValues
            P05.Cells.EntireRow.AutoFit
            Windows(Job).Activate
            Application.CutCopyMode = False
            ActiveWorkbook.Close
            Windows(Home).Activate
    '-------------------------  Primeira Seleção  -------------------------
            PrimeiraSelecao
            RelatorioProducao
        End If
    '-------------------------  LOG DE FIM DE OPERAÇÃO   ------------------------------------
        P11.Range("C" & Llog) = Now
        P11.Range("D" & Llog) = P11.Range("C" & Llog) - P11.Range("B" & Llog)
        ActiveWorkbook.Save
        Sheets("Capa").Select
        MsgBox "Primeira Etapa Concluída com Sucesso"
        Exit Sub
Erro:
    MsgBox "Ocorreu um erro durante o processamento da Informação, Favor avisar Programador", vbCritical + vbOKOnly, "Relatório de Produção"
End Sub

Private Sub PrimeiraSelecao()
    Dim Verif As Integer
    
    'Ajuste Verificar Se está disponível para Produção
    L = 2
    P05.Range("CD1") = "PrimeiraSeleção"
    While P05.Range("A" & L) <> ""
        If P05.Range("BG" & L) = "OK" Or P05.Range("BH" & L) = "OK" Or P05.Range("BG" & L) = "" Then
            P05.Range("CD" & L) = "Não Programar para Produção"
        Else
            If P05.Range("BE" & L) = "OK" Then
                P05.Range("CD" & L) = "Não Programar para Produção"
            Else
                If P05.Range("BE" & L) = "OK" Then
                    P05.Range("CD" & L) = "Não Programar para Produção"
                ElseIf P05.Range("U" & L) <> "" Then
                    Verif = Left(P05.Range("U" & L), 4)
                    Cont = WorksheetFunction.CountIf(P13.Range("C:C"), P05.Range("W" & L))
                Else
                    P05.Range("CD" & L) = "Não Programar para Produção"
                End If
                If P05.Range("U" & L) = "" Then
                    P05.Range("CD" & L) = "Indisponível para Produção"
                ElseIf Verif > 10 And Cont > 0 Then
                    P05.Range("CD" & L) = "Disponível para Produção"
                ElseIf Verif > 10 And Cont = 0 Then
                    P05.Range("CD" & L) = "Indisponível para Produção"
                ElseIf Verif < 10 Then
                    If P05.Range("BG" & L) > Date Then
                        P05.Range("CD" & L) = "Pendente no Prazo"
                    Else
                        P05.Range("CD" & L) = "Indisponível para Produção"
                    End If
                ElseIf P05.Range("BD" & L) <> "OK" Then
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
End Sub

Sub RelatorioProducao()
    'Obtendo Informações do Modelo
    '====================================
    L = 4
    With P04.PivotTables("tblPCP")
        .PivotCache.Refresh                                                         'Atualizar tabela
    End With
    
    While P04.Range("A" & L) <> ""
        If P04.Range("A" & L) = "Não Programar para Produção" Then
            L = L + 1
        Else
            Erase LstProd
            Verif = WorksheetFunction.CountIf(P01.Range("A:A"), P04.Range("B" & L))
            If Verif > 0 Then
                LstProd(0, 0) = P04.Range("B" & L)                      'Modelo
                LstProd(0, 1) = P04.Range("F" & L)                      'Quantidade
                LstProd(0, 13) = P04.Range("D" & L)                     'Prev Recu
                LstProd(0, 12) = P04.Range("C" & L)                     'Dt Cart
                LstProd(0, 14) = P04.Range("E" & L)                     'Prev Dep
                LstProd(0, 15) = P04.Range("A" & L)                     'Peças para Produzir
                EncontrarModelo
                GerarDemanda
                L = L + 1
            Else
                LstProd(0, 0) = P04.Range("B" & L)                      'Modelo
                LstProd(0, 1) = P04.Range("F" & L)                      'Quantidade
                Lt = P06.Range("A1048576").End(xlUp).Offset(1, 0).Row   'Posição do Relatório
                P06.Range("A" & Lt) = "Falta Cadastro"
                P06.Range("B" & Lt) = LstProd(0, 0)
                P06.Range("C" & Lt) = LstProd(0, 1)
                L = L + 1
            End If
        End If
    Wend
End Sub

Private Sub GerarDemanda()
'-----------------------        Não Programar Produção              -------------------------------------------
    If LstProd(0, 15) = "Não Programar para Produção" Then
        Exit Sub
    End If
'-----------------------        Informação de Máquina               -------------------------------------------
    If LstProd(0, 4) = "-" Or LstProd(0, 4) = "" Then
        Lprod = P06.Range("A1048576").End(xlUp).Offset(1, 0).Row
        P06.Range("A" & Lprod) = "Falta informação de Máquina de Usinagem"
        P06.Range("B" & Lprod) = LstProd(0, 0)
        P06.Range("C" & Lprod) = LstProd(0, 1)
        Exit Sub
    End If
'-----------------------        Informação de MESA                  -------------------------------------------
    If LstProd(0, 3) = "" Then
        Lprod = P06.Range("A1048576").End(xlUp).Offset(1, 0).Row
        P06.Range("A" & Lprod) = "Falta informação de Quantidade de Mesa"
        P06.Range("B" & Lprod) = LstProd(0, 0)
        P06.Range("C" & Lprod) = LstProd(0, 1)
        ActiveCell.Offset(1, 0).Select
        Exit Sub
    End If
'-----------------------        Disponibilidade de Peças            --------------------------------------------
    If LstProd(0, 15) = "Atraso Acabamento" Then
        Lprod = P07.Range("A1048576").End(xlUp).Offset(1, 0).Row
        'Colocando informação no Relatório de Atrasos do Acabamento
        P07.Range("A" & Lprod) = LstProd(0, 0)                     'Modelo
        P07.Range("B" & Lprod) = LstProd(0, 1)                     'Quantidade
        P07.Range("C" & Lprod) = LstProd(0, 13)                    'Data Previsão
        P07.Range("D" & Lprod) = LstProd(0, 14)                    'Data Deposito
        'Colocar Informação no Relatório de Produção e informar que não tem peças disponivel
        Lprod = P08.Range("A1048576").End(xlUp).Offset(1, 0).Row
        P08.Range("A" & Lprod) = LstProd(0, 0)                     'Modelo
        P08.Range("B" & Lprod) = LstProd(0, 1)                     'Quantidade
        P08.Range("C" & Lprod) = LstProd(0, 2) * LstProd(0, 1)     'peso
        P08.Range("D" & Lprod) = LstProd(0, 3)                     'Mesa
        P08.Range("E" & Lprod) = LstProd(0, 4)                     'Maquina
        P08.Range("F" & Lprod) = LstProd(0, 5)                     'Cliente
        P08.Range("G" & Lprod) = LstProd(0, 6) * LstProd(0, 1)     'Tempo
        P08.Range("H" & Lprod) = LstProd(0, 9)                     'Setup
        P08.Range("K" & Lprod) = FormatNumber(LstProd(0, 10), 0)  'Lote Minimo
        P08.Range("L" & Lprod) = LstProd(0, 11)                   'peça irmão
        P08.Range("M" & Lprod) = LstProd(0, 12)                   'Dt Cart
        P08.Range("N" & Lprod) = LstProd(0, 14)                   'Dt Dep
        P08.Range("O" & Lprod) = "Atraso no acabamento"  'Observação
        Exit Sub
    ElseIf LstProd(0, 15) = "Indisponível para Produção" Then
        Lprod = P08.Range("A1048576").End(xlUp).Offset(1, 0).Row
        P08.Range("A" & Lprod) = LstProd(0, 0)                     'Modelo
        P08.Range("B" & Lprod) = LstProd(0, 1)                     'Quantidade
        P08.Range("C" & Lprod) = LstProd(0, 2) * LstProd(0, 1)     'peso
        P08.Range("D" & Lprod) = LstProd(0, 3)                     'Mesa
        P08.Range("E" & Lprod) = LstProd(0, 4)                     'Maquina
        P08.Range("F" & Lprod) = LstProd(0, 5)                     'Cliente
        P08.Range("G" & Lprod) = LstProd(0, 6) * LstProd(0, 1)     'Tempo
        P08.Range("H" & Lprod) = LstProd(0, 9)                     'Setup
        P08.Range("K" & Lprod) = FormatNumber(LstProd(0, 10), 0)  'Lote Minimo
        P08.Range("L" & Lprod) = LstProd(0, 11)                   'peça irmão
        P08.Range("M" & Lprod) = LstProd(0, 12)                   'Dt Cart
        P08.Range("N" & Lprod) = LstProd(0, 14)                   'Dt Dep
        P08.Range("O" & Lprod) = "Não Disposnivel para Produção"  'Observação
        Exit Sub
    End If
    If LstProd(0, 6) = "" Then
'-----------------------         Tempo de Produção                  --------------------------------------------
        Lprod = P06.Range("A1048576").End(xlUp).Offset(1, 0).Row
        P06.Range("A" & Lprod) = "Falta informação de Tempo de Usinagem"
        P06.Range("B" & Lprod) = LstProd(0, 0)
        P06.Range("C" & Lprod) = LstProd(0, 1)
        Exit Sub
    End If
'-----------------------        Pendente no Prazo                   --------------------------------------------
    If LstProd(0, 15) = "Pendente no Prazo" Then
        Lprod = P08.Range("A1048576").End(xlUp).Offset(1, 0).Row    'Posicionar onde vai ser lançada a informação
        P08.Select
        P08.Range("A" & Lprod).Select
        P08.Range("A" & Lprod) = LstProd(0, 0)                      'Modelo
        P08.Range("B" & Lprod) = LstProd(0, 1)                      'Quantidade
        P08.Range("C" & Lprod) = LstProd(0, 2) * LstProd(0, 1)      'peso
        P08.Range("D" & Lprod) = LstProd(0, 3)                      'Mesa
        P08.Range("E" & Lprod) = LstProd(0, 4)                      'Maquina
        P08.Range("F" & Lprod) = LstProd(0, 5)                      'Cliente
        P08.Range("G" & Lprod) = LstProd(0, 6) * LstProd(0, 1)      'Tempo
        P08.Range("H" & Lprod) = LstProd(0, 9)                      'Setup
        P08.Range("K" & Lprod) = FormatNumber(LstProd(0, 10), 0)    'Lote Minimo
        P08.Range("L" & Lprod) = LstProd(0, 11)                     'peça irmão
        P08.Range("M" & Lprod) = LstProd(0, 12)                     'Dt Cart
        P08.Range("N" & Lprod) = LstProd(0, 14)                     'Dt Dep
        P08.Range("O" & Lprod) = "Pendente no Prazo"                'Observação
        Exit Sub
    End If
'-----------------------        Realizar Programação Produção       --------------------------------------------
    Lprod = P08.Range("A1048576").End(xlUp).Offset(1, 0).Row        'Posicionar onde vai ser lançada a informação
    P08.Range("A" & Lprod) = LstProd(0, 0)                          'Modelo
    P08.Range("B" & Lprod) = LstProd(0, 1)                          'Quantidade
    P08.Range("C" & Lprod) = LstProd(0, 2) * LstProd(0, 1)          'peso
    P08.Range("D" & Lprod) = LstProd(0, 3)                          'Mesa
    P08.Range("E" & Lprod) = LstProd(0, 4)                          'Maquina
    P08.Range("F" & Lprod) = LstProd(0, 5)                          'Cliente
    P08.Range("G" & Lprod) = LstProd(0, 6) * LstProd(0, 1)          'Tempo
    P08.Range("H" & Lprod) = LstProd(0, 9)                          'Setup
    P08.Range("I" & Lprod) = P00.Range("J9")                        'Inicio da operação
    P08.Range("J" & Lprod) = P08.Range("I" & Lprod) + _
    (P08.Range("H" & Lprod) + P08.Range("G" & Lprod)) * #12:01:00 AM#
    P08.Range("K" & Lprod) = FormatNumber(LstProd(0, 10), 0)        'Lote Minimo
    P08.Range("L" & Lprod) = LstProd(0, 11)                         'peça irmão
    P08.Range("M" & Lprod) = LstProd(0, 12)                         'Dt Cart
    P08.Range("N" & Lprod) = LstProd(0, 14)                         'Dt Dep
    'Informa atraso
    If P08.Range("j" & Lprod) > P08.Range("N" & Lprod) Then
        Range(P08.Range("A" & Lprod), P08.Range("Q" & Lprod)).Font.Color = vbRed
        Range(P08.Range("A" & Lprod), P08.Range("Q" & Lprod)).Font.Bold = True
        Range(P08.Range("A" & Lprod), P08.Range("Q" & Lprod)).Font.Size = 12
    Else
        Range(P08.Range("A" & Lprod), P08.Range("Q" & Lprod)).Font.Color = vbBlack
        Range(P08.Range("A" & Lprod), P08.Range("Q" & Lprod)).Font.Bold = False
        Range(P08.Range("A" & Lprod), P08.Range("Q" & Lprod)).Font.Size = 12
    End If
    'Atende Lote Mínimo
    Ltmin = WorksheetFunction.RoundDown((LstProd(0, 10) * P00.Range("J23") / 100), 0)
    If LstProd(0, 1) < Ltmin Then
        P08.Range("O" & Lprod) = "Não Atende Lote Mínimo"
    End If
    
End Sub

Private Sub LimpezaInicial()
    P03.Range("A2:EZ70000") = ""
    P05.Range("A2:EZ70000") = ""
    P06.AutoFilterMode = False
    P06.Range("A5:Z30000").ClearContents
    P07.AutoFilterMode = False
    P07.Range("A5:Z30000").ClearContents
    P08.AutoFilterMode = False
    P08.Range("A5:Z30000").ClearContents
    P09.AutoFilterMode = False
    P09.Range("A5:Z30000").ClearContents
    P10.AutoFilterMode = False
    P10.Range("A5:Z30000").ClearContents
    
End Sub

Private Sub EncontrarModelo()
    Dim Ver As Integer
    
    Lproc = P01.Columns("A:A").Find(What:=LstProd(0, 0), LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Row
        
    LstProd(0, 2) = P01.Range("E" & Lproc)                                          'Peso
    LstProd(0, 3) = P01.Range("T" & Lproc)                                          'Mesa
    LstProd(0, 4) = P01.Range("Y" & Lproc) & "-" & P01.Range("X" & Lproc)           'Maquina 1
    LstProd(0, 5) = P01.Range("G" & Lproc)                                          'Cliente
    LstProd(0, 6) = P01.Range("K" & Lproc)                                          'Tempo
    LstProd(0, 9) = P01.Range("R" & Lproc)                                          'Setup
    LstProd(0, 10) = P01.Range("L" & Lproc)                                         'Lote minimo
    LstProd(0, 11) = P01.Range("U" & Lproc)                                         'Modelo Irmão
      
End Sub


