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
        Exit Sub
    End If
'-----------------------        Disponibilidade de Peças            --------------------------------------------
    If LstProd(0, 15) = "Atraso Acabamento" Then
        Lprod = P07.Range("A1048576").End(xlUp).Offset(1, 0).Row
'-----------------------        Colocando informação no Relatório de Atrasos do Acabamento
        P07.Range("A" & Lprod) = LstProd(0, 0)                     'Modelo
        P07.Range("B" & Lprod) = LstProd(0, 1)                     'Quantidade
        P07.Range("C" & Lprod) = LstProd(0, 13)                    'Data Previsão
        P07.Range("D" & Lprod) = LstProd(0, 14)                    'Data Deposito
'-----------------------        Colocar Informação no Relatório de Produção e informar que não tem peças disponivel
        Lprod = P08.Range("A1048576").End(xlUp).Offset(1, 0).Row
        P08.Range("A" & Lprod) = LstProd(0, 0)                     'Modelo
        P08.Range("B" & Lprod) = LstProd(0, 1)                     'Quantidade
        P08.Range("C" & Lprod) = LstProd(0, 2)                     'peso
        P08.Range("D" & Lprod) = LstProd(0, 3)                     'Mesa
        P08.Range("E" & Lprod) = LstProd(0, 4)                     'Maquina
        P08.Range("F" & Lprod) = LstProd(0, 5)                     'Cliente
        P08.Range("G" & Lprod) = LstProd(0, 6)                      'Tempo
        P08.Range("H" & Lprod) = LstProd(0, 7)                     'Setup
        P08.Range("K" & Lprod) = LstProd(0, 10)                    'Lote Minimo
        P08.Range("L" & Lprod) = LstProd(0, 11)                    'peça irmão
        P08.Range("M" & Lprod) = LstProd(0, 12)                    'Dt Cart
        P08.Range("N" & Lprod) = LstProd(0, 14)                    'Dt Dep
        P08.Range("O" & Lprod) = "Atraso no acabamento"            'Observação
        P08.Range("R" & Lprod) = LstProd(0, 20)                    'Maquina 2
        Exit Sub
    ElseIf LstProd(0, 15) = "Indisponível para Produção" Then
        Lprod = P08.Range("A1048576").End(xlUp).Offset(1, 0).Row
        P08.Range("A" & Lprod) = LstProd(0, 0)                     'Modelo
        P08.Range("B" & Lprod) = LstProd(0, 1)                     'Quantidade
        P08.Range("C" & Lprod) = LstProd(0, 2)                     'peso
        P08.Range("D" & Lprod) = LstProd(0, 3)                     'Mesa
        P08.Range("E" & Lprod) = LstProd(0, 4)                     'Maquina
        P08.Range("F" & Lprod) = LstProd(0, 5)                     'Cliente
        P08.Range("G" & Lprod) = LstProd(0, 6)                     'Tempo
        P08.Range("H" & Lprod) = LstProd(0, 7)                     'Setup
        P08.Range("K" & Lprod) = FormatNumber(LstProd(0, 10), 0)  'Lote Minimo
        P08.Range("L" & Lprod) = LstProd(0, 11)                   'peça irmão
        P08.Range("M" & Lprod) = LstProd(0, 12)                   'Dt Cart
        P08.Range("N" & Lprod) = LstProd(0, 14)                   'Dt Dep
        P08.Range("O" & Lprod) = "Não Disposnivel para Produção"  'Observação
        P08.Range("R" & Lprod) = LstProd(0, 20)                    'Maquina 2
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
        P08.Range("A" & Lprod) = LstProd(0, 0)                      'Modelo
        P08.Range("B" & Lprod) = LstProd(0, 1)                      'Quantidade
        P08.Range("C" & Lprod) = LstProd(0, 2)                      'peso
        P08.Range("D" & Lprod) = LstProd(0, 3)                      'Mesa
        P08.Range("E" & Lprod) = LstProd(0, 4)                      'Maquina
        P08.Range("F" & Lprod) = LstProd(0, 5)                      'Cliente
        P08.Range("G" & Lprod) = LstProd(0, 6)                      'Tempo
        P08.Range("H" & Lprod) = LstProd(0, 7)                      'Setup
        P08.Range("K" & Lprod) = FormatNumber(LstProd(0, 10), 0)    'Lote Minimo
        P08.Range("L" & Lprod) = LstProd(0, 11)                     'peça irmão
        P08.Range("M" & Lprod) = LstProd(0, 12)                     'Dt Cart
        P08.Range("N" & Lprod) = LstProd(0, 14)                     'Dt Dep
        P08.Range("O" & Lprod) = "Pendente no Prazo"                'Observação
        P08.Range("R" & Lprod) = LstProd(0, 20)                    'Maquina 2
        Exit Sub
    End If
'-----------------------        Realizar Programação Produção       --------------------------------------------
    Lprod = P08.Range("A1048576").End(xlUp).Offset(1, 0).Row        'Posicionar onde vai ser lançada a informação
    P08.Range("A" & Lprod) = LstProd(0, 0)                          'Modelo
    P08.Range("B" & Lprod) = LstProd(0, 1)                          'Quantidade
    P08.Range("C" & Lprod) = LstProd(0, 2)                          'peso
    P08.Range("D" & Lprod) = LstProd(0, 3)                          'Mesa
    P08.Range("E" & Lprod) = LstProd(0, 4)                          'Maquina
    P08.Range("F" & Lprod) = LstProd(0, 5)                          'Cliente
    P08.Range("G" & Lprod) = LstProd(0, 6)                          'Tempo
    P08.Range("H" & Lprod) = LstProd(0, 7)                          'Setup
    P08.Range("I" & Lprod) = P00.Range("J9")                        'Inicio da operação
    P08.Range("J" & Lprod) = P08.Range("I" & Lprod) + _
    (P08.Range("H" & Lprod) + P08.Range("G" & Lprod)) * #12:01:00 AM#
    P08.Range("K" & Lprod) = FormatNumber(LstProd(0, 10), 0)        'Lote Minimo
    P08.Range("L" & Lprod) = LstProd(0, 11)                         'peça irmão
    P08.Range("M" & Lprod) = LstProd(0, 12)                         'Dt Cart
    P08.Range("N" & Lprod) = LstProd(0, 14)                         'Dt Dep
'-----------------------        Atende Lote Mínimo
    Dim Ltmin As Double
    Ltmin = WorksheetFunction.RoundDown((LstProd(0, 10) * P00.Range("J23") / 100), 0)
    If LstProd(0, 1) < Ltmin Then
        P08.Range("O" & Lprod) = "Não Atende Lote Mínimo"
    End If
    P08.Range("R" & Lprod) = LstProd(0, 20)                    'Maquina 2
End Sub