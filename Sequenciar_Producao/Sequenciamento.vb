Option Explicit
'-----------------------    Área de Variáveis do Sistema
    Public Rever()
    Public Seq(8, 20)               As Variant
    Public LstProd(200, 30)         As Variant
    Public Mesa(2, 4)               As Variant
    Public Pego                     As Integer
    Public RelImpresso(10, 10)      As Variant
    Public Dini, Dfim               As Integer
    Public CapProd                  As Double                                                                                                                                                           'Capacidade para Produzir no dia
    Public HParada                  As Date                                                                                                                                                             'Hora da Parada
    Public HRetorno                 As Date                                                                                                                                                             'Hora do Retorno
    Public Parada                   As String                                                                                                                                                           'Parada Fim de Semana
    Public Horizonte                As Integer, Resp        As Integer                                                                                                                                  'Dias para Programar
    Dim Vr1                         As Double, Vr2          As Double
    Public ProdDia                  As Double, Refeicao     As Double                                                                                                                                   'Produção do Dia
    Public Home                     As String, Job          As String, Cod  As String                                                                                                                   'Variáveis de posicionamento
    Public Prod                     As Double, Dt           As Double, Tp   As Double, Ltmin    As Double, Tempo    As Double                                                                           'Quantidade para Produzir
    Public L                        As Double, Llog         As Double, Lt   As Double, Lprod    As Double, Lproc    As Double, Lcapa    As Double, Lrel As Double                                       'Posicionamento
    Public A                        As Double, B            As Double, C    As Double, D        As Double, N        As Double, P        As Double, Enc  As Double, Verif As Double, Cont As Double      'Variáveis para contagem
    
'------------------------------------------------------------------------------------------------
Sub Sequenciamento()
'On Error GoTo Erro
'------------------------------------------------------------------------------------------------------------------------------------  Verificar se programar modelo
Resp = MsgBox("Gerar Sequenciamento de Produção", vbQuestion + vbYesNo, "Sequenciamento")
If Resp = 6 Then
'------------------------------------------------------------------------------------------------------------------------------------  LOG DE INICIO DE OPERAÇÃO
        Llog = P11.Range("A1048576").End(xlUp).Offset(1, 0).Row
        P11.Range("A" & Llog) = "Sequenciamento de Produção"
        P11.Range("B" & Llog) = Date
        Tempo = Timer
        Application.DisplayAlerts = False
        Home = ActiveWorkbook.Name
        LimpezaInicial
'------------------------------------------------------------------------------------------------------------------------------------  Bucar a informação
    If P00.Range("J25") <> "" Then
        A = 0
        Erase LstProd
        LstProd(A, 0) = P00.Range("J25")
        Lproc = P01.Columns("A:A").Find(What:=LstProd(A, 0), _
                LookAt:=xlPart, SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, MatchCase:=False, _
                SearchFormat:=False).Row
        If P01.Range("U" & Lproc) <> "-" Then   '====================================================================================> Tem modelo irmão
            Cont = 2
            LstProd(A, 1) = WorksheetFunction.SumIfs(P04.Range("G:G"), _
                            P04.Range("C:C"), LstProd(A, 0), P04.Range("A:A"), _
                            "Disponível para Produção")                                                                                 'Qt
            LstProd(A, 2) = P01.Range("E" & Lproc)                                                                                      'Peso
            LstProd(A, 3) = P01.Range("T" & Lproc)                                                                                      'Qt Mesa
            LstProd(A, 4) = P01.Range("Y" & Lproc) & "-" & P01.Range("X" & Lproc)                                                       'Maquina
            LstProd(A, 5) = P01.Range("G" & Lproc)                                                                                      'Cliente
            LstProd(A, 6) = P01.Range("K" & Lproc)                                                                                      'tempo(Min)
            LstProd(A, 7) = P01.Range("R" & Lproc)                                                                                      'Setup
            LstProd(A, 8) = P00.Range("J9")                                                                                             'Hora Inicio
            LstProd(A, 10) = P01.Range("L" & Lproc)                                                                                     'Lote min
            LstProd(A, 11) = P01.Range("U" & Lproc)                                                                                     'Peça Irmã
            LstProd(A, 12) = WorksheetFunction.SumIfs(P04.Range("D:D"), _
                            P04.Range("C:C"), LstProd(A, 0), P04.Range("A:A"), _
                            "Disponível para Produção")                                                                                 'Dt Cart
            LstProd(A, 13) = WorksheetFunction.SumIfs(P04.Range("F:F"), _
                            P04.Range("C:C"), LstProd(A, 0), P04.Range("A:A"), _
                            "Disponível para Produção")                                                                                 'Dt Dep
            LstProd(A + 1, 0) = P01.Range("U" & Lproc)  '==============================================================================> MODELO IRMÃO
            Lproc = P01.Columns("A:A").Find(What:=LstProd(A + 1, 0), _
                LookAt:=xlPart, SearchOrder:=xlByRows, _
                SearchDirection:=xlNext, MatchCase:=False, _
                SearchFormat:=False).Row
            LstProd(A + 1, 1) = WorksheetFunction.SumIfs(P04.Range("G:G"), _
                            P04.Range("C:C"), LstProd(A, 0), P04.Range("A:A"), _
                            "Disponível para Produção")                                                                                 'Qt
            LstProd(A + 1, 2) = P01.Range("E" & Lproc)                                                                                  'Peso
            LstProd(A + 1, 3) = P01.Range("T" & Lproc)                                                                                  'Qt Mesa
            LstProd(A + 1, 4) = P01.Range("Y" & Lproc) & "-" & P01.Range("X" & Lproc)                                                   'Maquina
            LstProd(A + 1, 5) = P01.Range("G" & Lproc)                                                                                  'Cliente
            LstProd(A + 1, 6) = P01.Range("K" & Lproc)                                                                                  'tempo(Min)
            LstProd(A + 1, 7) = P01.Range("R" & Lproc)                                                                                  'Setup
            LstProd(A + 1, 8) = P00.Range("J9")                                                                                         'Hora Inicio
            LstProd(A + 1, 10) = P01.Range("L" & Lproc)                                                                                 'Lote min
            LstProd(A + 1, 11) = P01.Range("U" & Lproc)                                                                                 'Peça Irmã
            LstProd(A + 1, 12) = WorksheetFunction.SumIfs(P04.Range("D:D"), _
                            P04.Range("C:C"), LstProd(A, 0), P04.Range("A:A"), _
                            "Disponível para Produção")                                                                                 'Dt Cart
            LstProd(A + 1, 13) = WorksheetFunction.SumIfs(P04.Range("F:F"), _
                            P04.Range("C:C"), LstProd(A, 0), P04.Range("A:A"), _
                            "Disponível para Produção")                                                                                 'Dt Dep
            '-------------------------------------------------------------------------------------------------------------------------  Informando no Relatório De Programação Geral de Produção
            L = P08.Range("A1048576").End(xlUp).Offset(1, 0).Row
            For A = 0 To 1
                P08.Range("A" & L) = LstProd(A, 0)                                                                                      'Modelo
                P08.Range("B" & L) = LstProd(A, 1)                                                                                      'Qt
                P08.Range("C" & L) = LstProd(A, 2)                                                                                      'Peso
                P08.Range("D" & L) = LstProd(A, 3)                                                                                      'Qt Mesa
                P08.Range("E" & L) = LstProd(A, 4)                                                                                      'Maquina
                P08.Range("F" & L) = LstProd(A, 5)                                                                                      'Cliente
                P08.Range("G" & L) = LstProd(A, 6)                                                                                      'tempo(Min)
                P08.Range("H" & L) = LstProd(A, 7)                                                                                      'Setup
                P08.Range("I" & L) = LstProd(A, 8)                                                                                      'Hora Inicio
                LstProd(A, 9) = (P08.Range("G" & L) * P08.Range("B" & L) + P08.Range("H" & L)) * #12:01:00 AM# + P08.Range("I" & L)     'calculo de termino do trabalho
                P08.Range("J" & L) = LstProd(A, 9)                                                                                      'Hora Fim
                P08.Range("K" & L) = LstProd(A, 10)                                                                                     'Lote min
                P08.Range("L" & L) = LstProd(A, 11)                                                                                     'Peça Irmã
                P08.Range("M" & L) = LstProd(A, 12)                                                                                     'Dt Cart
                P08.Range("N" & L) = LstProd(A, 13)                                                                                     'Dt Dep
                P08.Range("O" & L) = "Programar Modelo"                                                                                 'Observação
                P08.Range("P" & L) = "Programado"
                L = L + 1
            Next A

            For A = 0 To 1
                P = 0
                N = 0
                L = P10.Range("A1048576").End(xlUp).Offset(1, 0).Row
                While P < LstProd(0, 1)
                    If N = 0 Then       '===========================================================================================>   SETUP DE MÁQUINA
                        If Mesa(0, 0) > Mesa(1, 0) Then
                            LstProd(A, 8) = Mesa(1, 0)
                            Pego = 1
                        Else
                            LstProd(A, 8) = Mesa(0, 0)
                            Pego = 0
                        End If
                        LstProd(A, 18) = LstProd(A, 8)
                        P10.Range("A" & L) = LstProd(A, 4)                                                                              'Máquina
                        P10.Range("B" & L) = LstProd(A, 0)                                                                              'Modelo
                        P10.Range("C" & L) = LstProd(A, 5)                                                                              'Cliente
                        P10.Range("D" & L) = LstProd(A, 12)                                                                             'Data Carteira
                        P10.Range("E" & L) = LstProd(A, 13)                                                                             'Data Depósito
                        P10.Range("F" & L) = "Setup de Máquina"                                                                         'Descrição
                        P10.Range("G" & L) = LstProd(A, 3)                                                                              'Qt mesa
                        P10.Range("H" & L) = LstProd(A, 2)                                                                              'Peso
                        P10.Range("I" & L) = 0                                                                                          'Qt
                        P10.Range("J" & L) = LstProd(A, 1)                                                                              'Qt programada
                        P10.Range("K" & L) = LstProd(A, 8)                                                                              'Inicio
                        P10.Range("L" & L) = LstProd(A, 7) * #12:01:00 AM#                                                              'DURAÇÃO
                        Duracao
                        LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao                                              'calculo para termino do trabalho
                        P10.Range("M" & L) = LstProd(A, 9)                                                                              'Termino
                        LstProd(A, 8) = LstProd(A, 9)                                                                                   'Novo Inicio
                        LstProd(1, 8) = LstProd(A, 9)                                                                                   'Inicio Segunda mesa
                        N = N + 1
                        L = L + 1

                    Else                '==========================================================================================>    SEQUENCIAR PRODUÇÃO
                        P10.Range("A" & L) = LstProd(A, 4)                                                                              'Máquina
                        P10.Range("B" & L) = LstProd(A, 0)                                                                              'Modelo
                        P10.Range("C" & L) = LstProd(A, 5)                                                                              'Cliente
                        P10.Range("D" & L) = LstProd(A, 12)                                                                             'Data Carteira
                        P10.Range("E" & L) = LstProd(A, 13)                                                                             'Data Depósito
                        P10.Range("F" & L) = "Produção Dia " & Format(N, "00")                                                          'Descrição
                        P10.Range("G" & L) = LstProd(A, 3)                                                                              'Qt mesa
                        P10.Range("H" & L) = LstProd(A, 2)                                                                              'Peso
                        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        CalcularProducao
                        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        If P00.Range("H16") = Dt And P00.Range("J11") = "Sim" Then
                            ProdDia = WorksheetFunction.RoundUp(0.5625 / Prod, 0)
                            If (ProdDia + P) > LstProd(A, 1) Then
                                ProdDia = LstProd(A, 1) - P
                                P10.Range("I" & L) = ProdDia                                                                        'Qt
                                P10.Range("J" & L) = LstProd(A, 1)                                                                  'Qt programada
                                P10.Range("K" & L) = LstProd(A, 8)                                                                  'Inicio
                                P10.Range("L" & L) = ProdDia * Prod                                                                 'DURAÇÃO
                                Duracao
                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                P10.Range("M" & L) = LstProd(A, 9)                                                                  'Termino
                                LstProd(A, 8) = LstProd(A, 9)                                                                       'Novo Inicio
                                N = N + 1
                                P = P + ProdDia
                                L = L + 1
                            Else
                                P10.Range("I" & L) = ProdDia                                                                        'Qt
                                P10.Range("J" & L) = LstProd(A, 1)                                                                  'Qt programada
                                P10.Range("K" & L) = LstProd(A, 8)                                                                  'Inicio
                                P10.Range("L" & L) = ProdDia * Prod                                                                 'DURAÇÃO
                                Duracao
                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                P10.Range("M" & L) = LstProd(A, 9)                                                                  'Termino
                                LstProd(A, 8) = WorksheetFunction.RoundUp(LstProd(A, 9), 0) + 0.9166                                'Novo Inicio
                                N = N + 1
                                P = P + ProdDia
                                L = L + 1
                            End If
                        Else
                            ProdDia = WorksheetFunction.RoundDown(CapProd / Prod, 0)
                            If (ProdDia + P) > LstProd(A, 1) Then
                                ProdDia = LstProd(A, 1) - P
                                P10.Range("I" & L) = ProdDia                                                                        'Qt
                                P10.Range("J" & L) = LstProd(A, 1)                                                                  'Qt programada
                                P10.Range("K" & L) = LstProd(A, 8)                                                                  'Inicio
                                P10.Range("L" & L) = ProdDia * Prod                                                                 'DURAÇÃO
                                Duracao
                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                P10.Range("M" & L) = LstProd(A, 9)                                                                  'Termino
                                LstProd(A, 8) = LstProd(A, 9)                                                                       'Novo Inicio
                                N = N + 1
                                P = P + ProdDia
                                L = L + 1
                            Else
                                If ProdDia = 0 Then P10.Range("I" & L) = 1 Else P10.Range("I" & L) = ProdDia                        'Qt
                                P10.Range("J" & L) = LstProd(A, 1)                                                                  'Qt programada
                                P10.Range("K" & L) = LstProd(A, 8)                                                                  'Inicio
                                If ProdDia = 0 Then P10.Range("L" & L) = 1 * Prod Else P10.Range("L" & L) = ProdDia * Prod          'DURAÇÃO
                                Duracao
                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                P10.Range("M" & L) = LstProd(A, 9)                                                                  'Termino
                                LstProd(A, 8) = LstProd(A, 9)                                                                       'Novo Inicio
                                N = N + 1
                                P = P + ProdDia
                                L = L + 1
                            End If
                        End If
                    End If
                Wend
            Next A
            Enc = False
        Else
            LstProd(A, 1) = WorksheetFunction.SumIfs(P04.Range("G:G"), _
                            P04.Range("C:C"), LstProd(A, 0), P04.Range("A:A"), _
                            "Disponível para Produção")                                                                                 'Qt
            LstProd(A, 2) = P01.Range("E" & Lproc)                                                                                      'Peso
            LstProd(A, 3) = P01.Range("T" & Lproc)                                                                                      'Qt Mesa
            LstProd(A, 4) = P01.Range("Y" & Lproc) & "-" & P01.Range("X" & Lproc)                                                       'Maquina
            LstProd(A, 5) = P01.Range("G" & Lproc)                                                                                      'Cliente
            LstProd(A, 6) = P01.Range("K" & Lproc)                                                                                      'tempo(Min)
            LstProd(A, 7) = P01.Range("R" & Lproc)                                                                                      'Setup
            LstProd(A, 8) = P00.Range("J9")                                                                                             'Hora Inicio
            LstProd(A, 10) = P01.Range("L" & Lproc)                                                                                     'Lote min
            LstProd(A, 11) = P01.Range("U" & Lproc)                                                                                     'Peça Irmã
            LstProd(A, 12) = WorksheetFunction.SumIfs(P04.Range("D:D"), _
                            P04.Range("C:C"), LstProd(A, 0), P04.Range("A:A"), _
                            "Disponível para Produção")                                                                                 'Dt Cart
            LstProd(A, 13) = WorksheetFunction.SumIfs(P04.Range("F:F"), _
                            P04.Range("C:C"), LstProd(A, 0), P04.Range("A:A"), _
                            "Disponível para Produção")                                                                                 'Dt Dep
        '-----------------------------------------------------------------------------------------------------------------------------  Informando no Relatório De Programação Geral de Produção
        L = P08.Range("A1048576").End(xlUp).Offset(1, 0).Row
        P08.Range("A" & L) = LstProd(A, 0)                                                                                              'Modelo
        P08.Range("B" & L) = LstProd(A, 1)                                                                                              'Qt
        P08.Range("C" & L) = LstProd(A, 2)                                                                                              'Peso
        P08.Range("D" & L) = LstProd(A, 3)                                                                                              'Qt Mesa
        P08.Range("E" & L) = LstProd(A, 4)                                                                                              'Maquina
        P08.Range("F" & L) = LstProd(A, 5)                                                                                              'Cliente
        P08.Range("G" & L) = LstProd(A, 6)                                                                                              'tempo(Min)
        P08.Range("H" & L) = LstProd(A, 7)                                                                                              'Setup
        P08.Range("I" & L) = LstProd(A, 8)                                                                                              'Hora Inicio
        LstProd(A, 9) = (P08.Range("G" & L) * P08.Range("B" & L) + P08.Range("H" & L)) * #12:01:00 AM# + P08.Range("I" & L)             'calculo de termino do trabalho
        P08.Range("J" & L) = LstProd(A, 9)                                                                                              'Hora Fim
        P08.Range("K" & L) = LstProd(A, 10)                                                                                             'Lote min
        P08.Range("L" & L) = LstProd(A, 11)                                                                                             'Peça Irmã
        P08.Range("M" & L) = LstProd(A, 12)                                                                                             'Dt Cart
        P08.Range("N" & L) = LstProd(A, 13)                                                                                             'Dt Dep
        P08.Range("O" & L) = "Programar Modelo"                                                                                         'Observação
        P08.Range("P" & L) = "Programado"
        P = 0
        N = 0
        L = P10.Range("A1048576").End(xlUp).Offset(1, 0).Row
        While P < LstProd(0, 1)
            If N = 0 Then       '==================================================================================================>    SETUP DE MÁQUINA
                If Mesa(0, 0) > Mesa(1, 0) Then
                    LstProd(A, 8) = Mesa(1, 0)
                    Pego = 1
                Else
                    LstProd(A, 8) = Mesa(0, 0)
                    Pego = 0
                End If
                LstProd(A, 18) = LstProd(A, 8)
                P10.Range("A" & L) = LstProd(A, 4)                                                                                      'Máquina
                P10.Range("B" & L) = LstProd(A, 0)                                                                                      'Modelo
                P10.Range("C" & L) = LstProd(A, 5)                                                                                      'Cliente
                P10.Range("D" & L) = LstProd(A, 12)                                                                                     'Data Carteira
                P10.Range("E" & L) = LstProd(A, 13)                                                                                     'Data Depósito
                P10.Range("F" & L) = "Setup de Máquina"                                                                                 'Descrição
                P10.Range("G" & L) = LstProd(A, 3)                                                                                      'Qt mesa
                P10.Range("H" & L) = LstProd(A, 2)                                                                                      'Peso
                P10.Range("I" & L) = 0                                                                                                  'Qt
                P10.Range("J" & L) = LstProd(A, 1)                                                                                      'Qt programada
                P10.Range("K" & L) = LstProd(A, 8)                                                                                      'Inicio
                P10.Range("L" & L) = LstProd(A, 7) * #12:01:00 AM#                                                                      'DURAÇÃO
                Duracao
                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao                                                      'calculo para termino do trabalho
                P10.Range("M" & L) = LstProd(A, 9)                                                                                      'Termino
                LstProd(A, 8) = LstProd(A, 9)                                                                                           'Novo Inicio
                LstProd(1, 8) = LstProd(A, 9)                                                                                           'Inicio Segunda mesa
                N = N + 1
                L = L + 1
            Else                '==================================================================================================>    SEQUENCIAR PRODUÇÃO
                P10.Range("A" & L) = LstProd(A, 4)                                                                                      'Máquina
                P10.Range("B" & L) = LstProd(A, 0)                                                                                      'Modelo
                P10.Range("C" & L) = LstProd(A, 5)                                                                                      'Cliente
                P10.Range("D" & L) = LstProd(A, 12)                                                                                     'Data Carteira
                P10.Range("E" & L) = LstProd(A, 13)                                                                                     'Data Depósito
                P10.Range("F" & L) = "Produção Dia " & Format(N, "00")                                                                  'Descrição
                P10.Range("G" & L) = LstProd(A, 3)                                                                                      'Qt mesa
                P10.Range("H" & L) = LstProd(A, 2)                                                                                      'Peso
                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                CalcularProducao
                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                If P00.Range("H16") = Dt And P00.Range("J11") = "Sim" Then
                    ProdDia = WorksheetFunction.RoundDown(0.5625 / Prod, 0)
                    If (ProdDia + P) > LstProd(A, 1) Then
                        ProdDia = LstProd(A, 1) - P
                        P10.Range("I" & L) = ProdDia                                                                                'Qt
                        P10.Range("J" & L) = LstProd(A, 1)                                                                          'Qt programada
                        P10.Range("K" & L) = LstProd(A, 8)                                                                          'Inicio
                        P10.Range("L" & L) = ProdDia * Prod                                                                         'DURAÇÃO
                        Duracao
                        LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                        P10.Range("M" & L) = LstProd(A, 9)                                                                          'Termino
                        LstProd(A, 8) = LstProd(A, 9)                                                                               'Novo Inicio
                        N = N + 1
                        P = P + ProdDia
                        L = L + 1
                    Else
                        P10.Range("I" & L) = ProdDia                                                                                'Qt
                        P10.Range("J" & L) = LstProd(A, 1)                                                                          'Qt programada
                        P10.Range("K" & L) = LstProd(A, 8)                                                                          'Inicio
                        P10.Range("L" & L) = ProdDia * Prod                                                                         'DURAÇÃO
                        Duracao
                        LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                        P10.Range("M" & L) = LstProd(A, 9)                                                                          'Termino
                        LstProd(A, 8) = WorksheetFunction.RoundUp(LstProd(A, 9), 0) + 0.9166                                        'Novo Inicio
                        N = N + 1
                        P = P + ProdDia
                        L = L + 1
                    End If
                Else
                    ProdDia = WorksheetFunction.RoundDown(CapProd / Prod, 0)
                    If (ProdDia + P) > LstProd(A, 1) Then
                        ProdDia = LstProd(A, 1) - P
                        P10.Range("I" & L) = ProdDia                                                                                'Qt
                        P10.Range("J" & L) = LstProd(A, 1)                                                                          'Qt programada
                        P10.Range("K" & L) = LstProd(A, 8)                                                                          'Inicio
                        P10.Range("L" & L) = ProdDia * Prod                                                                         'DURAÇÃO
                        Duracao
                        LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                        P10.Range("M" & L) = LstProd(A, 9)                                                                          'Termino
                        LstProd(A, 8) = LstProd(A, 9)                                                                               'Novo Inicio
                        N = N + 1
                        P = P + ProdDia
                        L = L + 1
                    Else
                        If ProdDia = 0 Then P10.Range("I" & L) = 1 Else P10.Range("I" & L) = ProdDia                                'Qt
                        P10.Range("J" & L) = LstProd(A, 1)                                                                          'Qt programada
                        P10.Range("K" & L) = LstProd(A, 8)                                                                          'Inicio
                        If ProdDia = 0 Then P10.Range("L" & L) = 1 * Prod Else P10.Range("L" & L) = ProdDia * Prod                  'DURAÇÃO
                        Duracao
                        LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                        P10.Range("M" & L) = LstProd(A, 9)                                                                          'Termino
                        LstProd(A, 8) = LstProd(A, 9)                                                                               'Novo Inicio
                        N = N + 1
                        P = P + ProdDia
                        L = L + 1
                    End If
                End If
            End If
        Wend
        End If
    Else    '========================================================================================================================== Fazer Pergunta se Pelo SPS ou Contagem
        Resp = MsgBox("Informar como Deseja Realizar a Programação da Produção" _
                    & Chr(10) & "Sim - Programar pelo SPS" _
                    & Chr(10) & "Não - Programar pela Contagem de Peças Brutas", vbYesNo, "Programação da Produção")
        If Resp = 6 Then
            RelatorioProducao
            Classificar
        Else
            RelContagem
        End If
        Lcapa = 11
        While P00.Range("B" & Lcapa) <> ""
            If P00.Range("D" & Lcapa) = "Sim" Then  '=================================================================================> Sequenciar esta máquina
                Mesa(0, 0) = P00.Range("J9")
                Mesa(1, 0) = P00.Range("J9")
                Cod = P00.Range("B" & Lcapa) & "-" & P00.Range("C" & Lcapa)
                Cont = WorksheetFunction.CountIf(P08.Range("E:E"), Cod)
                Erase LstProd
                If Cont = 0 Then
                    GoTo outroCodigo
                End If
                Lproc = P08.Cells.Find(What:=Cod, After:=ActiveCell, LookIn:=xlFormulas, _
                            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                            MatchCase:=False, SearchFormat:=False).Row
                For A = 0 To Cont - 1
                    LstProd(A, 0) = P08.Range("A" & Lproc)                                                                              'modelo
                    LstProd(A, 1) = P08.Range("B" & Lproc)                                                                              'QT
                    LstProd(A, 2) = P08.Range("C" & Lproc)                                                                              'PESO
                    LstProd(A, 3) = P08.Range("D" & Lproc)                                                                              'QT MESA
                    LstProd(A, 4) = P08.Range("E" & Lproc)                                                                              'MAQUINA
                    LstProd(A, 5) = P08.Range("F" & Lproc)                                                                              'CLIENTE
                    LstProd(A, 6) = P08.Range("G" & Lproc)                                                                              'TEMPO MIN
                    LstProd(A, 7) = P08.Range("H" & Lproc)                                                                              'SETUP
                    LstProd(A, 8) = P08.Range("I" & Lproc)                                                                              'HORA INICIO
                    LstProd(A, 9) = P08.Range("J" & Lproc)                                                                              'HORA FIM
                    LstProd(A, 10) = P08.Range("K" & Lproc)                                                                             'LOTE MINIMO
                    LstProd(A, 11) = P08.Range("L" & Lproc)                                                                             'PEÇA IRMÃ
                    LstProd(A, 12) = P08.Range("M" & Lproc)                                                                             'DT CARTEIRA
                    LstProd(A, 13) = P08.Range("N" & Lproc)                                                                             'DT DEPOSITO
                    LstProd(A, 14) = P08.Range("O" & Lproc)                                                                             'OBSERVAÇAO
                    LstProd(A, 15) = P08.Range("P" & Lproc)                                                                             'STATUS
                    LstProd(A, 16) = P08.Range("Q" & Lproc)                                                                             'APROVAR
                    LstProd(A, 17) = P08.Range("Q" & Lproc).Row                                                                         'ENDEREÇO
                    Lproc = P08.Cells.FindNext(After:=P08.Range("Q" & Lproc)).Row
                Next A
                OrganizarSeguencia
                '------------------------------     PONTO DE VERIFICAÇÃO
                'VerificarParaAnalisar
                'Stop
                'P10.Select
                'P10.Range("A1048576").End(xlUp).Offset(1, 0).Select
                '-----------------------------
                A = 0
                If LstProd(A, 0) = LstProd(A + 2, 0) And LstProd(A, 14) = "" And LstProd(A + 2, 14) = "" Then
                    Dim troca(30) As Variant
                    Dim Tr As Integer
                    For Tr = 0 To 21        '==========================> Colocar no temporário
                        troca(Tr) = LstProd(A + 1, Tr)
                    Next Tr
                    
                    For Tr = 0 To 21        '==========================> Passa Terceira linha para Segunda
                        LstProd(A + 1, Tr) = LstProd(A + 2, Tr)
                    Next Tr
                    
                    For Tr = 0 To 21        '==========================> Passa temporária para terceira linha
                        LstProd(A + 2, Tr) = troca(Tr)
                    Next Tr
                '------------------------------
                'VerificarParaAnalisar
                'Stop
                'P10.Select
                'P10.Range("A1048576").End(xlUp).Offset(1, 0).Select
                '-----------------------------
                End If
                For C = 0 To Cont - 1
                    A = C
                    If LstProd(A, 3) = 1 Then       '===============================================================================>   UMA MESA
                        If LstProd(A + 1, 11) = "-" Then '==========================================================================>   PRÓXIMA NÃO É IRMÃ CONSIGO SEQUENCIAR AS DUAS
                            If LstProd(A, 14) = "" Then
                                P = 0
                                N = 0
                                L = P10.Range("A1048576").End(xlUp).Offset(1, 0).Row
                                For B = 0 To 1
                                    If LstProd(A, 14) = "" Then
                                        While P < LstProd(A, 1)
                                            If N = 0 Then       '=======================================================================>   SETUP DE MÁQUINA
                                                If Mesa(0, 0) > Mesa(1, 0) Then
                                                    LstProd(A, 8) = Mesa(1, 0)
                                                    Mesa(0, 1) = LstProd(A, 0)
                                                    Pego = 1
                                                Else
                                                    LstProd(A, 8) = Mesa(0, 0)
                                                    Mesa(0, 1) = LstProd(A, 0)
                                                    Pego = 0
                                                End If
                                                LstProd(A, 18) = LstProd(A, 8)
                                                P10.Range("A" & L) = LstProd(A, 4)                                                          'Máquina
                                                P10.Range("B" & L) = LstProd(A, 0)                                                          'Modelo
                                                P10.Range("C" & L) = LstProd(A, 5)                                                          'Cliente
                                                P10.Range("D" & L) = LstProd(A, 12)                                                         'Data Carteira
                                                P10.Range("E" & L) = LstProd(A, 13)                                                         'Data Depósito
                                                P10.Range("F" & L) = "Setup de Máquina"                                                     'Descrição
                                                If B = 0 Then
                                                    P10.Range("G" & L) = LstProd(A, 3) & "A"                                                'Qt mesa
                                                Else
                                                    P10.Range("G" & L) = LstProd(A, 3) & "B"                                                'Qt mesa
                                                End If
                                                P10.Range("H" & L) = LstProd(A, 2)                                                          'Peso
                                                P10.Range("I" & L) = 0                                                                      'Qt
                                                P10.Range("J" & L) = LstProd(A, 1)                                                          'Qt programada
                                                P10.Range("K" & L) = LstProd(A, 8)                                                          'Inicio
                                                P10.Range("L" & L) = LstProd(A, 7) * #12:01:00 AM# + LstProd(A, 6) * #12:01:00 AM#                                      'DURAÇÃO
                                                Duracao
                                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao                          'calculo para termino do trabalho
                                                P10.Range("M" & L) = LstProd(A, 9)                                                          'Termino
                                                If Mesa(0, 0) = Mesa(1, 0) Then
                                                    Mesa(0, 0) = LstProd(A, 9)
                                                    Mesa(1, 0) = LstProd(A, 9)
                                                End If
                                                LstProd(A, 8) = LstProd(A, 9)                                                               'Novo Inicio
                                                Mesa(Pego, 0) = LstProd(A, 9)
                                                N = N + 1
                                                L = L + 1
                                                LstProd(A, 19) = LstProd(A, 9)
                                            Else                '=======================================================================>   SEQUENCIAR PRODUÇÃO
                                                P10.Range("A" & L) = LstProd(A, 4)                                                          'Máquina
                                                P10.Range("B" & L) = LstProd(A, 0)                                                          'Modelo
                                                P10.Range("C" & L) = LstProd(A, 5)                                                          'Cliente
                                                P10.Range("D" & L) = LstProd(A, 12)                                                         'Data Carteira
                                                P10.Range("E" & L) = LstProd(A, 13)                                                         'Data Depósito
                                                If Mesa(Pego, 2) <> "" Then
                                                    P10.Range("F" & L) = "Produção Dia " & Format(N - 1, "00") & " Próximo Item"            'Descrição
                                                    N = N - 1
                                                    Mesa(Pego, 2) = Null
                                                Else
                                                    P10.Range("F" & L) = "Produção Dia " & Format(N, "00")                                  'Descrição
                                                End If
                                                If B = 0 Then
                                                    P10.Range("G" & L) = LstProd(A, 3) & "A"                                                'Qt mesa
                                                Else
                                                    P10.Range("G" & L) = LstProd(A, 3) & "B"                                                'Qt mesa
                                                End If
                                                P10.Range("H" & L) = LstProd(A, 2)                                                          'Peso
                                                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                                CalcularProducao
                                                '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                                If P00.Range("H16") = Dt And P00.Range("J11") = "Sim" Then
                                                    ProdDia = WorksheetFunction.RoundUp(0.5625 / Prod, 0)
                                                    If (ProdDia + P) > LstProd(A, 1) Then
                                                        ProdDia = LstProd(A, 1) - P
                                                        P10.Range("I" & L) = ProdDia                                                    'Qt
                                                        P10.Range("J" & L) = LstProd(A, 1)                                              'Qt programada
                                                        P10.Range("K" & L) = LstProd(A, 8)                                              'Inicio
                                                        P10.Range("L" & L) = ProdDia * Prod                                             'DURAÇÃO
                                                        Duracao
                                                        LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                        P10.Range("M" & L) = LstProd(A, 9)                                              'Termino
                                                        LstProd(A, 8) = LstProd(A, 9)                                                   'Novo Inicio
                                                        N = N + 1
                                                        P = P + ProdDia
                                                        L = L + 1
                                                        LstProd(A, 19) = LstProd(A, 9)
                                                        Mesa(Pego, 0) = LstProd(A, 9)
                                                    Else
                                                        P10.Range("I" & L) = ProdDia                                                    'Qt
                                                        P10.Range("J" & L) = LstProd(A, 1)                                              'Qt programada
                                                        P10.Range("K" & L) = LstProd(A, 8)                                              'Inicio
                                                        P10.Range("L" & L) = ProdDia * Prod                                             'DURAÇÃO
                                                        Duracao
                                                        LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                        P10.Range("M" & L) = LstProd(A, 9)                                              'Termino
                                                        LstProd(A, 8) = WorksheetFunction.RoundUp(LstProd(A, 9), 0) + 0.9166            'Novo Inicio
                                                        N = N + 1
                                                        P = P + ProdDia
                                                        L = L + 1
                                                        LstProd(A, 19) = LstProd(A, 9)
                                                        Mesa(Pego, 0) = LstProd(A, 9)
                                                    End If
                                                Else
                                                    ProdDia = WorksheetFunction.RoundUp(CapProd / Prod, 0)
                                                    If (ProdDia + P) > LstProd(A, 1) Then
                                                        ProdDia = LstProd(A, 1) - P
                                                        P10.Range("I" & L) = ProdDia                                                    'Qt
                                                        P10.Range("J" & L) = LstProd(A, 1)                                              'Qt programada
                                                        P10.Range("K" & L) = LstProd(A, 8)                                              'Inicio
                                                        P10.Range("L" & L) = ProdDia * Prod                                             'DURAÇÃO
                                                        Duracao
                                                        LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                        P10.Range("M" & L) = LstProd(A, 9)                                              'Termino
                                                        LstProd(A, 8) = LstProd(A, 9)                                                   'Novo Inicio
                                                        Mesa(Pego, 0) = LstProd(A, 9)
                                                        N = N + 1
                                                        P = P + ProdDia
                                                        L = L + 1
                                                        LstProd(A, 19) = LstProd(A, 9)
                                                        Mesa(Pego, 0) = LstProd(A, 9)
                                                        If P = LstProd(A, 1) Then   'verificar se próximo item é da mesma produção
                                                            If LstProd(A, 0) = LstProd(A + 1, 0) And LstProd(A + 1, 14) = "" Then
                                                                P = 0
                                                                Mesa(Pego, 2) = LstProd(A, 0)
                                                                LstProd(A + 1, 8) = LstProd(A, 9)                                       'Novo Inicio
                                                                A = A + 1
                                                            End If
                                                        End If
                                                    Else
                                                        If ProdDia = 0 Then P10.Range("I" & L) = 1 Else P10.Range("I" & L) = ProdDia    'Qt
                                                        P10.Range("J" & L) = LstProd(A, 1)                                              'Qt programada
                                                        P10.Range("K" & L) = LstProd(A, 8)                                              'Inicio
                                                        If ProdDia = 0 Then P10.Range("L" & L) = 1 * Prod _
                                                        Else P10.Range("L" & L) = ProdDia * Prod                                        'DURAÇÃO
                                                        '-----------------------------------------------------------
                                                        Duracao
                                                        LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                        P10.Range("M" & L) = LstProd(A, 9)                                              'Termino
                                                        LstProd(A, 8) = LstProd(A, 9)                                                   'Novo Inicio
                                                        N = N + 1
                                                        P = P + ProdDia
                                                        L = L + 1
                                                        C = A
                                                        LstProd(A, 19) = LstProd(A, 9)
                                                        Mesa(Pego, 0) = LstProd(A, 9)
                                                        If P = LstProd(A, 1) Then   'verificar se próximo item é da mesma produção
                                                            If LstProd(A, 0) = LstProd(A + 1, 0) Then
                                                                P = 0
                                                                Mesa(Pego, 2) = LstProd(A, 0)
                                                                LstProd(A + 1, 8) = LstProd(A, 9)                                       'Novo Inicio
                                                                A = A + 1
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Wend
                                        A = A + 1
                                        N = 0
                                        P = 0
                                    Else
                                        Lprod = P09.Range("A1048576").End(xlUp).Offset(1, 0).Row
                                        P09.Range("A" & Lprod) = LstProd(A, 4)
                                        P09.Range("C" & Lprod) = "Falta modelo para segunda Mesa"
                                        C = A
                                    End If
                                Next B
                                C = A - 1
                            Else
                                Lproc = LstProd(A, 17)
                                P08.Range("P" & Lproc) = "Não Programado"
                            End If
                        Enc = False
                        Else    '===================================================================================================>   PRÓXIMA É IRMÃ
                            If LstProd(A, 14) = "" Then
                            '=======================================================================================================>   SEQUENCIAR PRIMEIRO MODELO
                            P = 0
                            N = 0
                            L = P10.Range("A1048576").End(xlUp).Offset(1, 0).Row
                            While P < LstProd(A, 1)
                                If N = 0 Then       '===============================================================================>   SETUP DE MÁQUINA
                                    If Mesa(0, 0) > Mesa(1, 0) Then
                                        LstProd(A, 8) = Mesa(1, 0)
                                        Pego = 1
                                    Else
                                        LstProd(A, 8) = Mesa(0, 0)
                                        Pego = 0
                                    End If
                                    LstProd(A, 18) = LstProd(A, 8)
                                    P10.Range("A" & L) = LstProd(A, 4)                                                                  'Máquina
                                    P10.Range("B" & L) = LstProd(A, 0)                                                                  'Modelo
                                    P10.Range("C" & L) = LstProd(A, 5)                                                                  'Cliente
                                    P10.Range("D" & L) = LstProd(A, 12)                                                                 'Data Carteira
                                    P10.Range("E" & L) = LstProd(A, 13)                                                                 'Data Depósito
                                    P10.Range("F" & L) = "Setup de Máquina"                                                             'Descrição
                                    P10.Range("G" & L) = LstProd(A, 3)                                                                  'Qt mesa
                                    P10.Range("H" & L) = LstProd(A, 2)                                                                  'Peso
                                    P10.Range("I" & L) = 0                                                                              'Qt
                                    P10.Range("J" & L) = LstProd(A, 1)                                                                  'Qt programada
                                    P10.Range("K" & L) = LstProd(A, 8)                                                                  'Inicio
                                    P10.Range("L" & L) = LstProd(A, 7) * #12:01:00 AM#                                                  'DURAÇÃO
                                    Duracao
                                    LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao                                  'calculo para termino do trabalho
                                    P10.Range("M" & L) = LstProd(A, 9)                                                                  'Termino
                                    LstProd(A, 8) = LstProd(A, 9)                                                                       'Novo Inicio
                                    Mesa(Pego, 0) = LstProd(A, 9)                                                                         'Inicio Segunda mesa
                                    N = N + 1
                                    L = L + 1
                                Else    '==========================================================================================>    SEQUENCIAR PRODUÇÃO
                                    P10.Range("A" & L) = LstProd(A, 4)                                                                  'Máquina
                                    P10.Range("B" & L) = LstProd(A, 0)                                                                  'Modelo
                                    P10.Range("C" & L) = LstProd(A, 5)                                                                  'Cliente
                                    P10.Range("D" & L) = LstProd(A, 12)                                                                 'Data Carteira
                                    P10.Range("E" & L) = LstProd(A, 13)                                                                 'Data Depósito
                                    P10.Range("F" & L) = "Produção Dia " & Format(N, "00")                                              'Descrição
                                    P10.Range("G" & L) = LstProd(A, 3)                                                                  'Qt mesa
                                    P10.Range("H" & L) = LstProd(A, 2)                                                                  'Peso
                                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                    CalcularProducao
                                    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                    If P00.Range("H16") = Dt And P00.Range("J11") = "Sim" Then
                                        ProdDia = WorksheetFunction.RoundUp(0.5625 / Prod, 0)
                                        If (ProdDia + P) > LstProd(A, 1) Then
                                            ProdDia = LstProd(A, 1) - P
                                            P10.Range("I" & L) = ProdDia                                                            'Qt
                                            P10.Range("J" & L) = LstProd(A, 1)                                                      'Qt programada
                                            P10.Range("K" & L) = LstProd(A, 8)                                                      'Inicio
                                            P10.Range("L" & L) = ProdDia * Prod                                                     'DURAÇÃO
                                            Duracao
                                            LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                            P10.Range("M" & L) = LstProd(A, 9)                                                      'Termino
                                            LstProd(A, 8) = LstProd(A, 9)                                                           'Novo Inicio
                                            N = N + 1
                                            P = P + ProdDia
                                            L = L + 1
                                            LstProd(A, 19) = LstProd(A, 9)
                                            Mesa(Pego, 0) = LstProd(A, 9)
                                        Else
                                            P10.Range("I" & L) = ProdDia                                                            'Qt
                                            P10.Range("J" & L) = LstProd(A, 1)                                                      'Qt programada
                                            P10.Range("K" & L) = LstProd(A, 8)                                                      'Inicio
                                            P10.Range("L" & L) = ProdDia * Prod                                                     'DURAÇÃO
                                            Duracao
                                            LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                            P10.Range("M" & L) = LstProd(A, 9)                                                      'Termino
                                            LstProd(A, 8) = WorksheetFunction.RoundUp(LstProd(A, 9), 0) + 0.9167                    'Novo Inicio
                                            N = N + 1
                                            P = P + ProdDia
                                            L = L + 1
                                            LstProd(A, 19) = LstProd(A, 9)
                                        End If
                                    Else
                                        ProdDia = WorksheetFunction.RoundUp(CapProd / Prod, 0)
                                        If (ProdDia + P) > LstProd(A, 1) Then
                                            ProdDia = LstProd(A, 1) - P
                                            P10.Range("I" & L) = ProdDia                                                            'Qt
                                            P10.Range("J" & L) = LstProd(A, 1)                                                      'Qt programada
                                            P10.Range("K" & L) = LstProd(A, 8)                                                      'Inicio
                                            P10.Range("L" & L) = ProdDia * Prod                                                     'DURAÇÃO
                                            Duracao
                                            LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                            P10.Range("M" & L) = LstProd(A, 9)                                                      'Termino
                                            LstProd(A, 8) = LstProd(A, 9)                                                           'Novo Inicio
                                            Mesa(Pego, 0) = LstProd(A, 9)
                                            N = N + 1
                                            P = P + ProdDia
                                            L = L + 1
                                            LstProd(A, 19) = LstProd(A, 9)
                                        Else
                                            If ProdDia = 0 Then P10.Range("I" & L) = 1 _
                                            Else P10.Range("I" & L) = ProdDia                                                       'Qt
                                            P10.Range("J" & L) = LstProd(A, 1)                                                      'Qt programada
                                            P10.Range("K" & L) = LstProd(A, 8)                                                      'Inicio
                                            If ProdDia = 0 Then P10.Range("L" & L) = 1 * Prod _
                                            Else P10.Range("L" & L) = ProdDia * Prod                                                'DURAÇÃO
                                            Duracao
                                            LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                            P10.Range("M" & L) = LstProd(A, 9)                                                      'Termino
                                            LstProd(A, 8) = LstProd(A, 9)                                                           'Novo Inicio
                                            N = N + 1
                                            P = P + ProdDia
                                            L = L + 1
                                            LstProd(A, 19) = LstProd(A, 9)
                                            Mesa(Pego, 0) = LstProd(A, 9)
                                        End If
                                    End If
                                End If
                            Wend
                            A = A + 1
                            P = 0
                            N = 0
                            For B = 0 To 1
                                L = P10.Range("A1048576").End(xlUp).Offset(1, 0).Row
                                While P < LstProd(A, 1)
                                    If N = 0 Then       '==========================================================================>    SETUP DE MÁQUINA
                                        If Mesa(0, 0) > Mesa(1, 0) Then
                                            LstProd(A, 8) = Mesa(1, 0)
                                            Pego = 1
                                        Else
                                            LstProd(A, 8) = Mesa(0, 0)
                                            Pego = 0
                                        End If
                                        LstProd(A, 18) = LstProd(A, 8)
                                        P10.Range("A" & L) = LstProd(A, 4)                                                              'Máquina
                                        P10.Range("B" & L) = LstProd(A, 0)                                                              'Modelo
                                        P10.Range("C" & L) = LstProd(A, 5)                                                              'Cliente
                                        P10.Range("D" & L) = LstProd(A, 12)                                                             'Data Carteira
                                        P10.Range("E" & L) = LstProd(A, 13)                                                             'Data Depósito
                                        P10.Range("F" & L) = "Setup de Máquina"                                                         'Descrição
                                        P10.Range("G" & L) = LstProd(A, 3)                                                              'Qt mesa
                                        P10.Range("H" & L) = LstProd(A, 2)                                                              'Peso
                                        P10.Range("I" & L) = 0                                                                          'Qt
                                        P10.Range("J" & L) = LstProd(A, 1)                                                              'Qt programada
                                        P10.Range("K" & L) = LstProd(A, 8)                                                              'Inicio
                                        P10.Range("L" & L) = LstProd(A, 7) * #12:01:00 AM#                                              'DURAÇÃO
                                        Duracao
                                        LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao                              'calculo para termino do trabalho
                                        P10.Range("M" & L) = LstProd(A, 9)                                                              'Termino
                                        LstProd(A, 8) = LstProd(A, 9)                                                                   'Novo Inicio
                                        Mesa(Pego, 0) = LstProd(A, 9)
                                        N = N + 1
                                        L = L + 1
                                    Else    '======================================================================================>    SEQUENCIAR PRODUÇÃO
                                        P10.Range("A" & L) = LstProd(A, 4)                                                              'Máquina
                                        P10.Range("B" & L) = LstProd(A, 0)                                                              'Modelo
                                        P10.Range("C" & L) = LstProd(A, 5)                                                              'Cliente
                                        P10.Range("D" & L) = LstProd(A, 12)                                                             'Data Carteira
                                        P10.Range("E" & L) = LstProd(A, 13)                                                             'Data Depósito
                                        P10.Range("F" & L) = "Produção Dia " & Format(N, "00")                                          'Descrição
                                        P10.Range("G" & L) = LstProd(A, 3)                                                              'Qt mesa
                                        P10.Range("H" & L) = LstProd(A, 2)                                                              'Peso
                                        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                        CalcularProducao
                                        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                        If P00.Range("H16") = Dt And P00.Range("J11") = "Sim" Then
                                            ProdDia = WorksheetFunction.RoundUp(0.5625 / Prod, 0)
                                            If (ProdDia + P) > LstProd(A, 1) Then
                                                ProdDia = LstProd(A, 1) - P
                                                P10.Range("I" & L) = ProdDia                                                        'Qt
                                                P10.Range("J" & L) = LstProd(A, 1)                                                  'Qt programada
                                                P10.Range("K" & L) = LstProd(A, 8)                                                  'Inicio
                                                P10.Range("L" & L) = ProdDia * Prod                                                 'DURAÇÃO
                                                Duracao
                                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                P10.Range("M" & L) = LstProd(A, 9)                                                  'Termino
                                                LstProd(A, 8) = LstProd(A, 9)                                                       'Novo Inicio
                                                N = N + 1
                                                P = P + ProdDia
                                                L = L + 1
                                                LstProd(A, 19) = LstProd(A, 9)
                                            Else
                                                P10.Range("I" & L) = ProdDia                                                        'Qt
                                                P10.Range("J" & L) = LstProd(A, 1)                                                  'Qt programada
                                                P10.Range("K" & L) = LstProd(A, 8)                                                  'Inicio
                                                P10.Range("L" & L) = ProdDia * Prod                                                 'DURAÇÃO
                                                Duracao
                                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                P10.Range("M" & L) = LstProd(A, 9)                                                  'Termino
                                                LstProd(A, 8) = WorksheetFunction.RoundUp(LstProd(A, 9), 0) + 0.9166                'Novo Inicio
                                                N = N + 1
                                                P = P + ProdDia
                                                L = L + 1
                                                LstProd(A, 19) = LstProd(A, 9)
                                                Mesa(Pego, 0) = LstProd(A, 9)
                                            End If
                                        Else
                                            ProdDia = WorksheetFunction.RoundUp(CapProd / Prod, 0)
                                            If (ProdDia + P) > LstProd(A, 1) Then
                                                ProdDia = LstProd(A, 1) - P
                                                P10.Range("I" & L) = ProdDia                                                        'Qt
                                                P10.Range("J" & L) = LstProd(A, 1)                                                  'Qt programada
                                                P10.Range("K" & L) = LstProd(A, 8)                                                  'Inicio
                                                P10.Range("L" & L) = ProdDia * Prod                                                 'DURAÇÃO
                                                Duracao
                                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                P10.Range("M" & L) = LstProd(A, 9)                                                  'Termino
                                                LstProd(A, 8) = LstProd(A, 9)                                                       'Novo Inicio
                                                N = N + 1
                                                P = P + ProdDia
                                                L = L + 1
                                                LstProd(A, 19) = LstProd(A, 9)
                                                Mesa(Pego, 0) = LstProd(A, 9)
                                            Else
                                                If ProdDia = 0 Then P10.Range("I" & L) = 1 Else P10.Range("I" & L) = ProdDia        'Qt
                                                P10.Range("J" & L) = LstProd(A, 1)                                                  'Qt programada
                                                P10.Range("K" & L) = LstProd(A, 8)                                                  'Inicio
                                                If ProdDia = 0 Then P10.Range("L" & L) = 1 * Prod _
                                                Else P10.Range("L" & L) = ProdDia * Prod                                            'DURAÇÃO
                                                Duracao
                                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                P10.Range("M" & L) = LstProd(A, 9)                                                  'Termino
                                                LstProd(A, 8) = LstProd(A, 9)                                                       'Novo Inicio
                                                N = N + 1
                                                P = P + ProdDia
                                                L = L + 1
                                                LstProd(A, 19) = LstProd(A, 9)
                                                Mesa(Pego, 0) = LstProd(A, 9)
                                            End If
                                        End If
                                    End If
                                Wend
                            Next B
                            Enc = False
                            C = A
                            Else
                                Lproc = LstProd(A, 17)
                                P08.Range("P" & Lproc) = "Não Programado"
                            End If
                        End If
                    Else    '=======================================================================================================>   MAIS DE UMA MESA
                        If LstProd(A, 7) = "" Then '========================================>FALTA SETUP DE MAQUINA?
                            Lt = P06.Range("A1048576").End(xlUp).Offset(1, 0).Row
                            P06.Range("A" & Lt) = "Falta Setup"
                            P06.Range("B" & Lt) = LstProd(A, 0)
                            P06.Range("C" & Lt) = LstProd(A, 1)
                            Lproc = LstProd(A, 17)
                            P08.Range("O" & Lproc) = "Falta Setup da Máquina"
                            P08.Range("P" & Lproc) = "Não Programado"
                            P08.Range("I" & Lproc) = Null
                            P08.Range("J" & Lproc) = Null
                            LstProd(A, 14) = "Não Programado Falta Setup"
                        End If
                        If LstProd(A, 14) = "" Then     '=====================================> TEM ALGUMA OBSERVAÇÃO?
                            If LstProd(A, 11) = "-" Then        '=============================> NÃO TEM MODELO IRMÃO
                                P = 0
                                N = 0
                                L = P10.Range("A1048576").End(xlUp).Offset(1, 0).Row
                                While P < LstProd(A, 1)
                                    If N = 0 Then   '==============================================================================>    SETUP DE MÁQUINA
                                        If Mesa(0, 0) > Mesa(1, 0) Then
                                            LstProd(A, 8) = Mesa(1, 0)
                                            Pego = 1
                                        Else
                                            LstProd(A, 8) = Mesa(0, 0)
                                            Pego = 0
                                        End If
                                        LstProd(A, 18) = LstProd(A, 8)
                                        P10.Range("A" & L) = LstProd(A, 4)                                                              'Máquina
                                        P10.Range("B" & L) = LstProd(A, 0)                                                              'Modelo
                                        P10.Range("C" & L) = LstProd(A, 5)                                                              'Cliente
                                        P10.Range("D" & L) = LstProd(A, 12)                                                             'Data Carteira
                                        P10.Range("E" & L) = LstProd(A, 13)                                                             'Data Depósito
                                        P10.Range("F" & L) = "Setup de Máquina"                                                         'Descrição
                                        P10.Range("G" & L) = LstProd(A, 3)                                                              'Qt mesa
                                        P10.Range("H" & L) = LstProd(A, 2)                                                              'Peso
                                        P10.Range("I" & L) = 0                                                                          'Qt
                                        P10.Range("J" & L) = LstProd(A, 1)                                                              'Qt programada
                                        P10.Range("K" & L) = LstProd(A, 8)                                                              'Inicio
                                        P10.Range("L" & L) = LstProd(A, 7) * #12:01:00 AM#                                              'DURAÇÃO
                                        Duracao
                                        LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao                              'calculo para termino do trabalho
                                        P10.Range("M" & L) = LstProd(A, 9)                                                              'Termino
                                        LstProd(A, 8) = LstProd(A, 9)                                                                   'Novo Inicio
                                        
                                        If Mesa(0, 0) = Mesa(1, 0) Then
                                            Mesa(0, 0) = LstProd(A, 9)
                                            Mesa(1, 0) = LstProd(A, 9)
                                        End If
                                        Mesa(Pego, 0) = LstProd(A, 9)
                                        N = N + 1
                                        L = L + 1
                                    Else    '======================================================================================>    SEQUENCIAR PRODUÇÃO
                                        L = P10.Range("A1048576").End(xlUp).Offset(1, 0).Row
                                        P10.Range("A" & L) = LstProd(A, 4)                                                              'Máquina
                                        P10.Range("B" & L) = LstProd(A, 0)                                                              'Modelo
                                        P10.Range("C" & L) = LstProd(A, 5)                                                              'Cliente
                                        P10.Range("D" & L) = LstProd(A, 12)                                                             'Data Carteira
                                        P10.Range("E" & L) = LstProd(A, 13)                                                             'Data Depósito
                                        If LstProd(A, 18) = "Continua" Then
                                            N = N - 1
                                            P10.Range("F" & L) = "Produção Dia " & Format(N, "00 ") & "Próximo Item"                       'Descrição
                                            LstProd(A, 18) = Null
                                        Else
                                            P10.Range("F" & L) = "Produção Dia " & Format(N, "00 ")
                                        End If
                                        P10.Range("G" & L) = LstProd(A, 3)                                                              'Qt mesa
                                        P10.Range("H" & L) = LstProd(A, 2)                                                              'Peso
                                        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                        CalcularProducao
                                        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                        If P00.Range("H16") = Dt And P00.Range("J11") = "Sim" Then
                                            ProdDia = WorksheetFunction.RoundUp(0.5625 / Prod, 0)
                                            If (ProdDia + P) > LstProd(A, 1) Then
                                                ProdDia = LstProd(A, 1) - P
                                                P10.Range("I" & L) = ProdDia                                                        'Qt
                                                P10.Range("J" & L) = LstProd(A, 1)                                                  'Qt programada
                                                P10.Range("K" & L) = LstProd(A, 8)                                                  'Inicio
                                                P10.Range("L" & L) = ProdDia * Prod                                                 'DURAÇÃO
                                                Duracao
                                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                P10.Range("M" & L) = LstProd(A, 9)                                                  'Termino
                                                LstProd(A, 8) = LstProd(A, 9)                                                       'Novo Inicio
                                                N = N + 1
                                                P = P + ProdDia
                                                L = L + 1
                                                LstProd(A, 19) = LstProd(A, 9)
                                                Mesa(Pego, 0) = LstProd(A, 9)
                                            Else
                                                P10.Range("I" & L) = ProdDia                                                        'Qt
                                                P10.Range("J" & L) = LstProd(A, 1)                                                  'Qt programada
                                                P10.Range("K" & L) = LstProd(A, 8)                                                  'Inicio
                                                P10.Range("L" & L) = ProdDia * Prod                                                 'DURAÇÃO
                                                Duracao
                                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                P10.Range("M" & L) = LstProd(A, 9)                                                  'Termino
                                                LstProd(A, 8) = WorksheetFunction.RoundUp(LstProd(A, 9), 0) + 0.9166                'Novo Inicio
                                                N = N + 1
                                                P = P + ProdDia
                                                L = L + 1
                                                LstProd(A, 19) = LstProd(A, 9)
                                                Mesa(Pego, 0) = LstProd(A, 9)
                                            End If
                                        Else
                                            ProdDia = WorksheetFunction.RoundUp(CapProd / Prod, 0)
                                            If (ProdDia + P) > LstProd(A, 1) Then
                                                ProdDia = LstProd(A, 1) - P
                                                P10.Range("I" & L) = ProdDia                                                        'Qt
                                                P10.Range("J" & L) = LstProd(A, 1)                                                  'Qt programada
                                                P10.Range("K" & L) = LstProd(A, 8)                                                  'Inicio
                                                P10.Range("L" & L) = ProdDia * Prod                                                 'DURAÇÃO
                                                Duracao
                                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                P10.Range("M" & L) = LstProd(A, 9)                                                  'Termino
                                                LstProd(A, 8) = LstProd(A, 9)                                                       'Novo Inicio
                                                N = N + 1
                                                P = P + ProdDia
                                                L = L + 1
                                                LstProd(A, 19) = LstProd(A, 9)
                                                Mesa(Pego, 0) = LstProd(A, 9)
                                            Else
                                                If ProdDia = 0 Then P10.Range("I" & L) = 1 Else P10.Range("I" & L) = ProdDia        'Qt
                                                P10.Range("J" & L) = LstProd(A, 1)                                                  'Qt programada
                                                P10.Range("K" & L) = LstProd(A, 8)                                                  'Inicio
                                                If ProdDia = 0 Then P10.Range("L" & L) = 1 * Prod _
                                                Else P10.Range("L" & L) = ProdDia * Prod                                            'DURAÇÃO
                                                Duracao
                                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                P10.Range("M" & L) = LstProd(A, 9)                                                  'Termino
                                                LstProd(A, 8) = LstProd(A, 9)                                                       'Novo Inicio
                                                N = N + 1
                                                P = P + ProdDia
                                                L = L + 1
                                                LstProd(A, 19) = LstProd(A, 9)
                                                Mesa(Pego, 0) = LstProd(A, 9)
                                            End If
                                            If P = LstProd(A, 1) Then
                                                If LstProd(A, 0) = LstProd(A + 1, 0) Then
                                                    LstProd(A + 1, 18) = "Continua"
                                                    LstProd(A + 1, 8) = LstProd(A, 8)
                                                    A = A + 1
                                                    P = 0
                                                    C = A
                                                End If
                                            End If
                                        End If
                                    End If
                                Wend
                            Else    '==============================================================================================>    Modelo irmão
                                If LstProd(A, 14) = "" Then
                                P = 0
                                N = 0
                                L = P10.Range("A1048576").End(xlUp).Offset(1, 0).Row
                                While P < LstProd(A, 1)
                                    If N = 0 Then       '==========================================================================>    SETUP DE MÁQUINA
                                        If Mesa(0, 0) > Mesa(1, 0) Then
                                            LstProd(A, 8) = Mesa(1, 0)
                                            Pego = 1
                                        Else
                                            LstProd(A, 8) = Mesa(0, 0)
                                            Pego = 0
                                        End If
                                        LstProd(A, 18) = LstProd(A, 8)
                                        P10.Range("A" & L) = LstProd(A, 4)                                                              'Máquina
                                        P10.Range("B" & L) = LstProd(A, 0)                                                              'Modelo
                                        P10.Range("C" & L) = LstProd(A, 5)                                                              'Cliente
                                        P10.Range("D" & L) = LstProd(A, 12)                                                             'Data Carteira
                                        P10.Range("E" & L) = LstProd(A, 13)                                                             'Data Depósito
                                        P10.Range("F" & L) = "Setup de Máquina"                                                         'Descrição
                                        P10.Range("G" & L) = LstProd(A, 3)                                                              'Qt mesa
                                        P10.Range("H" & L) = LstProd(A, 2)                                                              'Peso
                                        P10.Range("I" & L) = 0                                                                          'Qt
                                        P10.Range("J" & L) = LstProd(A, 1)                                                              'Qt programada
                                        P10.Range("K" & L) = LstProd(A, 8)                                                              'Inicio
                                        P10.Range("L" & L) = LstProd(A, 7) * #12:01:00 AM#                                              'DURAÇÃO
                                        Duracao
                                        LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao                              'calculo para termino do trabalho
                                        P10.Range("M" & L) = LstProd(A, 9)                                                              'Termino
                                        LstProd(A, 8) = LstProd(A, 9)                                                                   'Novo Inicio
                                        Mesa(Pego, 0) = LstProd(A, 9)                                                                     'Inicio Segunda mesa
                                        N = N + 1
                                        L = L + 1
                                    Else    '======================================================================================>    SEQUENCIAR PRODUÇÃO
                                        P10.Range("A" & L) = LstProd(A, 4)                                                              'Máquina
                                        P10.Range("B" & L) = LstProd(A, 0)                                                              'Modelo
                                        P10.Range("C" & L) = LstProd(A, 5)                                                              'Cliente
                                        P10.Range("D" & L) = LstProd(A, 12)                                                             'Data Carteira
                                        P10.Range("E" & L) = LstProd(A, 13)                                                             'Data Depósito
                                        P10.Range("F" & L) = "Produção Dia " & Format(N, "00")                                          'Descrição
                                        P10.Range("G" & L) = LstProd(A, 3)                                                              'Qt mesa
                                        P10.Range("H" & L) = LstProd(A, 2)                                                              'Peso
                                        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                        CalcularProducao
                                        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                        If P00.Range("H16") = Dt And P00.Range("J11") = "Sim" Then
                                            ProdDia = WorksheetFunction.RoundUp(0.5625 / Prod, 0)
                                            If (ProdDia + P) > LstProd(A, 1) Then
                                                ProdDia = LstProd(A, 1) - P
                                                P10.Range("I" & L) = ProdDia                                                        'Qt
                                                P10.Range("J" & L) = LstProd(A, 1)                                                  'Qt programada
                                                P10.Range("K" & L) = LstProd(A, 8)                                                  'Inicio
                                                P10.Range("L" & L) = ProdDia * Prod                                                 'DURAÇÃO
                                                '-----------------------------------------------------------
                                                Duracao
                                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                P10.Range("M" & L) = LstProd(A, 9)                                                  'Termino
                                                LstProd(A, 8) = LstProd(A, 9)                                                       'Novo Inicio
                                                N = N + 1
                                                P = P + ProdDia
                                                L = L + 1
                                                LstProd(A, 19) = LstProd(A, 9)
                                                Mesa(Pego, 0) = LstProd(A, 9)
                                            Else
                                                P10.Range("I" & L) = ProdDia                                                        'Qt
                                                P10.Range("J" & L) = LstProd(A, 1)                                                  'Qt programada
                                                P10.Range("K" & L) = LstProd(A, 8)                                                  'Inicio
                                                P10.Range("L" & L) = ProdDia * Prod                                                 'DURAÇÃO
                                                Duracao
                                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                P10.Range("M" & L) = LstProd(A, 9)                                                  'Termino
                                                LstProd(A, 8) = WorksheetFunction.RoundUp(LstProd(A, 9), 0) + 0.9167                'Novo Inicio
                                                N = N + 1
                                                P = P + ProdDia
                                                L = L + 1
                                                LstProd(A, 19) = LstProd(A, 9)
                                                Mesa(Pego, 0) = LstProd(A, 9)
                                            End If
                                        Else
                                            ProdDia = WorksheetFunction.RoundUp(CapProd / Prod, 0)
                                            If (ProdDia + P) > LstProd(A, 1) Then
                                                ProdDia = LstProd(A, 1) - P
                                                P10.Range("I" & L) = ProdDia                                                        'Qt
                                                P10.Range("J" & L) = LstProd(A, 1)                                                  'Qt programada
                                                P10.Range("K" & L) = LstProd(A, 8)                                                  'Inicio
                                                P10.Range("L" & L) = ProdDia * Prod                                                 'DURAÇÃO
                                                Duracao
                                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                P10.Range("M" & L) = LstProd(A, 9)                                                  'Termino
                                                LstProd(A, 8) = LstProd(A, 9)                                                       'Novo Inicio
                                                Mesa(Pego, 0) = LstProd(A, 9)
                                                N = N + 1
                                                P = P + ProdDia
                                                L = L + 1
                                                LstProd(A, 19) = LstProd(A, 9)
                                                Mesa(Pego, 0) = LstProd(A, 9)
                                            Else
                                                If ProdDia = 0 Then P10.Range("I" & L) = 1 Else P10.Range("I" & L) = ProdDia        'Qt
                                                P10.Range("J" & L) = LstProd(A, 1)                                                  'Qt programada
                                                P10.Range("K" & L) = LstProd(A, 8)                                                  'Inicio
                                                If ProdDia = 0 Then P10.Range("L" & L) = 1 * Prod _
                                                Else P10.Range("L" & L) = ProdDia * Prod                                            'DURAÇÃO
                                                Duracao
                                                LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                P10.Range("M" & L) = LstProd(A, 9)                                                  'Termino
                                                LstProd(A, 8) = LstProd(A, 9)                                                       'Novo Inicio
                                                N = N + 1
                                                P = P + ProdDia
                                                L = L + 1
                                                LstProd(A, 19) = LstProd(A, 9)
                                                Mesa(Pego, 0) = LstProd(A, 9)
                                            End If
                                        End If
                                    End If
                                Wend
                                A = A + 1
                                P = 0
                                N = 0
                                For B = 0 To 1
                                    L = P10.Range("A1048576").End(xlUp).Offset(1, 0).Row
                                    While P < LstProd(A, 1)
                                        If N = 0 Then   '==========================================================================>    SETUP DE MÁQUINA
                                            If Mesa(0, 0) > Mesa(1, 0) Then
                                                LstProd(A, 8) = Mesa(1, 0)
                                                Pego = 1
                                            Else
                                                LstProd(A, 8) = Mesa(0, 0)
                                                Pego = 0
                                            End If
                                            LstProd(A, 18) = LstProd(A, 8)
                                            P10.Range("A" & L) = LstProd(A, 4)                                                          'Máquina
                                            P10.Range("B" & L) = LstProd(A, 0)                                                          'Modelo
                                            P10.Range("C" & L) = LstProd(A, 5)                                                          'Cliente
                                            P10.Range("D" & L) = LstProd(A, 12)                                                         'Data Carteira
                                            P10.Range("E" & L) = LstProd(A, 13)                                                         'Data Depósito
                                            P10.Range("F" & L) = "Setup de Máquina"                                                     'Descrição
                                            P10.Range("G" & L) = LstProd(A, 3)                                                          'Qt mesa
                                            P10.Range("H" & L) = LstProd(A, 2)                                                          'Peso
                                            P10.Range("I" & L) = 0                                                                      'Qt
                                            P10.Range("J" & L) = LstProd(A, 1)                                                          'Qt programada
                                            P10.Range("K" & L) = LstProd(A, 8)                                                          'Inicio
                                            P10.Range("L" & L) = LstProd(A, 7) * #12:01:00 AM#                                          'DURAÇÃO
                                            Duracao
                                            LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao                          'calculo para termino do trabalho
                                            P10.Range("M" & L) = LstProd(A, 9)                                                          'Termino
                                            LstProd(A, 8) = LstProd(A, 9)                                                               'Novo Inicio
                                            Mesa(Pego, 0) = LstProd(A, 9)
                                            N = N + 1
                                            L = L + 1
                                        Else    '==================================================================================>    SEQUENCIAR PRODUÇÃO
                                            P10.Range("A" & L) = LstProd(A, 4)                                                          'Máquina
                                            P10.Range("B" & L) = LstProd(A, 0)                                                          'Modelo
                                            P10.Range("C" & L) = LstProd(A, 5)                                                          'Cliente
                                            P10.Range("D" & L) = LstProd(A, 12)                                                         'Data Carteira
                                            P10.Range("E" & L) = LstProd(A, 13)                                                         'Data Depósito
                                            P10.Range("F" & L) = "Produção Dia " & Format(N, "00")                                      'Descrição
                                            P10.Range("G" & L) = LstProd(A, 3)                                                          'Qt mesa
                                            P10.Range("H" & L) = LstProd(A, 2)                                                          'Peso
                                            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                            CalcularProducao
                                            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                                            If P00.Range("H16") = Dt And P00.Range("J11") = "Sim" Then
                                                ProdDia = WorksheetFunction.RoundUp(0.5625 / Prod, 0)
                                                If (ProdDia + P) > LstProd(A, 1) Then
                                                    ProdDia = LstProd(A, 1) - P
                                                    P10.Range("I" & L) = ProdDia                                                    'Qt
                                                    P10.Range("J" & L) = LstProd(A, 1)                                              'Qt programada
                                                    P10.Range("K" & L) = LstProd(A, 8)                                              'Inicio
                                                    P10.Range("L" & L) = ProdDia * Prod                                             'DURAÇÃO
                                                    Duracao
                                                    LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                    P10.Range("M" & L) = LstProd(A, 9)                                              'Termino
                                                    LstProd(A, 8) = LstProd(A, 9)                                                   'Novo Inicio
                                                    N = N + 1
                                                    P = P + ProdDia
                                                    L = L + 1
                                                    LstProd(A, 19) = LstProd(A, 9)
                                                    Mesa(Pego, 0) = LstProd(A, 9)
                                                Else
                                                    P10.Range("I" & L) = ProdDia                                                    'Qt
                                                    P10.Range("J" & L) = LstProd(A, 1)                                              'Qt programada
                                                    P10.Range("K" & L) = LstProd(A, 8)                                              'Inicio
                                                    P10.Range("L" & L) = ProdDia * Prod                                             'DURAÇÃO
                                                    Duracao
                                                    LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                    P10.Range("M" & L) = LstProd(A, 9)                                              'Termino
                                                    LstProd(A, 8) = WorksheetFunction.RoundUp(LstProd(A, 9), 0) + 0.9166            'Novo Inicio
                                                    N = N + 1
                                                    P = P + ProdDia
                                                    L = L + 1
                                                    LstProd(A, 19) = LstProd(A, 9)
                                                    Mesa(Pego, 0) = LstProd(A, 9)
                                                End If
                                            Else
                                                ProdDia = WorksheetFunction.RoundUp(CapProd / Prod, 0)
                                                If (ProdDia + P) > LstProd(A, 1) Then
                                                    ProdDia = LstProd(A, 1) - P
                                                    P10.Range("I" & L) = ProdDia                                                    'Qt
                                                    P10.Range("J" & L) = LstProd(A, 1)                                              'Qt programada
                                                    P10.Range("K" & L) = LstProd(A, 8)                                              'Inicio
                                                    P10.Range("L" & L) = ProdDia * Prod                                             'DURAÇÃO
                                                    Duracao
                                                    LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                    P10.Range("M" & L) = LstProd(A, 9)                                              'Termino
                                                    LstProd(A, 8) = LstProd(A, 9)                                                   'Novo Inicio
                                                    N = N + 1
                                                    P = P + ProdDia
                                                    L = L + 1
                                                    LstProd(A, 19) = LstProd(A, 9)
                                                    Mesa(Pego, 0) = LstProd(A, 9)
                                                Else
                                                    If ProdDia = 0 Then P10.Range("I" & L) = 1 Else P10.Range("I" & L) = ProdDia    'Qt
                                                    P10.Range("J" & L) = LstProd(A, 1)                                              'Qt programada
                                                    P10.Range("K" & L) = LstProd(A, 8)                                              'Inicio
                                                    If ProdDia = 0 Then P10.Range("L" & L) = 1 * Prod _
                                                    Else P10.Range("L" & L) = ProdDia * Prod                                        'DURAÇÃO
                                                    Duracao
                                                    LstProd(A, 9) = P10.Range("K" & L) + P10.Range("L" & L) + Refeicao
                                                    P10.Range("M" & L) = LstProd(A, 9)                                              'Termino
                                                    LstProd(A, 8) = LstProd(A, 9)                                                   'Novo Inicio
                                                    N = N + 1
                                                    P = P + ProdDia
                                                    L = L + 1
                                                    LstProd(A, 19) = LstProd(A, 9)
                                                    Mesa(Pego, 0) = LstProd(A, 9)
                                                End If
                                            End If
                                        End If
                                    Wend
                                Next B
                                Enc = False
                                C = A
                                End If
                            End If
                        Else
                            Lproc = LstProd(A, 17)
                            P08.Range("P" & Lproc) = "Não Programado"
                        End If
                    End If
                Next C
            End If
            'LANÇAR AS INFORMAÇÕES SEQUENCIADAS
            For A = 0 To Cont - 1
                If LstProd(A, 14) = "" Then
                    Lt = LstProd(A, 17)
                    P08.Range("P" & Lt) = "Programado"
                    P08.Range("I" & Lt) = LstProd(A, 18)
                    P08.Range("J" & Lt) = LstProd(A, 19)
                Else
                    Lt = LstProd(A, 17)
                    P08.Range("P" & Lt) = "Não Programado"
                    P08.Range("I" & Lt) = ""
                    P08.Range("J" & Lt) = ""
                End If

            Next A
outroCodigo:
            Lcapa = Lcapa + 1
        Wend
    End If
Else
    Exit Sub    '===================================================================================================================>   Desistir do seguenciamento de produção

End If
P11.Range("C" & Llog) = Timer - Tempo   '--------------------------------------------------------------------------------------------   LOG DE FIM DE OPERAÇÃO
ActiveWorkbook.Save
Sheets("Capa").Select
RelataAtraso
RelataAtrasoSeq
MsgBox "Sequenciamento Concluída com Sucesso"
Exit Sub
Erro:
MsgBox "Ocorreu um erro durante o sequenciamento da Produção, favor avisar programador", vbCritical + vbOKOnly, "Sequenciamento da Produção"
End Sub
Private Sub RelataAtraso()
    L = 5
    While P08.Range("A" & L) <> ""
        If P08.Range("P" & L) = "" Then
            P08.Range("P" & L) = "Não Programado"
            P08.Range("I" & L & ":J" & L) = ""
        End If
        If P08.Range("N" & L) <> "" Then
            If P08.Range("N" & L) <= P08.Range("J" & L) Then
                P08.Range("A" & L & ":R" & L).Font.Color = vbRed
                P08.Range("A" & L & ":R" & L).Font.Bold = True
                P08.Range("A" & L & ":R" & L).Font.Size = 12
                P08.Range("S" & L) = "AT"
            Else
                P08.Range("A" & L & ":R" & L).Font.Color = vbBlack
                P08.Range("A" & L & ":R" & L).Font.Bold = False
                P08.Range("A" & L & ":R" & L).Font.Size = 12
                P08.Range("S" & L) = "OK"
            End If
        Else
            P08.Range("A" & L & ":R" & L).Font.Color = vbBlack
            P08.Range("A" & L & ":R" & L).Font.Bold = False
            P08.Range("A" & L & ":R" & L).Font.Size = 12
        End If
        L = L + 1
    Wend
End Sub