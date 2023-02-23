Option Explicit
Public Vdt1 As Integer
Public Vdt2 As Integer

'Módulo Programar Produção

Sub Main()

'On Error GoTo Erro
'-------------------------  Perguntar ----------------------------------------------
    Resp = MsgBox("Realizar Relatório Produção", vbQuestion + vbYesNo, "Relatório Produção")

    If Resp = 6 Then
        MsgBox "Informar Local do Apontamento de Produção", vbInformation, "Programação de Produção"
        PrimeiraFase
'-------------------------  LOG DE INICIO DE OPERAÇÃO   ------------------------------------
        Llog = P11.Range("A1048576").End(xlUp).Offset(1, 0).Row
        P11.Range("A" & Llog) = "Relatório de Produção"
        P11.Range("B" & Llog) = Date
        Tempo = Timer
        Application.DisplayAlerts = False
        Home = ActiveWorkbook.Name
        LimpezaInicial
'-------------------------  Bucar a informação ----------------------------------------------
        MsgBox "Escolha o arquivo Inicial da Programação do PCP", vbInformation, "Programação"
        Job = Application.GetOpenFilename(Title:="Escolha o arquivo Inicial da Programação do PCP")
        If Job = "Falso" Then
            Exit Sub
        End If
            Workbooks.Open Filename:=Job
            Job = ActiveWorkbook.Name
            A = Sheets("MPS-UPR").Range("A1048576").End(xlUp).Row
            B = Sheets("MPS-UPR").Range("XFD2").End(xlToLeft).Column
            Sheets("MPS-UPR").Range(Cells(2, 1), Cells(A, B)).Copy
            Windows(Home).Activate
            P05.Select
            P05.Range("A1").PasteSpecial xlValues
            P05.Cells.EntireRow.AutoFit
            Windows(Job).Activate
            Application.CutCopyMode = False
            ActiveWorkbook.Close
            Windows(Home).Activate
'-------------------------  Primeira Seleção  -------------------------
            PrimeiraSelecao
            RelatorioProducao
    ElseIf Resp = 7 Then
        Exit Sub
    End If
'-------------------------  LOG DE FIM DE OPERAÇÃO   ------------------------------------
        P11.Range("C" & Llog) = Timer - Tempo
        ActiveWorkbook.Save
        Sheets("Capa").Select
        MsgBox "Primeira Etapa Concluída com Sucesso"
        Exit Sub
Erro:
    MsgBox "Ocorreu um erro durante o processamento da Informação, Favor avisar Programador", _
            vbCritical + vbOKOnly, "Relatório de Produção"
End Sub