Private Sub Classificar()
    Dim Tabela As Range

    Lt = P08.Range("A1048576").End(xlUp).Row
    
    Set Tabela = P08.Range("A4").CurrentRegion
    
    P08.Sort.SortFields.Clear

    P08.Sort.SortFields.Add Tabela.Columns(14), xlSortOnValues, xlAscending
    P08.Sort.SortFields.Add Tabela.Columns(5), xlSortOnValues, xlAscending
    P08.Sort.SetRange Tabela
    P08.Sort.Header = xlYes
    P08.Sort.Apply
    
End Sub