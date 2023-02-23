Private Sub RelataNaoAtendeLoteMinimo()
    Lprod = P09.Range("a1048576").End(xlUp).Offset(1, 0).Row
    P09.Range("A" & Lprod) = Maq(0, 0)
    P09.Range("B" & Lprod) = Maq(0, 1)
    P09.Range("C" & Lprod) = "Não atende lote mínimo para Programação de produção"
End Sub