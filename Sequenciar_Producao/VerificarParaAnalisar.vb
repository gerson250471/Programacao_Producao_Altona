Private Sub VerificarParaAnalisar()
    P17.Select
    P17.AutoFilterMode = False
    P17.Range("A2:Z2000").Clear
    L = 2
    For P = 0 To Cont - 1
    P17.Range("A" & L) = LstProd(P, 0)                                           'MODELO
    P17.Range("B" & L) = LstProd(P, 1)                                           'QT
    P17.Range("C" & L) = LstProd(P, 2)                                           'PESO
    P17.Range("D" & L) = LstProd(P, 3)                                           'QT MESA
    P17.Range("E" & L) = LstProd(P, 4)                                           'MAQUINA
    P17.Range("F" & L) = LstProd(P, 5)                                           'CLIENTE
    P17.Range("G" & L) = LstProd(P, 6)                                           'TEMPO (MIN)
    P17.Range("H" & L) = LstProd(P, 7)                                           'SETUP
    P17.Range("I" & L) = LstProd(P, 8)                                           'HORA INICIO
    P17.Range("J" & L) = LstProd(P, 9)                                           'HORA FIM
    P17.Range("K" & L) = LstProd(P, 10)                                          'LOTE MINIMO
    P17.Range("L" & L) = LstProd(P, 11)                                          'PEÇA IRMÃ
    P17.Range("M" & L) = LstProd(P, 12)                                          'DT CART
    P17.Range("N" & L) = LstProd(P, 13)                                          'DT DEPOSITO
    P17.Range("O" & L) = LstProd(P, 14)                                          'OBSERVAÇÃO
    P17.Range("P" & L) = LstProd(P, 15)                                          'STATUS
    P17.Range("Q" & L) = LstProd(P, 16)                                          'APROVAR
    P17.Range("R" & L) = LstProd(P, 17)                                          'Endereço no Relatorio
    L = L + 1
    Next P
End Sub