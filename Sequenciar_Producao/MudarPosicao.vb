Private Sub MudarPosicao()
    '-------------------------------------      PASSAR DE B PARA TEMPORÁRIA
    For C = 0 To 17
        LstProd(Cont, C) = LstProd(B, C)
    Next C
    '-------------------------------------      PASSAR DE B PARA A
    For C = 0 To 17
        LstProd(B, C) = LstProd(A, C)
    Next C
    '-------------------------------------      PASSAR DE TEMPORÁRIA PARA B
    For C = 0 To 17
        LstProd(A, C) = LstProd(Cont, C)
    Next C
End Sub