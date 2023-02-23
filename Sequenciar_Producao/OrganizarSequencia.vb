Private Sub OrganizarSeguencia()
'-----------------------------------------------        ENCONTRAR A MENOR DATA
For A = 0 To Cont - 1
    If LstProd(A, 14) <> "" Then                                'COLOCAR PARA CIMA MODELOS APTOS PARA PROGRAMAR
        
        For B = A To Cont - 1
            If LstProd(B, 14) = "" Then
                Enc = True
                Exit For
            End If
        Next B
            If Enc = True Then
                Enc = False
                For C = 0 To 17
                    LstProd(Cont, C) = LstProd(B, C)
                Next C
                
                For C = 0 To 17
                    LstProd(B, C) = LstProd(A, C)
                Next C
                
                For C = 0 To 17
                    LstProd(A, C) = LstProd(Cont, C)
                Next C
                
                For N = 0 To 17
                    LstProd(Cont, N) = Null
                Next N
            Else
                Exit For
            End If
        End If
    Next A
End Sub