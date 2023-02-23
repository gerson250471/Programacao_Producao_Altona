Private Sub PriorizaData()
    For A = 0 To Cont - 1
        For B = A + 1 To Cont - 1
            If LstProd(A, 13) > LstProd(B, 13) Then
                If LstProd(B, 14) = "" Then
                    MudarPosicao
                End If
            End If
        Next B
    Next A
End Sub