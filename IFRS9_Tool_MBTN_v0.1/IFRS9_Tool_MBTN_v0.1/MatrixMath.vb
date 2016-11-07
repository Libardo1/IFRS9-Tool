Public Class MatrixMathOriginal
    '    Public Function M_EXP(Mat() As Double, Optional Algo As String = "P", Optional N As Long = 0)
    '        '-------------------------------------------------------------------
    '        'return the exponential matrix expansion
    '        'it use two alternative algorithms:
    '        'Algo = P  uses the Pade'expansion, kindly developed by Gregory Klein
    '        'Algo = W  uses the power's serie.
    '        'n (only for algo=W) sets the max for the serie
    '        '-------------------------------------------------------------------

    '        If UCase(Algo) = "W" Then
    '            M_EXP = M_exp2(Mat, N) 'TN: dimension?

    '        Else
    '            M_EXP = M_Exp3(Mat)
    '        End If
    '    End Function

    '    '------------------------------------------------------------------------------
    '    'matrix exponential with pade approximation
    '    'by Gregory Klein  12.5.2003
    '    '------------------------------------------------------------------------------
    '    Private Function M_Exp3(Mat)
    '        Dim p As Boolean
    '        ' Scale Mat by power of 2 so that its norm is < 1/2 .

    '        Dim s, A, N, x, c, e, d, q, cX ' but what are they 

    '        s = Math.Max(0, Int(Math.Log(M_ABS(Mat), 2)) + 1)
    '        A = M_PRODS(Mat, 0.5 ^ s)
    '        N = UBound(A)
    '        '% Pade approximation for exp(A)
    '        x = A
    '        c = 1 / 2
    '        e = M_ADD(M_IDENT_(N), M_PRODS(A, c))
    '        d = M_SUB(M_IDENT_(N), M_PRODS(A, c))
    '        q = 6
    '        p = 1
    '        For k = 2 To q
    '            c = c * (q - k + 1) / (k * (2 * q - k + 1))
    '            x = Math.MMult(A, x)
    '            cX = M_PRODS(x, c)
    '            e = M_ADD(e, cX)
    '            If p Then
    '                d = M_ADD(d, cX)
    '            Else
    '                d = M_SUB(d, cX)
    '            End If
    '            p = Not (p)
    '        Next
    '        e = Application.MMult(M_INV(d), e)
    '        '% Undo scaling by repeated squaring
    '        For k = 1 To s
    '            e = Application.MMult(e, e)
    '        Next
    '        M_Exp3 = e
    '    End Function
    '    Private Function M_exp2(Mat, Optional N)
    '        'returns the matrix series expansion
    '        'exp(A)= I + A + 1/2*A^2 +1/6*A^3 +...1/n!*A^n + error
    '        Dim A, b, b1, c
    '        Dim Flag_End_Loop As Boolean
    '        Const tiny = 10 ^ -15
    '        A = Mat
    '        If UBound(A, 1) <> UBound(A, 2) Then
    '            M_exp2 = "?"   'only square matrix please !
    '            Exit Function
    '        End If
    '        'series expansion begins
    '        b = A
    '        m = UBound(A, 1)
    '        c = M_ADD(M_IDENT_(m), A)  'C=I+A
    '        'For k = 2 To n
    '        k = 1
    '    GoSub Check_End_Loop
    '    Do Until Flag_End_Loop
    '            k = k + 1
    '            b1 = Application.WorksheetFunction.MMult(b, A)
    '            b = M_PRODS(b1, 1 / k)  'B= 1/k*B*A
    '            c = M_ADD(c, b)
    '        GoSub Check_End_Loop
    '    Loop
    '        M_exp2 = c
    '        Exit Function
    '        '-----------------------------
    'Check_End_Loop:
    '        If IsMissing(N) Then
    '            myErr = M_ABS(b)
    '            Flag_End_Loop = (myErr < tiny)
    '        Else
    '            Flag_End_Loop = (k >= N)
    '        End If
    '        Return

    '    End Function

    '    Function M_EXP_ERR(Mat, N)
    '        'returns the truncation error of matrix series expansion
    '        'exp(A)= I + A + 1/2*A^2 +1/6*A^3 +...1/n!*A^n + error
    '        Dim A, b, b1, c
    '        A = Mat
    '        If UBound(A, 1) <> UBound(A, 2) Then
    '            M_EXP_ERR = "?"   'only square matrix please !
    '            Exit Function
    '        End If
    '        'series expansion begins
    '        b = A
    '        m = UBound(A, 1)
    '        c = M_ADD(M_IDENT_(m), A)  'C=I+A
    '        For k = 2 To N
    '            b1 = Application.WorksheetFunction.MMult(b, A)
    '            b = M_PRODS(b1, 1 / k)  'B= 1/k*B*A
    '            c = M_ADD(c, b)
    '        Next
    '        myErr = M_ABS(b)
    '        M_EXP_ERR = myErr
    '    End Function

    '    Function M_PRODS(Mat, scalar)
    '        'multiplies a scalar for a matrix
    '        Dim b, k
    '        k = scalar
    '        b = Mat
    '        For i = 1 To UBound(b, 1)
    '            For j = 1 To UBound(b, 2)
    '                b(i, j) = k * b(i, j)
    '            Next j
    '        Next i
    '        M_PRODS = b
    '    End Function

    '    Function M_ABS(v)
    '        'Absolute of matrix (Euclidean norm)
    '        Dim A, N As Integer, m As Integer, i As Integer, j As Integer
    '        A = v
    '        N = UBound(A, 1)
    '        m = UBound(A, 2)
    '        s = 0
    '        For i = 1 To N
    '            For j = 1 To m
    '                s = s + A(i, j) ^ 2
    '            Next j
    '        Next i
    '        M_ABS = Sqr(s)
    '    End Function

    '    Private Function M_IDENT_(N)
    '        Dim A() As Double
    '        ReDim A(1 To N, 1 To N)
    '        For i = 1 To N
    '            A(i, i) = 1
    '        Next
    '        M_IDENT_ = A
    '    End Function

    '    Function M_ADD(Mat1, Mat2)
    '        'matrix addition
    '        Dim A, b, c()
    '        Dim na As Integer, ma As Integer, nb As Integer, mb As Integer
    '        Dim i As Integer, j As Integer
    '        A = Mat1 : b = Mat2
    '        na = UBound(A, 1) : ma = UBound(A, 2)
    '        ReDim c(1 To na, 1 To ma)
    '        For i = 1 To na
    '            For j = 1 To ma
    '                c(i, j) = A(i, j) + b(i, j)
    '            Next j
    '        Next i
    '        M_ADD = c
    '    End Function

    '    Function M_SUB(Mat1, Mat2)
    '        'matrix subtraction
    '        Dim A, b, c()
    '        Dim na As Integer, ma As Integer, nb As Integer, mb As Integer
    '        Dim i As Integer, j As Integer
    '        A = Mat1 : b = Mat2
    '        na = UBound(A, 1) : ma = UBound(A, 2)
    '        ReDim c(1 To na, 1 To ma)
    '        For i = 1 To na
    '            For j = 1 To ma
    '                c(i, j) = A(i, j) - b(i, j)
    '            Next j
    '        Next i
    '        M_SUB = c
    '    End Function
End Class

