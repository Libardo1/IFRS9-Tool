Imports DataManager

Public Class Matrix
    Implements ICloneable

    Public NumCol As Integer
    Public NumRow As Integer
    Public Mat(,) As Double 'Mat(NumCol, NumRow)
    Protected CurrentNumCol As Integer ' Current Column index
    Protected CurrentNumRow As Integer ' Current Row index

    Public Sub New()
        CurrentNumCol = 0
        CurrentNumRow = 0
    End Sub

    Public Sub New(ByVal NumCol As Integer, ByVal NumRow As Integer)
        setdimensions(NumCol, NumRow)
        CurrentNumCol = 0
        CurrentNumRow = 0
    End Sub

    Public Sub New(ByVal x As Integer)
        setdimensions(x, x)
        CurrentNumCol = 0
        CurrentNumRow = 0
    End Sub

    ' Not best practice
    Public Sub setdimensions(ByVal NumCol As Integer, ByVal NumRow As Integer)
        Me.NumCol = NumCol
        Me.NumRow = NumRow
        ReDim Mat(NumCol - 1, NumRow - 1) ' Because index starts at 0 
    End Sub

    Public Sub AddElement(ByVal element As Double)

        If CurrentNumCol > (NumCol - 1) Then
            CurrentNumRow += 1
            CurrentNumCol = 0
        End If

        Try
            Mat(CurrentNumCol, CurrentNumRow) = element
        Catch e As Exception
            Throw New Exception("Matrix filled with values. " & "element(" _
                & CurrentNumCol & "," & CurrentNumRow & ") doesn't exist.")
        End Try
        CurrentNumCol += 1

    End Sub

    Public Overridable Function Clone() As Object Implements ICloneable.Clone
        Dim temp As Matrix = New Matrix(NumCol, NumRow)
        temp.Mat = Mat.Clone

        'for i as integer = 0 to m - 1
        'for j as integer = 0 to n - 1
        'temp.a(i, j) = a(i, j)
        'next
        'next
        Return temp
    End Function

    Public Function Add(ByVal c As Matrix) As Matrix
        If NumCol <> c.NumCol Or NumRow <> c.NumRow Then
            Throw New Exception("Matrices size mismatch.")
        End If

        Dim b As Matrix = New Matrix(NumCol, NumRow)
        For i As Integer = 0 To NumCol - 1
            For j As Integer = 0 To NumRow - 1
                b.Mat(i, j) = Mat(i, j) + c.Mat(i, j)
            Next
        Next

        Return b
    End Function

    Public Function Subtract(ByVal c As Matrix) As Matrix
        If NumCol <> c.NumCol Or NumRow <> c.NumRow Then
            Throw New Exception("Matrices size mismatch.")
        End If

        Dim b As Matrix = New Matrix(NumCol, NumRow)
        For i As Integer = 0 To NumCol - 1
            For j As Integer = 0 To NumRow - 1
                b.Mat(i, j) = Mat(i, j) - c.Mat(i, j)
            Next
        Next

        Return b
    End Function

    Public Function Multiply(ByVal c As Matrix) As Matrix
        If NumCol <> c.NumRow Then
            Throw New Exception("Matrices size mismatch.")
        End If

        Dim b As Matrix = New Matrix(c.NumCol, NumRow)
        For j As Integer = 0 To NumRow - 1
            For i As Integer = 0 To c.NumCol - 1
                For k As Integer = 0 To NumCol - 1 ' or 0 to c.n - 1
                    b.Mat(i, j) += Mat(k, j) * c.Mat(i, k)
                Next
            Next
        Next

        Return b
    End Function

    'DEBUG: OK
    Public Overridable Overloads Function MultiplyPower(Power As Integer) As Matrix
        If NumCol <> NumRow Then
            Throw New Exception("Matrices size mismatch.")
        End If

        Dim b As Matrix = New Matrix(NumCol, NumRow)
        b = Me
        For n As Integer = 2 To Power
            b = b.Multiply(Me)
        Next

        Return b
    End Function

    Public Function SubMatrix(ByVal x As Integer, ByVal y As Integer) As Matrix
        Dim s As Matrix = New Matrix(NumCol - 1, NumRow - 1)
        For j As Integer = 0 To NumRow - 1
            For i As Integer = 0 To NumCol - 1
                If (i <> x And j <> y) Then
                    s.AddElement(Mat(i, j))
                End If
            Next
        Next

        Return s
    End Function

    Public Function Determinant() As Double
        If NumCol = 1 And NumRow = 1 Then
            Return Mat(0, 0)
        End If

        Dim temp As Double
        Dim mysubmatrix As Matrix
        Dim j As Integer = 0
        For i As Integer = 0 To NumCol - 1
            mysubmatrix = SubMatrix(i, j)
            temp = temp + ((-1) ^ (i + j) * Mat(i, j) * mysubmatrix.Determinant())
        Next
        Return temp

    End Function

    Public Function GeDeterminant() As Double
        If issquare() Then
            Return GeDeterminant(True)
        Else
            Throw New Exception("Determinant exists only possible" &
                "for a sqaure matrix.")
        End If
    End Function

    Private Function GeDeterminant(ByVal doclone As Boolean) As Double
        If NumCol = 1 And NumRow = 1 Then
            Return Mat(0, 0)
        End If

        Dim y As Integer = 0
        Dim k As Integer
        For i As Integer = 0 To NumCol - 1
            If Mat(i, y) <> 0 Then
                k = i
                Exit For
            End If
        Next

        Dim temp As Double
        Dim newmatrix As Matrix
        If doclone Then
            newmatrix = Clone()
        Else
            newmatrix = Me
        End If
        Dim f As Double
        For i As Integer = k + 1 To NumCol - 1
            If Mat(i, y) <> 0 Then
                f = Mat(i, y) / Mat(k, y)
                For j As Integer = 0 To NumRow - 1
                    newmatrix.Mat(i, j) = Mat(i, j) - Mat(k, j) * f
                Next
                'msgbox(newmatrix.tostring())
            End If
        Next
        newmatrix = newmatrix.SubMatrix(k, y) 'save space
        temp += ((-1) ^ (k + y)) * Mat(k, y) * newmatrix.GeDeterminant(False)

        Return temp

    End Function

    Public Function Det() As Double
        Dim srcmatrix As Matrix = Me
        Dim destmatrix As Matrix

        Dim r As Integer = 0
        Dim k As Integer
        Dim mul As Double = 1

        Do
            destmatrix = New Matrix(srcmatrix.NumCol - 1, srcmatrix.NumRow - 1)
            For x As Integer = 0 To srcmatrix.NumCol - 1
                If srcmatrix.Mat(x, r) <> 0 Then
                    k = x
                    Exit For
                End If
            Next

            Dim f As Double
            mul = mul * ((-1) ^ (k + r)) * srcmatrix.Mat(k, r)
            For i As Integer = 0 To srcmatrix.NumCol - 1
                If i <> k Then
                    f = srcmatrix.Mat(i, r) / srcmatrix.Mat(k, r)
                    For j As Integer = 0 To srcmatrix.NumRow - 1
                        If j <> r Then
                            destmatrix.Mat(subindex(i, k), subindex(j, r)) = srcmatrix.Mat(i, j) - srcmatrix.Mat(k, j) * f
                        End If
                    Next
                End If
            Next
            srcmatrix = destmatrix
        Loop Until srcmatrix.NumCol = 1

        Return mul * srcmatrix.Mat(0, 0)
    End Function

    Private Function subindex(ByVal x As Integer, ByVal k As Integer) As Integer
        Return IIf(x <= k, x, x - 1)
    End Function

    Public Function rotate(ByVal theetadeg As Double) As Matrix
        Dim theetarad As Double = theetadeg * Math.PI / 180 ' angle in radians
        Dim rotationmatrix As Matrix = New Matrix(2)
        rotationmatrix.AddElement(Math.Cos(theetarad))
        rotationmatrix.AddElement(-Math.Sin(theetarad))
        rotationmatrix.AddElement(Math.Sin(theetarad))
        rotationmatrix.AddElement(Math.Cos(theetarad))
        MsgBox(rotationmatrix.tostring)
        Return rotationmatrix.Multiply(Me)
    End Function

    Public Overrides Function tostring() As String
        Dim temp As String = ""
        For y As Integer = 0 To NumRow - 1
            For x As Integer = 0 To NumCol - 1
                temp &= Mat(x, y) & ","
            Next x
            temp &= Chr(13)
        Next y

        Return temp
    End Function

    Public Function issquare() As Boolean
        Return (NumCol = NumRow)
    End Function

    Public ReadOnly Property GetElement(numCol As Integer, numRow As Integer) As Double
        Get
            GetElement = Mat(numCol, numRow)
        End Get
    End Property

    'DEBUG: OK
    Public Function RetrieveTPs(Ref_Date As Date, DB As AccessIO) As Matrix
        DB.OpenConnection()

        'MAXIME: ID_Matrix = 1 is needed in the DB
        'Do we need it ? Yes if several TM models have been set up
        Dim ID_Matrix As Integer = 1

        'Set Matrix size
        Dim QuerySize As String = "Select DISTINCT RatingFrom
                               FROM [TransitionProbabilities]
                               WHERE ID_Matrix = " + ID_Matrix.ToString + ""
        DB.CreateCommand(QuerySize)
        DB.ExecuteReader()

        Dim MatSize As Integer = 0
        While (DB.Reader.Read())
            MatSize += 1
        End While

        setdimensions(MatSize, MatSize)

        'select TMs and create transition probability
        Dim Query As String = "Select * 
                               FROM [TransitionProbabilities]
                               WHERE ID_Matrix = " + ID_Matrix.ToString + " And
                                    Date_Start <= #" + Ref_Date.ToString + "# AND Date_End > #" + Ref_Date.ToString + "#
                               ORDER BY RatingFrom, RatingTo"

        DB.CreateCommand(Query)
        DB.ExecuteReader()

        Dim TP As TransitionProbability
        Dim Prob As Double

        Dim count As Integer = 0
        While (DB.Reader.Read())
            TP = New TransitionProbability
            TP.RetrieveTP(DB)
            Prob = TP.GetProb
            AddElement(Prob)
            count += 1
        End While

        DB.CloseReader()
        DB.CloseConnection()

        Return Me
    End Function


End Class
