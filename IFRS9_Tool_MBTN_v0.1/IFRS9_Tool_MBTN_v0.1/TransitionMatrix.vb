'Imports DataManager

'Public Class TransitionMatrix
'    Inherits Matrix

'    Public Sub New()
'        MyBase.CurrentNumCol = 0
'        MyBase.CurrentNumRow = 0
'    End Sub


'    Public Function RetrieveTPs(Ref_Date As Date, DB As AccessIO) As TransitionMatrix
'        DB.OpenConnection()

'        'MAXIME: ID_Matrix = 1 is needed in the DB
'        'Do we need it ? Yes if several TM models have been set up

'        Dim ID_Matrix As Integer = 1
'        Dim Query As String = "Select * 
'                               FROM [TransitionProbabilities]
'                               WHERE ID_Matrix = " + ID_Matrix.ToString + " AND
'                                     Date_Start <= #" + Ref_Date.ToString + "# AND Date_End > #" + Ref_Date.ToString + "#"

'        DB.CreateCommand(Query)
'        DB.ExecuteReader()

'        'Set Matrix size
'        Dim MatSize As Integer = DB.Reader.FieldCount 'CHECK IF IT WORKS
'        setdimensions(MatSize, MatSize)

'        'create and add transition probabilities
'        Dim TP As TransitionProbability
'        Dim Prob As Double
'        While (DB.Reader.Read())
'            TP = New TransitionProbability
'            TP.RetrieveTP(DB)
'            Prob = TP.GetProb
'            AddElement(Prob)
'        End While

'        DB.CloseReader()
'        DB.CloseConnection()

'        Return Me
'    End Function

'    Public Function Multiply(ByVal p As TransitionMatrix) As TransitionMatrix
'        Dim c As New TransitionMatrix
'        c = p
'        If NumCol <> c.NumRow Then
'            Throw New Exception("Matrices size mismatch.")
'        End If

'        Dim b As TransitionMatrix = New TransitionMatrix()
'        setdimensions(c.NumCol, NumRow)
'        For j As Integer = 0 To NumRow - 1
'            For i As Integer = 0 To c.NumCol - 1
'                For k As Integer = 0 To NumCol - 1 ' or 0 to c.n - 1
'                    b.Mat(i, j) += Mat(k, j) * c.Mat(i, k)
'                Next
'            Next
'        Next

'        Return b
'    End Function

'    'Note: both MultiplyPower should work. Try them.
'    Public Function MultiplyPower(ByVal c As TransitionMatrix, Power As Integer) As TransitionMatrix
'        If NumCol <> c.NumRow Then
'            Throw New Exception("Matrices size mismatch.")
'        End If

'        Dim b As TransitionMatrix = New TransitionMatrix()
'        setdimensions(c.NumCol, NumRow)
'        If Power = 1 Then
'            b = c
'        Else
'            b = c
'            For n As Integer = 2 To Power
'                b = b.Multiply(c)
'            Next

'        End If

'        Return b
'    End Function

'    Public Function MultiplyPower(Power As Integer) As TransitionMatrix
'        If NumCol <> NumRow Then
'            Throw New Exception("Matrices size mismatch.")
'        End If

'        Dim b As TransitionMatrix = New TransitionMatrix()
'        setdimensions(NumCol, NumRow)
'        b = Me
'        For n As Integer = 1 To Power
'            b = b.Multiply(Me)
'        Next

'        Return b
'    End Function

'    ''Note: overrides doesnt work cause return type different
'    'Public Overloads Overrides Function MultiplyPower(ByVal c As Matrix, Power As Integer) As TransitionMatrix
'    '    If NumCol <> c.NumRow Then
'    '        Throw New Exception("Matrices size mismatch.")
'    '    End If

'    '    Dim b As Matrix = New Matrix(c.NumCol, NumRow)
'    '    If Power = 1 Then
'    '        b = c
'    '    Else
'    '        b = c
'    '        For n As Integer = 2 To Power
'    '            b = b.Multiply(c)
'    '        Next

'    '    End If

'    '    Return b
'    'End Function

'    'Public Overloads Overrides Function MultiplyPower(Power As Integer) As TransitionMatrix
'    '    If NumCol <> NumRow Then
'    '        Throw New Exception("Matrices size mismatch.")
'    '    End If

'    '    Dim b As Matrix = New Matrix(NumCol, NumRow)
'    '    b = Me
'    '    For n As Integer = 1 To Power
'    '        b = b.Multiply(Me)
'    '    Next

'    '    Return b
'    'End Function

'End Class
