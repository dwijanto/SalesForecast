Imports Npgsql
Imports System.IO

Public Class PostgreSQLDbAdapter

    Implements IDisposable
    Implements IDBAdapter
    Dim _userid As String
    Dim _password As String
    Dim _connectionstring As String
    Public ErrorMessage As String
    Private Builder As New NpgsqlConnectionStringBuilder


    Public Property ConnectionString As String Implements IDBAdapter.ConnectionString
        Get
            Return _connectionstring
        End Get
        Set(value As String)
            _connectionstring = value
        End Set
    End Property

    Private Shared myInstance As PostgreSQLDbAdapter

    Public ReadOnly Property UserId As String
        Get
            Return _userid
        End Get
    End Property

    Public ReadOnly Property Password As String
        Get
            Return _password
        End Get
    End Property
    Public Shared Function getInstance() As PostgreSQLDbAdapter
        If myInstance Is Nothing Then
            myInstance = New PostgreSQLDbAdapter
        End If
        Return myInstance
    End Function

    Public ReadOnly Property Host As String
        Get
            Return Builder.Host
        End Get
    End Property

    Public ReadOnly Property DB As String
        Get
            Return Builder.Database
        End Get
    End Property

    Private Sub New()
        DBAdapterInitialize()
    End Sub

    Private Sub DBAdapterInitialize() Implements IDBAdapter.DBAdapterInitialize
        _userid = "admin"
        _password = "admin"
        _connectionstring = String.Format("{0}{1}", My.Settings.PostgreSQLCon, "User=admin;Password=admin")
        Builder.ConnectionString = _connectionstring
    End Sub

    Public Function CanConnect() As Boolean Implements IDBAdapter.CanConnect
        Dim myret As Boolean = False
        Using conn As New NpgsqlConnection(_connectionstring)
            Try
                conn.Open()
                myret = True
            Catch ex As Exception
                ErrorMessage = ex.Message
            End Try
        End Using
        Return myret
    End Function

    Public Function ExecuteNonQuery(sqlstr As String, Optional params As List(Of IDataParameter) = Nothing, Optional ByRef recordAffected As Long = 0, Optional ByRef message As String = "") As Boolean Implements IDBAdapter.ExecuteNonQuery
        Dim myret As Boolean = False
        Using conn As New NpgsqlConnection(_connectionstring)
            conn.Open()
            Using cmd As NpgsqlCommand = New NpgsqlCommand()
                cmd.CommandText = sqlstr
                cmd.Connection = conn
                Try
                    If Not IsNothing(params) Then
                        For Each param As IDataParameter In params
                            cmd.Parameters.Add(param)
                        Next
                    End If
                    recordAffected = cmd.ExecuteNonQuery
                    myret = True
                Catch ex As Exception
                    ErrorMessage = ex.Message
                End Try
            End Using
        End Using
        Return myret
    End Function

    Public Function ExecuteScalar(sqlstr As String, Optional params As List(Of IDataParameter) = Nothing, Optional ByRef recordAffected As Object = Nothing, Optional ByRef message As String = "") As Boolean Implements IDBAdapter.ExecuteScalar
        Dim myret As Boolean = False
        Using conn As New NpgsqlConnection(_connectionstring)
            conn.Open()
            Using cmd As NpgsqlCommand = New NpgsqlCommand()
                cmd.CommandText = sqlstr
                cmd.Connection = conn
                Try
                    If Not IsNothing(params) Then
                        For Each param As NpgsqlParameter In params
                            cmd.Parameters.Add(param)
                        Next
                    End If
                    recordAffected = cmd.ExecuteScalar
                    myret = True
                Catch ex As Exception
                    ErrorMessage = ex.Message
                End Try
            End Using
        End Using
        Return myret
    End Function

    Public Function GetDataset(sqlstr As String, ds As DataSet, Optional params As List(Of IDataParameter) = Nothing) As Boolean Implements IDBAdapter.GetDataset
        Dim DataAdapter As IDbDataAdapter = New NpgsqlDataAdapter
        Dim myret As Boolean = False
        Using conn As New NpgsqlConnection(_connectionstring)
            conn.Open()
            Using cmd As NpgsqlCommand = New NpgsqlCommand()
                cmd.CommandText = sqlstr
                cmd.Connection = conn
                DataAdapter.SelectCommand = cmd
                If Not IsNothing(params) Then
                    For Each param As IDataParameter In params
                        cmd.Parameters.Add(param)
                    Next
                End If
                DataAdapter.Fill(ds)
                myret = True
            End Using
        End Using
        Return myret
    End Function

    Public Function GetTable(sqlstr As String, dt As DataTable, Optional params As List(Of IDataParameter) = Nothing) As Boolean Implements IDBAdapter.GetTable
        Return Nothing
    End Function
    Public Function getConnection() As NpgsqlConnection
        'If IsNothing(_userid) Or IsNothing(_password) Then
        '    Throw New DbAdapterExeption("User Id or Password is blank.")
        'End If
        Return New NpgsqlConnection(_connectionstring)
    End Function
    Public Function getDbDataAdapter() As NpgsqlDataAdapter
        Return New NpgsqlDataAdapter
    End Function

    Public Function getCommandObject() As NpgsqlCommand
        Return New NpgsqlCommand
    End Function

    Public Function getCommandObject(ByVal sqlstr As String, ByVal connection As Object) As NpgsqlCommand
        Return New NpgsqlCommand(sqlstr, connection)
    End Function

    Public Sub onRowInsertUpdate(ByVal sender As Object, ByVal e As NpgsqlRowUpdatedEventArgs)
        If e.StatementType = StatementType.Insert Or e.StatementType = StatementType.Update Then
            If e.Status <> UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If
        End If
    End Sub

    Public Function getParam(ByVal ParameterName As String,
                            Optional ByVal value As Object = Nothing,
                        Optional ByVal dbType As DbType = Nothing,
                        Optional ByVal direction As ParameterDirection = ParameterDirection.Input,
                        Optional isNullable As Boolean = False,
                        Optional Precision As Byte = 0,
                        Optional scale As Byte = 0,
                        Optional size As Integer = Integer.MaxValue,
                        Optional SourceColumn As String = "",
                        Optional sourceversion As DataRowVersion = DataRowVersion.Current) As NpgsqlParameter
        Dim myparam = New NpgsqlParameter
        With myparam
            .ParameterName = ParameterName
            .Value = value
            .DbType = dbType
            .Direction = direction
            .IsNullable = isNullable
            .Precision = Precision
            .Scale = scale
            .Size = size
            .SourceColumn = SourceColumn
            .SourceVersion = sourceversion
        End With
        Return myparam
    End Function


    Public Function copy(ByVal sqlstr As String, ByVal InputString As String, Optional ByRef result As Boolean = False) As String
        result = False
        Dim CopyIn1 As NpgsqlCopyIn
        Dim myReturn As String = ""
        'Convert string to MemoryStream
        Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.ASCII.GetBytes(InputString))
        'Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.Default.GetBytes(InputString))
        Dim buf(9) As Byte
        Dim CopyInStream As Stream = Nothing
        Dim i As Long
        Using conn = New NpgsqlConnection(_connectionstring)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                CopyIn1 = New NpgsqlCopyIn(command, conn)
                Try
                    CopyIn1.Start()
                    CopyInStream = CopyIn1.CopyStream
                    i = MemoryStream1.Read(buf, 0, buf.Length)
                    While i > 0
                        CopyInStream.Write(buf, 0, i)
                        i = MemoryStream1.Read(buf, 0, buf.Length)
                        Application.DoEvents()
                    End While
                    CopyInStream.Close()
                    result = True
                Catch ex As NpgsqlException
                    Try
                        CopyIn1.Cancel("Undo Copy")
                        myReturn = ex.Message & ", " & ex.Detail
                    Catch ex2 As NpgsqlException
                        If ex2.Message.Contains("Undo Copy") Then
                            myReturn = ex2.Message & ", " & ex2.Detail
                        End If
                    End Try
                End Try

            End Using
        End Using

        Return myReturn
    End Function
    Public Shared Function validint(ByVal p1 As Object) As Object
        If p1 = "" Then
            Return "Null"
        Else
            Return CInt((Replace(p1, ",", "")))
        End If
    End Function
    Public Shared Function validnumeric(ByVal p1 As Object) As Object
        If p1 = "" Then
            Return "Null"
        Else
            Return CDec((Replace(p1, ",", "")))
        End If
    End Function
    Public Shared Function validlong(ByVal p1 As Object) As Object
        If p1 = "" Then
            Return "Null"
        Else
            Return CLng(p1)
        End If
    End Function
    Public Shared Function validstr(ByVal p1 As Object) As Object
        If p1 = "" Then
            Return "Null"
        Else
            Return Replace(Replace(p1, Chr(9), " "), "'", "''").Replace(Chr(10), "")
        End If
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                myInstance = Nothing
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region






End Class
Public Class DbAdapterExeption
    Inherits ApplicationException
    Public Sub New(ByVal errormessage As String)
        MyBase.New(errormessage)
    End Sub
End Class