Public Interface IModel
    ReadOnly Property tablename As String
    ReadOnly Property sortField As String
    Function LoadData(ByVal DS As DataSet) As Boolean
    Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean
End Interface

Public Class ContentBaseEventArgs
    Inherits EventArgs
    Public Property dataset As DataSet
    Public Property message As String
    Public Property hasChanges As Boolean
    Public Property ra As Integer
    Public Property continueonerror As Boolean

    Public Sub New(ByVal dataset As DataSet, ByRef haschanges As Boolean, ByRef message As String, ByRef recordaffected As Integer, ByVal continueonerror As Boolean)
        Me.dataset = dataset
        Me.message = message
        Me.ra = ra
        Me.continueonerror = continueonerror
    End Sub
End Class
