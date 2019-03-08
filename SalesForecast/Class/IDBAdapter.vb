Public Interface IDBAdapter
    Sub DBAdapterInitialize()
    Property ConnectionString() As String
    Function GetDataset(sqlstr As String, ds As DataSet, Optional params As List(Of IDataParameter) = Nothing) As Boolean
    Function GetTable(sqlstr As String, dt As DataTable, Optional params As List(Of IDataParameter) = Nothing) As Boolean
    Function ExecuteScalar(ByVal sqlstr As String, Optional ByVal params As List(Of IDataParameter) = Nothing, Optional ByRef recordAffected As Object = Nothing, Optional ByRef message As String = "") As Boolean
    Function ExecuteNonQuery(ByVal sqlstr As String, Optional ByVal params As List(Of IDataParameter) = Nothing, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
    Function CanConnect() As Boolean
End Interface
