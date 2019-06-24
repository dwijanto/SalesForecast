Imports Npgsql
Imports System.Text

Public Class BrandModel
    Implements IModel

    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "sales.sfbrand"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "brand"
        End Get
    End Property

    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Dim SB = New StringBuilder

        Using conn As Object = myadapter.getConnection
            conn.Open()
            SB.Append(String.Format("select u.* from {0} u order by {1};", TableName, SortField))
            SB.Append(String.Format("select 0 as id, null::character varying catbrandname union all (select * from sales.catbrandview);"))
            Dim sqlstr = SB.ToString
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Return Nothing
    End Function


End Class
