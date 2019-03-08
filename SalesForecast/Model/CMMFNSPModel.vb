Imports Npgsql

Public Class CMMFNSPModel
    Implements IModel

    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance
    Public Property cmmf As Long
    Public Property nsp1 As Decimal
    Public Property nsp2 As Decimal

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "sales.sfcmmfnsp"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "cmmf"
        End Get
    End Property
    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.* from {0} u order by {1}", TableName, SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function
    Public Function LoadData(ByVal DS As DataSet, ByVal criteria As String) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u. from {0} {1} order by {2}", TableName, criteria, SortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, TableName)
            myret = True
        End Using
        Return myret
    End Function
    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf myadapter.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = myadapter.getConnection
            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            Dim sqlstr = "sales.sp_updatesfcmmfnsp"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nsp1").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nsp2").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_insertsfcmmfnsp"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nsp1").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "nsp2").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_deletesfcmmfnsp"
            dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "cmmf").Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(TableName))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function
End Class
