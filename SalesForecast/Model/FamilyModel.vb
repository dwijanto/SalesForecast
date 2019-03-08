Imports Npgsql

Public Class FamilyModel
    Implements IModel

    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance
    Public ReadOnly Property FilterField
        Get
            Return "[familyname] like '*{0}*' or [familylv2] like '*{0}*' or [productlinedesc] like '*{0}*'  "
        End Get
    End Property

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "sales.sffamily"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "familyid"
        End Get
    End Property

    Public Function LoadData(ByVal DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.*,to_char(u.familyid,'FM000') || ' - ' || u.familyname || ' - ' || u.familylv2 as familydesc from {0} u order by {1}; select distinct productlinedesc from sales.sffamily order by productlinedesc;", TableName, SortField)
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
            Dim sqlstr = "sales.sp_updatesffamily"
            dataadapter.UpdateCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familyname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familylv2").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "productlinedesc").SourceVersion = DataRowVersion.Current


            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_insertsffamily"
            dataadapter.InsertCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familyname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familylv2").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "productlinedesc").SourceVersion = DataRowVersion.Current            
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_deletesffamily"
            dataadapter.DeleteCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").Direction = ParameterDirection.Input
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
