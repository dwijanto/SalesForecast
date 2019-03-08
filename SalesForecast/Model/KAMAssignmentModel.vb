Imports Npgsql

Public Class KAMAssignmentModel
    Implements IModel
    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance

    Public Property recid As Integer
    Public Property username As String
    Public Property mla As String
    Public Property cardname As String
    Public Property assingment As String
    Public Property sdasd As Double
    Public Property tefalsd As Double
    Public Property lagosd As Double
    Public Property wmfsd As Double


    Public Property kam As String
    Public Property mlacardnameid As Long

    Dim KAMAssingmentList As List(Of KAMAssignmentModel)

    Public Function LoadData(DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myadapter.getConnection
            conn.Open()
            Dim sqlstr = String.Format("select u.* from {0} u order by {1}", tablename, sortField)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, tablename)
            myret = True
        End Using
        Return myret
    End Function

    Public Function GetLastId() As Long
        Dim sqlstr = String.Format("select id from {0}  order by id desc limit 1", tablename)
        Dim myret As Long = 0
        Using conn As Object = myadapter.getConnection
            conn.open()
            myadapter.ExecuteScalar(sqlstr, recordAffected:=myret)
        End Using
        Return myret
    End Function

    Public Function PopulateKAMAssignment(criteria As String) As List(Of KAMAssignmentModel)
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim DS As DataSet = New DataSet
        Using conn As Object = myadapter.getConnection
            conn.Open()
            'Dim sqlstr = String.Format("select 1 as id,row_number() over (order by mc.mla,mc.cardname) as recid,kam.username,mc.mla,mc.cardname,trim(mc.mla) || ' - ' || mc.cardname as assignment from sales.sfkam kam" &
            '                           " left join sales.sfkamassignment ka on ka.kam = kam.username left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid where kam.username = '{0}'" &
            '                           " order by mc.mla", criteria)

            Dim sqlstr = String.Format("with ka as (select distinct kam.username,mc.mla,mc.cardname,trim(mc.mla) || ' - ' || mc.cardname as assignment,sdasd,tefalsd,lagosd,wmfsd  from sales.sfkam kam  left join sales.sfkamassignment ka on ka.kam = kam.username " &
                                       " left join sales.sfmlacardname mc on mc.id = ka.mlacardnameid where kam.username = '{0}' order by mc.mla) select 1 as id,row_number() over (order by ka.mla,ka.cardname) as recid,* from ka", criteria)

            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, tablename)
            KAMAssingmentList = New List(Of KAMAssignmentModel)
            For Each dr As DataRow In DS.Tables(0).Rows
                KAMAssingmentList.Add(New KAMAssignmentModel With {.recid = dr.Item("recid"),
                                                                   .username = dr.Item("username"),
                                                                   .mla = "" & dr.Item("mla"),
                                                                   .cardname = "" & dr.Item("cardname"),
                                                                   .assingment = "" & dr.Item("assignment"),
                                                                   .sdasd = dr.Item("sdasd"),
                                                                   .tefalsd = dr.Item("tefalsd"),
                                                                   .lagosd = dr.Item("lagosd"),
                                                                   .wmfsd = dr.Item("wmfsd")
                                                                   })
            Next
        End Using
        Return KAMAssingmentList
    End Function

    Public Function Add(ByVal model As KAMAssignmentModel) As Long
        Dim myret As Long
        Dim sqlstr = String.Format("insert into sales.sfkamassignment(mlacardnameid,kam) values('{0}','{1}'); select currval('sales.sfkamassignment_id_seq');", model.mlacardnameid, model.kam)
        If Not myadapter.ExecuteScalar(sqlstr, recordAffected:=myret) Then
            Throw New Exception(String.Format("KamAssignmentModel.Add {0}", myadapter.ErrorMessage))
        End If
        Return myret
    End Function

    Public Function save(obj As Object, mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Return Nothing
    End Function

    Public ReadOnly Property sortField As String Implements IModel.sortField
        Get
            Return "kam"
        End Get
    End Property

    Public ReadOnly Property tablename As String Implements IModel.tablename
        Get
            Return "Sales.sfkamassignment"
        End Get
    End Property
End Class
