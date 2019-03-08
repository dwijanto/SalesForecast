Imports Npgsql
Public Class KAMGroupSGModel
    Implements IModel
    Dim myadapter As PostgreSQLDbAdapter = PostgreSQLDbAdapter.getInstance


    Public Property kam As String
    Public Property groupid As Long
    Public Property groupname As String
    Public Property TDSDA As Decimal
    Public Property TDRowenta As Decimal
    Public Property TDCW As Decimal
    Public Property SDSDA As Decimal
    Public Property SDCW As Decimal

    Dim KAMGroupSGList As List(Of KAMGroupSGModel)

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


    Public Function PopulateKAMGroupSG(criteria As String) As List(Of KAMGroupSGModel)
        Dim dataadapter As NpgsqlDataAdapter = myadapter.getDbDataAdapter
        Dim DS As DataSet = New DataSet
        Using conn As Object = myadapter.getConnection
            conn.Open()

            'Dim sqlstr = String.Format("select distinct kam.username,kg.groupid,g.groupname from sales.sfkam kam  " &
            '                           " left join sales.sfkamgroupsg kg on kg.kam = kam.username " &
            '                           " left join sales.sfgroup g on g.id = kg.groupid  " &
            '                           " where kam.username = '{0}' order by 1,3", criteria)
            Dim sqlstr = String.Format("with kg as (select distinct kam.username,kg.groupid,g.groupname from sales.sfkam kam   " &
                " left join sales.sfkamgroupsg kg on kg.kam = kam.username  left join sales.sfgroup g on g.id = kg.groupid   " &
                " where kam.username = '{0}' order by 1,3)" &
                " select kg.*,sdsda.sd as sdsda,sdcw.sd as sdcw,td.sda as tdsda,td.rowenta as tdrowenta,td.cw as tdcw from kg" &
                " left join sales.sfgroupsdsg sdsda on sdsda.groupid = kg.groupid and sdsda.producttype = 1" &
                " left join sales.sfgroupsdsg sdcw on sdcw.groupid = kg.groupid and sdcw.producttype = 4" &
                " left join sales.sfgrouptradediscsg td on td.groupid = kg.groupid", criteria)
            dataadapter.SelectCommand = myadapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, tablename)
            KAMGroupSGList = New List(Of KAMGroupSGModel)
            For Each dr As DataRow In DS.Tables(0).Rows
                KAMGroupSGList.Add(New KAMGroupSGModel With {.groupid = dr.Item("groupid"),
                                                             .kam = dr.Item("username"),
                                                             .groupname = "" & dr.Item("groupname"),
                                                             .SDSDA = dr.Item("sdsda"),
                                                             .SDCW = dr.Item("sdcw"),
                                                             .TDCW = dr.Item("tdcw"),
                                                             .TDRowenta = dr.Item("tdrowenta"),
                                                             .TDSDA = dr.Item("tdsda")
                                                             })
            Next
        End Using
        Return KAMGroupSGList
    End Function

    Public Function Add(ByVal model As KAMGroupMSModel) As Long
        Dim myret As Long
        Dim sqlstr = String.Format("insert into sales.sfkamgroupsg(groupid,kam) values('{0}','{1}'); select currval('sales.sfkamgroupsg_id_seq');", model.groupid, model.kam)
        If Not myadapter.ExecuteScalar(sqlstr, recordAffected:=myret) Then
            Throw New Exception(String.Format("KAMGroupSGModel.Add {0}", myadapter.ErrorMessage))
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
            Return "Sales.sfkamgroupsg"
        End Get
    End Property
End Class
