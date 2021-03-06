﻿Public Class CMMFNSPTWController
    Implements IController
    Implements IToolbarAction

    Dim Model As New CMMFNSPTWModel
    Public BS As BindingSource
    Dim DS As DataSet

    Public ReadOnly Property GetTable As DataTable Implements IController.GetTable
        Get
            Return DS.Tables(Model.TableName).Copy()
        End Get
    End Property

    Public ReadOnly Property GetBindingSource As BindingSource
        Get
            Dim BS As New BindingSource
            BS.DataSource = GetTable
            BS.Sort = Model.SortField
            Return BS
        End Get
    End Property
    Public Function getRecord(ByVal cmmf As Long) As DataRowView
        Dim itemfound As Integer = BS.Find("cmmf", cmmf)
        If itemfound >= 0 Then
            BS.Position = itemfound
            Return BS.Current
        Else
            Return BS.AddNew()
        End If
    End Function

    Public Function RemoveRecord(ByVal cmmf As Long) As Boolean
        Dim itemfound As Integer = BS.Find("cmmf", cmmf)
        If itemfound >= 0 Then
            BS.RemoveAt(itemfound)
        End If
        Return True
    End Function

    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = False
        Model = New CMMFNSPTWModel
        DS = New DataSet
        If Model.LoadData(DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("cmmf")
            DS.Tables(0).PrimaryKey = pk
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
            myret = True
        End If
        Return myret

    End Function
    Public Function loaddata(ByRef criteria As String) As Boolean
        Dim myret As Boolean = False
        Model = New CMMFNSPTWModel
        DS = New DataSet
        If Model.LoadData(DS, criteria) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("cmmf")
            DS.Tables(0).PrimaryKey = pk
            BS = New BindingSource
            BS.DataSource = DS.Tables(0)
            myret = True
        End If
        Return myret

    End Function
    Public Function save() As Boolean Implements IController.save
        Dim myret As Boolean = False
        BS.EndEdit()

        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If save(mye) Then
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    'MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If

        Return myret
    End Function

    Public Function Save(ByVal mye As ContentBaseEventArgs) As Boolean Implements IToolbarAction.Save
        Dim myret As Boolean = False
        If Model.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function

    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return BS.Filter
        End Get
        Set(value As String)

        End Set
    End Property

    Public Function GetCurrentRecord() As DataRowView Implements IToolbarAction.GetCurrentRecord
        Return BS.Current
    End Function

    Public Function GetNewRecord() As DataRowView Implements IToolbarAction.GetNewRecord
        Return BS.AddNew
    End Function

    Public Sub RemoveAt(value As Integer) Implements IToolbarAction.RemoveAt
        BS.RemoveAt(value)
    End Sub
End Class
