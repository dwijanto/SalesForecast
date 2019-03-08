Imports System.Text
Imports System.Threading

Public Class FormUserAssignment
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Private myAdapter As New UserAssignmentController
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Private userId As Object
    Dim userRole As List(Of SalesForecast.Role)
    Dim roles As List(Of SalesForecast.Item)
    Public Sub New(userid)
        InitializeComponent()
        Me.userId = userid
    End Sub

    Private Sub FormUserAssignment_Load(sender As Object, e As EventArgs) Handles Me.Load
        loadData()
    End Sub

    Private Sub loadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoWork()
        'Get all roles - user's role -> assign to Listbox1
        'Get user's role -> assign to Listbox2

        Dim RBAC = New DbManager
        userRole = RBAC.getRolesByUser(userId)
        roles = RBAC.getRoles
        ProgressReport(4, "Fill Data..")
        ProgressReport(1, "Done.")


    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4
                    ListBox2.Items.Clear()
                    Dim userRoleList As New StringBuilder
                    For Each Role As Role In userRole
                        ListBox2.Items.Add(Role.name)
                        If userRoleList.Length > 0 Then
                            userRoleList.Append(",")
                        End If
                        userRoleList.Append(Role.name)
                    Next

                    ListBox1.Items.Clear()

                    Dim myroleArr = userRoleList.ToString.Split(",")


                    For Each Role As Role In roles
                        Dim addIt As Boolean = True
                        For i = 0 To myroleArr.Length - 1
                            If myroleArr(i) = Role.name Then
                                addIt = False
                            End If
                        Next

                        If addIt Then
                            ListBox1.Items.Add(Role.name)
                        End If
                    Next
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
            End Select
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim RBAC As New DbManager
        Dim myrole = ListBox1.SelectedItem
        Try
            RBAC.assign(RBAC.getRole(myrole), userId)
            ListBox1.Items.Remove(myrole)
            ListBox2.Items.Add(myrole)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim RBAC As New DbManager
        Dim myrole = ListBox2.SelectedItem
        Try
            RBAC.revoke(RBAC.getRole(myrole), userId)
            ListBox1.Items.Add(myrole)
            ListBox2.Items.Remove(myrole)
        Catch ex As Exception

        End Try
    End Sub
End Class