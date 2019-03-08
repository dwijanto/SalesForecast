Imports System.Threading

Public Class FormMLATemplateHK
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    'Dim mytemplate As New MLATemplateHK
    Dim mytemplate As New MLATemplateHKSD
    Dim KAMBS As BindingSource
    Dim myuser As UserController = New UserController

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim drv As DataRowView = ComboBox1.SelectedItem
        mytemplate.Generate(Me, New MLATemplateHKEventArgs With {.startPeriod = DateTimePicker1.Value.Date, .userName = drv.Item("username"), .blanktemplate = CheckBox1.Checked})
    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Try
                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 2
                        ToolStripStatusLabel2.Text = message
                    Case 4

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        'KamAdapter.LoadData(New DataSet, "where u.location = 1;")
        Dim Identitiy As UserController = User.getIdentity
        Dim KAMController As New KAMController
        If Identitiy.isAdmin Then
            KAMController.loaddata(String.Format("where u.location = 1"))

        Else
            KAMController.loaddata(String.Format("where u.location = 1 and userid = '{0}'", Identitiy.username))
        End If
        KAMBS = KAMController.GetBindingSource
        ComboBox1.DataSource = KAMBS
        ComboBox1.DisplayMember = "fullname"
        ComboBox1.ValueMember = "username"

    End Sub

    Private Sub FormMLATemplateHK_Load(sender As Object, e As EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Private Sub loaddata()
        'Throw New NotImplementedException
    End Sub

End Class

Public Class MLATemplateHKEventArgs
    Inherits EventArgs

    Public Property startPeriod As Date
    Public Property userName As String
    Public Property blanktemplate As Boolean
End Class