

Public Class FormForecastGroupTH
    Dim KAMBS As BindingSource
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Dim mytemplate As New ForecastGroupTemplateTH
    Dim TBParamDetailController1 As New TBParamDetailController
    Dim exRate As Decimal

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim drv As DataRowView = ComboBox1.SelectedItem
        mytemplate.Generate(Me, New FGTemplateTHEventArgs With {.startPeriod = DateTimePicker1.Value.Date, .userName = drv.Item("username"), .blanktemplate = CheckBox1.Checked})
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
        Dim Identitiy As UserController = User.getIdentity
        Dim KAMController As New KAMController
        If Identitiy.isAdmin Then
            KAMController.loaddata(String.Format("where u.location = 5"))

        Else
            KAMController.loaddata(String.Format("where u.location = 5 and userid = '{0}'", Identitiy.username))
        End If
        KAMBS = KAMController.GetBindingSource
        ComboBox1.DataSource = KAMBS
        ComboBox1.DisplayMember = "fullname"
        ComboBox1.ValueMember = "username"

        'exRate = TBParamDetailController1.Model.getCurrency(country.MS, "EX-Rate")


    End Sub

    Private Sub FormMLATemplateHK_Load(sender As Object, e As EventArgs) Handles Me.Load
        loaddata()
    End Sub

    Private Sub loaddata()
        'Throw New NotImplementedException
    End Sub
End Class

Public Class FGTemplateTHEventArgs
    Inherits EventArgs
    Public Property startPeriod As Date
    Public Property userName As String
    Public Property blanktemplate As Boolean
    'Public Property ExRate As Decimal
End Class