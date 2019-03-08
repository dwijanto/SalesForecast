Public Class FormMYALLKAM

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Dim mytemplate As New ALLKAMMY


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        mytemplate.Generate(Me, New ALLKAMEventArgs With {.startperiod = DateTimePicker1.Value.Date,
                                                          .endperiod = DateTimePicker2.Value.Date})
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


    End Sub

    Private Sub FormMLATemplateHK_Load(sender As Object, e As EventArgs) Handles Me.Load
        loaddata()
        DateTimePicker2.Value = DateTimePicker1.Value.AddMonths(12)
    End Sub

    Private Sub loaddata()
        'Throw New NotImplementedException
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        DateTimePicker2.Value = DateTimePicker1.Value.AddMonths(12)
    End Sub
End Class