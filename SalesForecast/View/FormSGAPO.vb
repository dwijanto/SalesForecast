Public Class FormSGAPO

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myAPO As New APOSG


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        myAPO.Generate(Me, New APOSGEventArgs With {.startPeriod = DateTimePicker1.Value.Date})
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
    End Sub

    Private Sub loaddata()
        'Throw New NotImplementedException
    End Sub
End Class