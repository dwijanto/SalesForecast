Imports System.Threading

Public Class FormHKAPOPrice

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myAPO As New APOHK
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim MlaNSPList As List(Of String)
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not myThread.IsAlive Then
            myAPO.GenerateAPOPrice(Me, New APOHKEventArgs With {.startPeriod = DateTimePicker1.Value.Date,
                                                               .MLA = ComboBox1.SelectedItem})
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If

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
                        If MlaNSPList.Count > 0 Then
                            For Each mla In MlaNSPList
                                ComboBox1.Items.Add(mla)
                            Next
                            ComboBox1.SelectedIndex = 0
                        End If
                        

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
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If


    End Sub

    Private Sub DoWork()
        ProgressReport(6, "Marque")
        Try
            Dim MlaNspModel1 As New MlaNspModel
            MlaNSPList = MlaNspModel1.GetMLAList
            ProgressReport(4, "")
        Catch ex As Exception
        Finally
            ProgressReport(5, "Marque")
        End Try

    End Sub

End Class