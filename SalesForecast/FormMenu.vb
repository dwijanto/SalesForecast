Public Class FormMenu
    Dim myuser As UserController = New UserController
    Dim HKReportProperty1 As HKReportProperty
    'Dim HKGroupReportProperty1 As HKGroupReportProperty
    Dim TWReportProperty1 As TWReportProperty
    Dim MSReportProperty1 As MSReportProperty
    Dim SGReportProperty1 As SGReportProperty
    Dim THReportProperty1 As THReportProperty
    Dim username As String

    Private Sub KAMToolStripMenuItem1_Click(sender As Object, e As EventArgs)

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        'HK MLA Template
        HKReportProperty1 = HKReportProperty.getInstance

        HKReportProperty1.ColumnStartKey = 6
        HKReportProperty1.ColumnStartData = 10
        HKReportProperty1.RowStartDataI = 19 '14    :: 18 For HKMLA with New Sales Deduction  
        HKReportProperty1.RowStartdData = "A19" 'A18 :: A19 -> Added WMF HKReportProperty1.RowStartDataI.ToString      

        'HK Forecast Group Template
        HKReportProperty1.FGRowStartDataI = 14
        HKReportProperty1.FGRowStartData = "A14" '& HKReportProperty1.FGRowStartDataI.ToString
        HKReportProperty1.FGColumnStartData = 9 '8 :: 9-> added Brand Column 2018-04-11
        HKReportProperty1.FGColumnStartKey = 1



        'HKGroupReportProperty1 = HKGroupReportProperty.getInstance
        'HKGroupReportProperty1.ColumnStartKey = 6
        'HKGroupReportProperty1.ColumnStartData = 10
        'HKGroupReportProperty1.RowStartDataI = 14 '    :: 18 For HKMLA with New Sales Deduction
        'HKGroupReportProperty1.FGColumnStartData = 8
        'HKGroupReportProperty1.FGRowStartDataI = 14
        'HKGroupReportProperty1.FGColumnStartKey = 1
        'HKGroupReportProperty1.RowStartdData = "A" & HKGroupReportProperty1.RowStartDataI.ToString

        TWReportProperty1 = TWReportProperty.getInstance

        TWReportProperty1.ColumnStartKey = 2
        TWReportProperty1.ColumnStartData = 13
        TWReportProperty1.RowStartDataI = 11
        TWReportProperty1.RowStartdData = "A" & TWReportProperty1.RowStartDataI.ToString

        MSReportProperty1 = MSReportProperty.getInstance

        MSReportProperty1.ColumnStartKey = 2
        MSReportProperty1.ColumnStartData = 12
        MSReportProperty1.RowStartDataI = 11 '9
        MSReportProperty1.RowStartdData = "A" & MSReportProperty1.RowStartDataI.ToString

        SGReportProperty1 = SGReportProperty.getInstance

        SGReportProperty1.ColumnStartKey = 2
        SGReportProperty1.ColumnStartData = 10
        SGReportProperty1.RowStartDataI = 10 '9
        SGReportProperty1.RowStartdData = "A" & SGReportProperty1.RowStartDataI.ToString

        THReportProperty1 = THReportProperty.getInstance
        THReportProperty1.ColumnStartKey = 1
        THReportProperty1.ColumnStartData = 10
        THReportProperty1.RowStartDataI = 6
        THReportProperty1.RowStartdData = "A" & THReportProperty1.RowStartDataI.ToString

    End Sub
    Private Sub displayMenuBar()
        Dim identity As UserController = User.getIdentity
        MasterToolStripMenuItem.Visible = User.can("View Master")
        CMMFToolStripMenuItem.Visible = User.can("View CMMF")
        CMMFHKToolStripMenuItem1.Visible = User.can("View CMMF HK")
        CMMFTWToolStripMenuItem.Visible = User.can("View CMMF TW")
        CMMFMYToolStripMenuItem.Visible = User.can("View CMMF MY")
        CMMFSGToolStripMenuItem.Visible = User.can("View CMMF SG")
        CMMFTHToolStripMenuItem.Visible = User.can("View CMMF TH")

        ToolsToolStripMenuItem.Visible = User.can("View Tools")
        ActionToolStripMenuItem.Visible = User.can("View Action")
        ExtractTemplateToolStripMenuItem.Visible = User.can("View MLA HK")
        ExtractForecastGroupTemplateToolStripMenuItem.Visible = User.can("View Forecast Group HK")
        ImportTemplateToolStripMenuItem.Visible = User.can("View Import HK")
        ExportAPOToolStripMenuItem.Visible = User.can("View APO HK")
        ExportAPOHKPriceToolStripMenuItem.Visible = User.can("View APO HK")
        ExportForecastGroupTWToolStripMenuItem.Visible = User.can("View Forecast Group TW")
        ImportSalesForecastTWToolStripMenuItem.Visible = User.can("View Import TW")
        ExportAPOTWToolStripMenuItem.Visible = User.can("View APO TW")
        KAMToolStripMenuItem.Visible = User.can("View KAM")
        KAMTargetToolStripMenuItem.Visible = User.can("View KAM Target HK")
        KAMTargetTWToolStripMenuItem.Visible = User.can("View KAM Target TW")
        RToolStripMenuItem.Visible = User.can("View Report")
        HKALLKAMToolStripMenuItem.Visible = User.can("View Report ALL HK")
        TWALLKAMToolStripMenuItem.Visible = User.can("View Report ALL TW")
        MYALLKAMToolStripMenuItem.Visible = User.can("View Report ALL MY")
        SGALLKAMToolStripMenuItem.Visible = User.can("View Report ALL SG")
        THALLKAMToolStripMenuItem.Visible = User.can("View Report ALL TH")

        MLAToolStripMenuItem.Visible = User.can("View MLA HK")
        MLASDHKToolStripMenuItem.Visible = User.can("View KAM Target HK") 'Supervisor HK

        SalesDeductionToolStripMenuItem.Visible = User.can("View Sales Deduction")
        GrossSalesTWToolStripMenuItem.Visible = User.can("View Gross Sales TW")
        GrossSalesBudgetTWToolStripMenuItem.Visible = User.can("View Gross Sales Target TW")
        ExRateToolStripMenuItem.Visible = User.can("View Ex Rate TW")
        RawDataToolStripMenuItem.Visible = User.can("View Rawdata HK")
        ExportMSTemplateToolStripMenuItem.Visible = User.can("View MS Template")
        ImportSalesForecastMSToolStripMenuItem.Visible = User.can("View Import MS")
        ImportFromRAWDATAMYToolStripMenuItem.Visible = User.can("View Import RAWDATA MS")
        ExportAPOMSToolStripMenuItem.Visible = User.can("View APO MS")
        ExportSGTemplateToolStripMenuItem.Visible = User.can("View SG Template")
        ImportSalesForecastSGToolStripMenuItem.Visible = User.can("View Import SG")
        ExportAPOSGToolStripMenuItem.Visible = User.can("View APO SG")
        MasterKAMToolStripMenuItem.Visible = User.can("View Master KAM")
        KAMBudgetMSToolStripMenuItem.Visible = User.can("View Budget KAM MS")
        HKALLKAMTARGETToolStripMenuItem.Visible = User.can("View Report ALL HK TARGET")



        ExportTHTemplateToolStripMenuItem.Visible = User.can("View TH Template")
        ImportSalesForecastTHToolStripMenuItem.Visible = User.can("View Import TH")
        ExportAPOTHToolStripMenuItem.Visible = User.can("View APO TH")
    End Sub

    Private Sub disableMenuBar()
        'AdminToolStripMenuItem.Visible = False
        'QueryToolStripMenuItem.Visible = False
        MessageBox.Show(String.Format("You're not authorized to use this function. If not please contact Admin.User id: {0}.", Environment.UserDomainName & "\" & Environment.UserName))
        Me.Close()
    End Sub

    Private Sub FormMenu_Load(sender As Object, e As EventArgs) Handles Me.Load
        username = Environment.UserDomainName & "\" & Environment.UserName
        Dim mydata = myuser.findByUserName(username.ToLower)
        If mydata.Tables(0).rows.count > 0 Then
            Dim identity = myuser.findIdentity(mydata.Tables(0).rows(0).item("id"))
            User.setIdentity(identity)
            User.login(identity)
            User.IdentityClass = myuser
            Try
                loglogin(username)
            Catch ex As Exception

            End Try
            displayMenuBar()
            Me.Text = getMenudesc()
        Else
            disableMenuBar()
        End If
    End Sub
    Private Function getMenudesc() As String
        Dim myAdapter = PostgreSQLDbAdapter.getInstance
        Return "App.Version: " & My.Application.Info.Version.ToString & " :: Server: " & myAdapter.Host & ", Database: " & myAdapter.DB & ", Userid: " & username
    End Function

    Private Sub loglogin(ByVal userid As String)
        Dim applicationname As String = "Sales Forecast"
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim computername As String = My.Computer.Name
        Dim time_stamp As DateTime = Now
        myuser.loglogin(applicationname, userid, username, computername, time_stamp)
    End Sub

    Private Sub MasterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MasterToolStripMenuItem.Click

    End Sub

    Private Sub ToolsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToolsToolStripMenuItem.Click

    End Sub

    Private Sub ExtractTemplateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExtractTemplateToolStripMenuItem.Click
        Dim myform = New FormMLATemplateHK
        myform.ShowDialog()
    End Sub

    Private Sub CMMFToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles CMMFHKToolStripMenuItem1.Click
        Dim myform = New FormCMMF
        myform.ShowDialog()
    End Sub

    Private Sub ExtractForecastGroupTemplateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExtractForecastGroupTemplateToolStripMenuItem.Click
        Dim myform = New FormForecastGroupHK
        myform.showdialog()
    End Sub

    Private Sub ImportTemplateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportTemplateToolStripMenuItem.Click
        Dim myform = New FormHKImport
        myform.ShowDialog()
    End Sub

    Private Sub KAMTargetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KAMTargetToolStripMenuItem.Click
        Dim myform = New FormKAMTarget
        myform.ShowDialog()
    End Sub

    Private Sub ExportAPOToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportAPOToolStripMenuItem.Click
        Dim myform = New FormHKAPO
        myform.showdialog()
    End Sub

    Private Sub CMMFMLANSPToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub UserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UserToolStripMenuItem.Click
        Dim myform = New FormUser
        myform.ShowDialog()
    End Sub

    Private Sub TestPostgreSQLConnectionToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub


    Private Sub ExportForecastGroupTWToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportForecastGroupTWToolStripMenuItem.Click
        Dim myform = New FormForecastGroupTW
        myform.ShowDialog()
    End Sub

    Private Sub HKALLKAMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HKALLKAMToolStripMenuItem.Click
        Dim myform = New FormHKALLKAM
        myform.ShowDialog()
    End Sub

    Private Sub ImportSalesForecastTWToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportSalesForecastTWToolStripMenuItem.Click
        Dim myform = New FormTWImport
        myform.ShowDialog()
    End Sub

    Private Sub ExportAPOTWToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportAPOTWToolStripMenuItem.Click
        Dim myform = New FormTWAPO
        myform.ShowDialog()
    End Sub

    Private Sub CMMFTWToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CMMFTWToolStripMenuItem.Click
        Dim myform = New FormCMMFTW
        myform.ShowDialog()
    End Sub

    Private Sub KAMTargetTWToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KAMTargetTWToolStripMenuItem.Click
        Dim myform = New FormKAMTargetTW
        myform.ShowDialog()
    End Sub

    Private Sub TWALLKAMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TWALLKAMToolStripMenuItem.Click
        Dim myform = New FormTWALLKAM
        myform.ShowDialog()
    End Sub

    Private Sub SalesDeductionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalesDeductionToolStripMenuItem.Click
        Dim myform = New FormSalesDeductionTW
        myform.ShowDialog()
    End Sub

    Private Sub MLAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MLAToolStripMenuItem.Click
        Dim myform = New FormMLA
        myform.ShowDialog()
    End Sub

    Private Sub UserGuideToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UserGuideToolStripMenuItem.Click
        Dim p As New System.Diagnostics.Process
        'p.StartInfo.FileName = "\\172.22.10.44\SharedFolder\PriceCMMF\New\template\Supplier Management Task User Guide-Admin.pdf"
        p.StartInfo.FileName = Application.StartupPath & "\help\Sales Forecast User Guide.pdf"
        p.Start()
    End Sub

    Private Sub GrossSalesTWToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GrossSalesTWToolStripMenuItem.Click
        Dim myform = New FormGrossSalesTW
        myform.ShowDialog()
    End Sub

    Private Sub GrossSalesBudgetTWToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GrossSalesBudgetTWToolStripMenuItem.Click
        Dim myform = New FormGrossSalesBudgetorTarget
        myform.ShowDialog()
    End Sub

    Private Sub ExRateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExRateToolStripMenuItem.Click
        Dim myform = New DialogTBParamDetail(country.TW)
        myform.ShowDialog()

    End Sub

    Private Sub KAMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KAMToolStripMenuItem.Click

    End Sub

    Private Sub RawDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RawDataToolStripMenuItem.Click
        Dim myform As New FormTousRawData
        myform.ShowDialog()
    End Sub

    Private Sub ExportMSTemplateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportMSTemplateToolStripMenuItem.Click
        Dim myform = New FormForecastGroupMS
        myform.ShowDialog()
    End Sub

    Private Sub ImportSalesForecastMSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportSalesForecastMSToolStripMenuItem.Click
        Dim myform = New FormMSImport
        myform.ShowDialog()
    End Sub

    Private Sub ExporAPOMSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportAPOMSToolStripMenuItem.Click
        Dim myform = New FormMYAPO
        myform.ShowDialog()
    End Sub

    Private Sub ExportSGTemplateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportSGTemplateToolStripMenuItem.Click
        Dim myform = New FormForecastGroupSG
        myform.ShowDialog()
    End Sub

    Private Sub ImportSalesForecastSGToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportSalesForecastSGToolStripMenuItem.Click
        Dim myform = New FormSGImport
        myform.ShowDialog()
    End Sub

    Private Sub ExportAPOSGToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportAPOSGToolStripMenuItem.Click
        Dim myform = New FormSGAPO
        myform.ShowDialog()
    End Sub

    Private Sub MYALLKAMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MYALLKAMToolStripMenuItem.Click
        Dim myform = New FormMYALLKAM
        myform.ShowDialog()
    End Sub

    Private Sub SGALLKAMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SGALLKAMToolStripMenuItem.Click
        Dim myform = New FormSGALLKAM
        myform.ShowDialog()
    End Sub

    Private Sub CMMFMYToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CMMFMYToolStripMenuItem.Click
        Dim myform = New FormCMMFMY
        myform.ShowDialog()
    End Sub

    Private Sub CMMFSGToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CMMFSGToolStripMenuItem.Click
        Dim myform = New FormCMMFSG
        myform.ShowDialog()
    End Sub

    Private Sub MasterKAMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MasterKAMToolStripMenuItem.Click
        Dim myform = New FormKAM
        myform.ShowDialog()
    End Sub

    Private Sub KAMBudgetMSToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KAMBudgetMSToolStripMenuItem.Click
        Dim myform = New FormMYBudget
        myform.showdialog()
    End Sub

    Private Sub CMMFToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CMMFToolStripMenuItem.Click
        
    End Sub

    Private Sub HKALLKAMTARGETToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HKALLKAMTARGETToolStripMenuItem.Click
        Dim myform = New FormHKALLKAMTarget
        myform.ShowDialog()
    End Sub

    Private Sub ExportTHTemplateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportTHTemplateToolStripMenuItem.Click
        Dim myform = New FormForecastGroupTH
        myform.ShowDialog()
    End Sub

    Private Sub ImportSalesForecastTHToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportSalesForecastTHToolStripMenuItem.Click
        Dim myform = New FormTHImport
        myform.ShowDialog()
    End Sub

    Private Sub ExportAPOTHToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportAPOTHToolStripMenuItem.Click
        Dim myform = New FormTHAPO
        myform.ShowDialog()
    End Sub

    Private Sub CMMFTHToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CMMFTHToolStripMenuItem.Click
        Dim myform = New FormCMMFTH
        myform.ShowDialog()
    End Sub

    Private Sub THALLKAMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles THALLKAMToolStripMenuItem.Click
        Dim myform = New FormTHALLKAM
        myform.ShowDialog()
    End Sub

    Private Sub ImportFromRAWDATAMYToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportFromRAWDATAMYToolStripMenuItem.Click
        Dim myform = New FormImportRAWDATAMS
        myform.ShowDialog()
    End Sub

    Private Sub FamilyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FamilyToolStripMenuItem.Click
        Dim myform = New FormFamily
        myform.Show()
    End Sub

    Private Sub ExportAPOHKPriceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportAPOHKPriceToolStripMenuItem.Click
        Dim myform = New FormHKAPOPrice
        myform.ShowDialog()
    End Sub

    Private Sub MLASDHKToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MLASDHKToolStripMenuItem.Click
        Dim myform = New FormMLACardNameSD
        myform.ShowDialog()
    End Sub

   


End Class
