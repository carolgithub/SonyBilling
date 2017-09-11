VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Begin VB.Form ReportForm 
   Caption         =   "Contacts Report"
   ClientHeight    =   7050
   ClientLeft      =   -195
   ClientTop       =   855
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ReportForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7050
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer CRViewer 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _cx             =   20770
      _cy             =   12515
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1033
   End
End
Attribute VB_Name = "ReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CRViewer_CloseButtonClicked(UseDefault As Boolean)
    Set report = Nothing
    Set crystal = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
    'YOU NEED REFERENCES TO: Microsoft ActiveX Data Objects 2.7 Library
    '                        Crystal Reports 9 ActiveX Designer Run Time Library
    'YOU NEED THE COMPONENT: Crystal Reports Viewer Control 9
    'ADD THE CRVIEWER91 COMPONENT TO YOUR FORM AND NAME IT CRViewer.
    
    '    Dim rs As ADODB.Recordset               'HOLDS ALL DATA RETURNED FROM QUERY
    Dim crystal As CRAXDRT.Application   'LOADS REPORT FROM FILE
    Dim report As CRAXDRT.report            'HOLDS REPORT
    Dim crysConn As CRAXDRT.ConnectionProperty
    
'    Dim MyApp As New CRAXDRT.Application
'    Dim MyRpt As New CRAXDRT.report
    
    Dim eomCurrentStart As Date
    Dim eomCurrentEnd As Date
        
    Dim crParamDefs As CRAXDRT.ParameterFieldDefinitions
    Dim crParamDef As CRAXDRT.ParameterFieldDefinition
    
'    CRViewer.BorderStyle = False          'MAKES REPORT FILL ENTIRE FORM
''    CRViewer.DisplayTabs = False            'THIS REPORT DOES NOT DRILL DOWN, NOT NEEDED
'    CRViewer.EnableDrilldown = False        'REPORT DOES NOT SUPPORT DRILL-DOWN
'    CRViewer.EnableRefreshButton = False    'ADO RECORDSET WILL NOT CHANGE, NOT NEEDED
    
If IntStat = 2 Then
    If RSConnect1.RecordCount <> 0 Then
    
        Set crystal = New CRAXDRT.Application         'MANAGES REPORTS
        Set report = crystal.OpenReport(App.Path & "\InvoiceRep.rpt")  'OPEN OUR REPORT
    
        report.DiscardSavedData                      'CLEARS REPORT SO WE WORK FROM RECORDSET
        report.Database.SetDataSource RSConnect1         'LINK REPORT TO RECORDSET
      
        report.DiscardSavedData             'CLEARS REPORT SO WE WORK FROM RECORDSET
        
'    ------------------------------------------------
    
''        eomCurrentStart = Format(CDate(frm_Billing.txt_DateFrom.Text), "dd/mm/yyyy 00:00:00")
''        eomCurrentEnd = Format(CDate(frm_Billing.txt_DateTo.Text), "dd/mm/yyyy 23:59:59")
'        eomCurrentStart = Format(CDate(frm_Billing.txt_DateFrom.Text), "mm/dd/yyyy 00:00:00")
'        eomCurrentEnd = Format(CDate(frm_Billing.txt_DateTo.Text), "mm/dd/yyyy 23:59:59")
'
        report.DisplayProgressDialog = False
        report.EnableParameterPrompting = False
                
'        'Fill report parameters
'        Set crParamDefs = report.ParameterFields
'        For Each crParamDef In crParamDefs
'            Select Case crParamDef.ParameterFieldName
'                Case "p1"
'                    crParamDef.AddCurrentValue eomCurrentStart
'                Case "p2"
'                    crParamDef.AddCurrentValue eomCurrentEnd
'            End Select
'        Next

'------------------------------------------
'
'        CRViewer.ReportSource = report                      'LINK VIEWER TO REPORT
        
'        Set MyRpt = MyApp.OpenReport("c:\windows\sample.rpt", 1)

        
        '    report.ParameterFields.Item(1).Value = "04/30/2004"
        '    Public WithEvents cr_Generic_Report As crysta
        '    cr_Generic_Report.SortFields(1) = "{View.Company}"
        '    cr_Generic_Report.SortFields(2) = "{View.Field2}"
        
        '    report.Areas.Item("GH1").SortDirection = crDescendingOrder
        ''   report.GroupSortFields(0) = "+{rs.Company}"
        '    report.ReportTitle = "By Work Location"

        ReportForm.WindowState = 2
'        CRViewer.ViewReport                   'SHOW REPORT
        
'        Do While CRViewer.IsBusy              'ZOOM METHOD DOES NOT WORK WHILE
'            DoEvents                          'REPORT IS LOADING, SO WE MUST PAUSE
'        Loop                                  'WHILE REPORT LOADS.
        
'        CRViewer.Zoom 100

        ReportForm.CRViewer.ReportSource = report
        ReportForm.Show
        ReportForm.CRViewer.ViewReport
        
        RSConnect1.Close                              'ALL BELOW HERE IS CLEANUP
        Set RSConnect1 = Nothing
        Set crystal = Nothing
        Set report = Nothing

    End If
    
ElseIf IntStat = 4 Then
    If RSConnect4.RecordCount <> 0 Then
        ''''''''''''''''''''''''''''''''''''''''
        'To Print the total Charges and Credits
        Set crystal = New CRAXDRT.Application         'MANAGES REPORTS
        Set report = crystal.OpenReport(App.Path & "\BillCode_Fees.rpt")  'OPEN OUR REPORT

        report.DiscardSavedData                      'CLEARS REPORT SO WE WORK FROM RECORDSET
        report.Database.SetDataSource RSConnect4         'LINK REPORT TO RECORDSET
      
        report.DiscardSavedData                  'CLEARS REPORT SO WE WORK FROM RECORDSET

        Call LastDayOfMonth
        
        report.DisplayProgressDialog = False
        report.EnableParameterPrompting = False
        
'        CRViewer.ReportSource = report              'LINK VIEWER TO REPORT
        report.ReportTitle = MEnd

        ReportForm.WindowState = 2
'        CRViewer.ViewReport                         'SHOW REPORT

'        Do While CRViewer.IsBusy              'ZOOM METHOD DOES NOT WORK WHILE
'            DoEvents                          'REPORT IS LOADING, SO WE MUST PAUSE
'        Loop                                  'WHILE REPORT LOADS.
'
'        CRViewer.Zoom 100

        ReportForm.CRViewer.ReportSource = report
        ReportForm.Show
        ReportForm.CRViewer.ViewReport

        RSConnect4.Close                              'ALL BELOW HERE IS CLEANUP
        Set RSConnect4 = Nothing
        Set crystal = Nothing
        Set report = Nothing

    End If
    
ElseIf IntStat = 3 Then
    Dim RsError As New ADODB.Recordset
    
    Set RsError = New ADODB.Recordset
    RsError.Open "select * from BillTransactionError", DBConnect, adOpenKeyset, adLockOptimistic, adCmdText
    
    If RsError.RecordCount <> 0 Then
    
        Set crystal = New CRAXDRT.Application
        Set report = crystal.OpenReport(App.Path & "\ErrorReport.rpt")
    
        report.DiscardSavedData
        report.Database.SetDataSource RsError

        report.DiscardSavedData
        report.DisplayProgressDialog = False
        report.EnableParameterPrompting = False
              
'        CRViewer.ReportSource = report
        ReportForm.WindowState = 2
'        CRViewer.ViewReport
'
'        Do While CRViewer.IsBusy
'            DoEvents                          'REPORT IS LOADING, SO WE MUST PAUSE
'        Loop                                  'WHILE REPORT LOADS.
'
'        CRViewer.Zoom 100

        ReportForm.CRViewer.ReportSource = report
        ReportForm.Show
        ReportForm.CRViewer.ViewReport
        
        RsError.Close                              'ALL BELOW HERE IS CLEANUP
        Set RsError = Nothing
        Set crystal = Nothing
        Set report = Nothing
    End If
''    RSConnect1.Close                              'ALL BELOW HERE IS CLEANUP
''    Set RSConnect1 = Nothing
End If
    
End Sub

Private Sub Form_Resize()               'MAKE SURE REPORT FILLS FORM
    CRViewer.Top = 0                    'WHEN FORM IS RESIZED
    CRViewer.Left = 0
    CRViewer.Height = ScaleHeight
    CRViewer.Width = ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

'Private Sub CrysLogon()
'    Dim LogOnInfo As CRAXDRT.lo
'    Dim strDSN As String
'    Dim strUID As String
'    Dim strPWD As String
'    '
'    ' Set up login parms
'    '
'    strDSN = "ECSFBilling"
'    strUID = "ECSFBilling"
'    strPWD = "ECSFBilling"
'
'    '
'    ' Establish ODBC connection
'    '
'    LogOnInfo.StructSize = Len(LogOnInfo)
'    LogOnInfo.ServerName = strDSN + Chr$(0)
'    LogOnInfo.DatabaseName = "ECSFBilling" + Chr$(0)
'    LogOnInfo.UserID = strUID + Chr$(0)
'    LogOnInfo.Password = strPWD + Chr$(0)
'    If PELogOnServer("PDSODBC.DLL", LogOnInfo) <> 1 Then
'        Call PrintErrorHandler("Failed to log onto database: Could not connect.")
'    End If
'
'End Sub

