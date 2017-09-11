VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frm_Billing_old 
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   315
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin ComCtl2.Animation Animation1 
      Height          =   345
      Left            =   3000
      TabIndex        =   22
      Top             =   8160
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      _Version        =   327681
      BackColor       =   12632256
      FullWidth       =   65
      FullHeight      =   23
   End
   Begin VB.CommandButton cmdExec 
      Height          =   495
      Left            =   11040
      Picture         =   "BillingReport_old.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Execute"
      Top             =   240
      Width           =   735
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   8160
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   22806529
      CurrentDate     =   38134
   End
   Begin VB.Frame frame_Print 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   2760
      TabIndex        =   17
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton Opt_MInv 
         Caption         =   "Monthly Invoice Chart"
         Height          =   195
         Left            =   0
         TabIndex        =   19
         Top             =   120
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton Opt_SInv 
         Caption         =   "Individual Invoice"
         Height          =   195
         Left            =   0
         TabIndex        =   18
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame frame_search 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   5160
      TabIndex        =   5
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton cmd_to 
         Caption         =   "..."
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton cmd_from 
         Caption         =   "..."
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   120
         Width           =   375
      End
      Begin VB.TextBox txt_DateFrom 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   345
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   120
         Width           =   1335
      End
      Begin VB.TextBox txt_DateTo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   345
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "(dd/mm/yyyy)"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label13 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "(dd/mm/yyyy)"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   600
         TabIndex        =   11
         Top             =   1305
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "(dd/mm/yyyy)"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmd_close 
      Height          =   495
      Left            =   11040
      Picture         =   "BillingReport_old.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Exit"
      Top             =   6720
      Width           =   720
   End
   Begin VB.CommandButton cmdrep 
      Height          =   495
      Left            =   10200
      Picture         =   "BillingReport_old.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Export to Excel"
      Top             =   6720
      Width           =   720
   End
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   11775
      Begin MSDataGridLib.DataGrid DataGrid_Billing 
         Bindings        =   "BillingReport_old.frx":075E
         Height          =   5325
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   9393
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         ForeColor       =   16777215
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txt_ClientCode 
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   8
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   "ecmsfilereference"
      Password        =   "ecms"
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   7230
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5116
            MinWidth        =   5116
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10584
            MinWidth        =   10584
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   3000
            MinWidth        =   3000
            TextSave        =   "01/06/2004"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2700
            MinWidth        =   2700
            TextSave        =   "18:48"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl_fileref 
      Caption         =   "Client Code"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frm_Billing_old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Refobj As New five_char
Dim xlApp As Excel.Application
Dim xlwb As Excel.Workbook
Dim file_obj As New Scripting.FileSystemObject

Private Sub cmdExec_Click()
If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
    If CDate(txt_DateFrom.Text) < CDate(txt_DateTo.Text) Then
        animationstart
        Call CallProc
        Call GetData
        Call data_refresh
        animationend
        Exit Sub
    End If
End If
End Sub

'Dim blnStat As Boolean
'Dim blnRecStat As Boolean
'Dim strdel As String
'Dim strins As String
'Dim StrDeptVal As String
'Dim strFileref_LF As String
'
'Private Sub cmd_close_Click()
'On Error GoTo Errclose
'Me.MousePointer = 1
'Screen.MousePointer = vbArrow
'
''''Adodc1.Recordset.Close
''''Set Adodc1.Recordset = Nothing
'''fm_nxtcode.dbc_ho.Text = current_country
'
'Unload Me
'
'Exit Sub
'Errclose:
'    MsgBox Err.Description & " #" & Err.Number
'End Sub
'
'Private Sub cmd_Other_Click()
'frm_OtherRef.Show
'End Sub
'
'Private Sub cmdCrysrep_Click()
'IntRepNo = 1
'ReportForm.Show vbModal
'End Sub
'
'Private Sub cmdGoback_Click()
'Call data_refresh
'cmdsearch.Visible = True
'txt_fileref.Visible = False
'lbl_fileref.Visible = False
'cmdGoback.Visible = False
'End Sub
'
Private Sub cmdrep_Click()
On Error GoTo trap

Dim strSQL As String
Dim xlApp As Object
Dim xlwb As Object
Dim xlWs As Object
    
Dim recArray As Variant
Dim fldCount As Integer
Dim recCount As Long
Dim iCol As Integer
Dim iRow As Integer
    
Dim ExFile As String
Dim strMName As String
Dim strYear As String
    
If RSConnect1.RecordCount > 0 Then
        
    If IntStat = 1 Then
        
        strMName = MonthName(Month(CDate(txt_DateFrom.Text)))
        strYear = Year(CDate(txt_DateFrom.Text))
        
        ExFile = App.Path & "\MonthlyInvoiceChart.xls"
        
        If file_obj.FileExists(App.Path & "\MonthlyInvoice_" & strMName & strYear & ".xls ") Then
            Kill (App.Path & "\MonthlyInvoice_" & strMName & strYear & ".xls ")
        End If
        
        'Open the MonthlyInvoice Template and save as MonthYear.xls
        'Create an instance of Excel and add a workbook
        Set xlApp = CreateObject("Excel.Application")
        'Open the worksheet
        Set xlwb = xlApp.Workbooks.Open(ExFile)
        Set xlWs = xlwb.Worksheets("Invoice")

        xlApp.DisplayAlerts = False
        xlApp.Visible = False
'        MsgBox (App.Path & "\MonthlyInvoice_" & strMName & " " & strYear & ".xls")
        
        xlwb.SaveAs App.Path & "\MonthlyInvoice_" & strMName & strYear & ".xls "
        
        xlwb.Close
        xlApp.Quit
        
        Set xlwb = Nothing
        Set xlApp = Nothing
        
        'Open the MonthlyInvoice saved in Year and Month Name
        'Create an instance of Excel and add a workbook
        Set xlApp = CreateObject("Excel.Application")
        'Open the worksheet
        Set xlwb = xlApp.Workbooks.Open(App.Path & "\MonthlyInvoice_" & strMName & strYear & ".xls")
        Set xlWs = xlwb.Worksheets("Invoice")

        xlApp.DisplayAlerts = False
        
        ' Display Excel and give user control of Excel's lifetime
        xlApp.Visible = True
        xlApp.UserControl = True
    
        With xlWs.Cells(1, 3)
            .Value = strYear
            .Font.Bold = True
            .Font.ColorIndex = 3
        End With
        
        With xlWs.Cells(1, 4)
            .Value = strMName
            .Font.Bold = True
            .Font.ColorIndex = 3
        End With
        
'        xlWs.Cells(1).For
        
        ' Copy field names to the first row of the worksheet

'       If IntSel = 1 Then
            fldCount = RSConnect1.Fields.Count
'       ElseIf IntSel = 2 Then
'           fldCount = RSConnect2.Fields.Count
'       ElseIf IntSel = 3 Then
'           fldCount = RSConnect3.Fields.Count
'       End If
    
        For iCol = 1 To fldCount
'           If IntSel = 1 Then
                xlWs.Cells(6, iCol).Value = RSConnect1.Fields(iCol - 1).Name
'           ElseIf IntSel = 2 Then
'               xlWs.Cells(1, iCol).Value = RSConnect2.Fields(iCol - 1).Name
'           ElseIf IntSel = 3 Then
'               xlWs.Cells(1, iCol).Value = RSConnect3.Fields(iCol - 1).Name
'           End If
            
        Next
        
    '   Check version of Excel
        If Val(Mid(xlApp.Version, 1, InStr(1, xlApp.Version, ".") - 1)) > 8 Then
            'EXCEL 2000 or 2002: Use CopyFromRecordset
         
            ' Copy the recordset to the worksheet, starting in cell A2
'           If IntSel = 1 Then
                xlWs.Cells(6, 1).CopyFromRecordset RSConnect1
'           ElseIf IntSel = 2 Then
'               xlWs.Cells(2, 1).CopyFromRecordset RSConnect2
'           ElseIf IntSel = 3 Then
'               xlWs.Cells(2, 1).CopyFromRecordset RSConnect3
'           End If
        
            'Note: CopyFromRecordset will fail if the recordset
            'contains an OLE object field or array data such
            'as hierarchical recordsets
        
        Else
            ' Copy recordset to an array
'           recArray = DE1.rsCommand1.GetRows
'           If IntSel = 1 Then
                recArray = RSConnect1.GetRows
'           ElseIf IntSel = 2 Then
'               recArray = RSConnect2.GetRows
'           ElseIf IntSel = 3 Then
'               recArray = RSConnect3.GetRows
'           End If
        
            'Note: GetRows returns a 0-based array where the first
            'dimension contains fields and the second dimension
            'contains records. We will transpose this array so that
            'the first dimension contains records, allowing the
            'data to appears properly when copied to Excel
        
            'Determine number of records

            recCount = UBound(recArray, 2) + 1 '+ 1 since 0-based array

        '   Check the array for contents that are not valid when
        '   copying the array to an Excel worksheet
            For iCol = 0 To fldCount - 1
                For iRow = 0 To recCount - 1
                    ' Take care of Date fields
                    If IsDate(recArray(iCol, iRow)) Then
                        recArray(iCol, iRow) = Format(recArray(iCol, iRow))
                        ' Take care of OLE object fields or array fields
                    ElseIf IsArray(recArray(iCol, iRow)) Then
                        recArray(iCol, iRow) = "Array Field"
                    End If
                Next iRow 'next record
            Next iCol 'next field
            
        '   Transpose and Copy the array to the worksheet,
        '   starting in cell A2
'           xlWs.Cells(22, 1).Resize(recCount, fldCount).Value = _
'           TransposeDim(recArray)
            xlwb.Cells(6, 1).Resize(recCount, fldCount).Value = _
            TransposeDim(recArray)
        End If
        Dim i, j As Integer
        For i = 6 To 5 + RSConnect1.RecordCount
            For j = 1 To 19
                With xlWs.Cells(i, j)
'                    .Font.Bold = True
'                    .Font.ColorIndex = 3
                    .Borders(xlTop).LineStyle = xlSingle
                    .Borders(xlBottom).LineStyle = xlSingle
                    .Borders(xlRight).LineStyle = xlSingle
                    .Borders(xlLeft).LineStyle = xlSingle
                End With
            Next
        Next

    '   Auto-fit the column widths and row heights
'        xlApp.Selection.CurrentRegion.Columns.AutoFit
        xlApp.Selection.CurrentRegion.Rows.AutoFit

'        xlwb.Close
'        xlApp.Quit

    '   Release Excel references
'       Set xlWs = Nothing
        xlwb.SaveAs App.Path & "\MonthlyInvoice_" & strMName & strYear & ".xls "
        Set xlwb = Nothing
        Set xlApp = Nothing
        
    Else
        ReportForm.Show
    End If
End If
    

Exit Sub

trap:
    MsgBox Err.Description & " #" & Err.Number

End Sub

Function TransposeDim(v As Variant) As Variant
' Custom Function to Transpose a 0-based array (v)

    Dim X As Long, Y As Long, Xupper As Long, Yupper As Long
    Dim tempArray As Variant

    Xupper = UBound(v, 2)
    Yupper = UBound(v, 1)

    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = v(Y, X)
        Next Y
    Next X

    TransposeDim = tempArray

End Function


Private Sub cmdsearch_Click()
''''fileref_sf = refconvert.ShortForm
txt_fileref.Text = ""
txt_fileref.Visible = True
lbl_fileref.Visible = True
cmdsearch.Visible = False
End Sub

'Private Sub DataGrid_warehouse_AfterColEdit(ByVal ColIndex As Integer)
'On Error GoTo editerr
'Dim StrUpdate As String
'
'
'Adodc1.Recordset.Update
'
'Exit Sub
'editerr:
'    MsgBox Err.Description
'End Sub

'Private Sub DataGrid_warehouse_Click()
'On Error GoTo Errclick
''''MsgBox DataGrid_warehouse.Row
''''MsgBox DataGrid_warehouse.Col
'Dim intI As Integer
'If blnRecStat = True Then
'    If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
'        StatusBar1.Panels.Item(1).Text = "Records:  " & (Adodc1.Recordset.AbsolutePosition) & "/" & (Adodc1.Recordset.RecordCount)
'    End If
'End If
'
'Exit Sub
'Errclick:
' MsgBox Err.Description & " #" & Err.Number
'End Sub
'
'Private Sub DataGrid_warehouse_ColEdit(ByVal ColIndex As Integer)
'On Error GoTo ErrColClick
'
'Dim intI As Integer
'If blnRecStat = True Then
'    If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
'        StatusBar1.Panels.Item(1).Text = "Records:  " & (Adodc1.Recordset.AbsolutePosition) & "/" & (Adodc1.Recordset.RecordCount)
'    End If
'End If
'
'Exit Sub
'ErrColClick:
' MsgBox Err.Description & " #" & Err.Number
'End Sub
'
'Private Sub DataGrid_warehouse_HeadClick(ByVal ColIndex As Integer)
'On Error GoTo ErrHeadClick
'
'Dim sortField As String
'Dim sortString As String
'
'sortField = DataGrid_warehouse.Columns(ColIndex).Caption
'If InStr(1, UCase(sortField), "FUNCTION") = 0 Then
'    If InStr(Adodc1.Recordset.Sort, "Asc") Then
'        sortString = sortField & " Desc"
'    Else
'        sortString = sortField & " Asc"
'    End If
'    Adodc1.Recordset.Sort = sortString
'End If
'
'Exit Sub
'ErrHeadClick:
' MsgBox Err.Description & " #" & Err.Number
'
'End Sub

'Private Sub DataGrid_warehouse_LostFocus()
''If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
'    Adodc1.Refresh
''End If
'End Sub

'Private Sub Form_Load()
'On Error GoTo ErrLoad
'
'Me.MousePointer = 1
''Height = Screen.Height * 0.9  ' Set height of form.
'Left = (Screen.Width - Width) / 2   ' Center form horizontally.
'Top = (Screen.Height - Height) / 2  ' Center form vertically.
'txt_filingyr.Enabled = True
'UpDownyr.Enabled = True
'blnStat = False
'StatusBar1.Panels.Item(2).Text = ""
'StatusBar1.Panels.Item(1).Text = ""
'blnRecStat = False
'StrDeptVal = "PAT"
'Adodc1.LockType = adLockOptimistic
'Adodc1.CursorType = adOpenKeyset
'Adodc1.CommandType = adCmdUnknown
'''Adodc1.ConnectionString = "dsn=ecmsfref"
'Adodc1.ConnectionString = "DSN=ecmsfref;UID=ecmsfilereference;PWD=ecms;"
'
'DataGrid_warehouse.AllowUpdate = False
'DataGrid_warehouse.AllowAddNew = False
'DataGrid_warehouse.AllowDelete = False
'
'UpDownyr.Value = Format(Date, "yyyy")
'
'StatusBar1.Visible = True
'blnStat = True
'
'If UCase(user_country) = "SG" Then
'    cmd_Other.Enabled = True
'Else
'    cmd_Other.Enabled = False
'End If
'
'Call data_refresh
'
'DataGrid_warehouse.Visible = True
'
'Me.MousePointer = 1
'StatusBar1.Panels.Item(2).Text = ""
'
'Exit Sub
'ErrLoad:
' MsgBox Err.Description & " #" & Err.Number
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
''If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
'Me.MousePointer = 1
'Screen.MousePointer = vbArrow
''''Adodc1.Refresh
''End If
'End Sub

'''Private Sub opt_CM_Click()
'''StrDeptVal = "CASE_MANAGEMENT"
'''Call data_refresh
'''    If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
'''        DataGrid_warehouse.Refresh
'''        StatusBar1.Panels.Item(1).Text = "Records: " & Adodc1.Recordset.RecordCount
'''    Else
'''        StatusBar1.Panels.Item(1).Text = "Records: 0"
'''    End If
'''    Exit Sub
'''End Sub
'
'Private Sub optmth_Click()
'UpDownyr.Enabled = False
'txt_filingyr.Text = ""
'UpDownmth.Enabled = True
'UpDownyrmth.Enabled = True
'UpDownmth.Value = Format(Date, "mm")
'UpDownyrmth.Value = Format(Date, "yyyy")
'End Sub
'
'Private Sub optyr_Click()
'txt_filingyr.Enabled = True
'UpDownyr.Enabled = True
'UpDownmth.Enabled = False
'UpDownyrmth.Enabled = False
'txt_filingmth.Text = ""
'txt_filingyrmth.Text = ""
'UpDownyr.Value = Format(Date, "yyyy")
'End Sub
'
'Private Sub txt_fileref_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
'    If Trim(txt_fileref.Text) <> Empty And Trim(txt_fileref.Text) <> "" Then
'        refconvert.Text = txt_fileref.Text
'        strFileref_LF = refconvert.LongForm
'        Call selsearch
'        If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
''            refconvert.Text = txt_fileref.Text
''            strFileref_LF = refconvert.LongForm
''            Call selsearch
'            cmdGoback.Visible = True
'        Else
'            cmdGoback.Visible = True
'            cmdsearch.Visible = True
'            lbl_fileref.Visible = False
'            txt_fileref.Visible = False
'            MsgBox "Record: " & txt_fileref.Text & " does not exist, Please verify", vbInformation, "Verify Grant No"
'        End If
'    End If
'End If
'
'End Sub
'
'Private Sub UpDownmth_Change()
'On Error GoTo ErrUpDn
'
'txt_filingmth.Text = UpDownmth.Value
'If Trim(txt_filingyrmth.Text) <> "" And txt_filingmth.Text <> "" Then
'    Call data_refresh
'    If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
'        DataGrid_warehouse.Refresh
'        StatusBar1.Panels.Item(1).Text = "Records: " & Adodc1.Recordset.RecordCount
'    Else
'        StatusBar1.Panels.Item(1).Text = "Records: 0"
'    End If
'    Exit Sub
'End If
'
'Exit Sub
'ErrUpDn:
' MsgBox Err.Description & " #" & Err.Number
'End Sub
'
'Private Sub UpDownyr_Change()
'On Error GoTo ErrUpDnYr
'
'txt_filingyr.Text = UpDownyr.Value
'If Trim(txt_filingyr.Text) <> "" Then
'    Call data_refresh
'    If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
'        DataGrid_warehouse.Refresh
'        StatusBar1.Panels.Item(1).Text = "Records: " & Adodc1.Recordset.RecordCount
'    Else
'        StatusBar1.Panels.Item(1).Text = "Records: 0"
'    End If
'    Exit Sub
'End If
'
'Exit Sub
'ErrUpDnYr:
' MsgBox Err.Description & " #" & Err.Number
'End Sub
'
'Private Sub UpDownyr_DownClick()
'On Error GoTo ErrUpDnClick
'
'txt_filingyr.Text = UpDownyr.Value
'If Trim(txt_filingyr.Text) <> "" Then
'    Call data_refresh
'    If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
'        DataGrid_warehouse.Refresh
'        StatusBar1.Panels.Item(1).Text = "Records: " & Adodc1.Recordset.RecordCount
'    Else
'        StatusBar1.Panels.Item(1).Text = "Records: 0"
'    End If
'    Exit Sub
'End If
'
'Exit Sub
'ErrUpDnClick:
' MsgBox Err.Description & " #" & Err.Number
'End Sub
'
'Private Sub UpDownyr_UpClick()
'If Trim(txt_filingyr.Text) <> "" Then
'''    Call Proc_Exec
'End If
'End Sub
'
'Private Sub data_refresh()
'On Error GoTo DataRefresh
'strYear = ""
'strMonth = ""
'strMod = ""
'strQry1 = ""
'
'If Len(Trim(txt_filingmth.Text)) = 1 Then
'    txt_filingmth.Text = "0" & txt_filingmth.Text
'Else
'End If
'
'If fm_Whouseref.Opt_dtarch.Value = True Then
'
'    If opt_All.Value = True Then
'
'        If optyr.Value = True Then
'                Adodc1.RecordSource = "select Module as Department, File_Ref as Filereference, warehouse_ref as Warehouse_Ref, " _
'                & "convert(varchar(20),date_archived,103) as Date_Archived, convert(varchar(20),date_destroyed,103) as Date_Destroyed, Remarks " _
'                & "from ecmsfile_Reference " _
'                & "where datepart(yyyy,date_archived) = '" & Trim(txt_filingyr.Text) & "' " _
'                & "and File_ref like '%" & user_country & "%' order by module, file_ref "
'
'                strYear = Trim(txt_filingyr.Text)
'                strMod = ""
'                strMonth = ""
'
'                strQry1 = "select distinct(warehouse_ref)  " _
'                & "from ecmsfile_Reference " _
'                & "where datepart(yyyy,date_archived) = '" & Trim(txt_filingyr.Text) & "' " _
'                & "and File_ref like '%" & user_country & "%'  "
'        Else
'            Adodc1.RecordSource = "select Module as Department, File_Ref as Filereference, warehouse_ref as Warehouse_Ref, " _
'            & "convert(varchar(20),date_archived,103) as Date_Archived, convert(varchar(20),date_destroyed,103) as Date_Destroyed, Remarks " _
'            & "from ecmsfile_Reference " _
'            & "where datepart(yyyy,date_archived) = '" & Trim(txt_filingyrmth.Text) & "' " _
'            & "and datepart(mm,date_archived) = '" & Trim(txt_filingmth.Text) & "' " _
'            & "and File_ref like '%" & user_country & "%' order by module, file_ref "
'
'            strYear = Trim(txt_filingyrmth.Text)
'            strMod = ""
'            strMonth = Trim(txt_filingmth.Text)
'
'            strQry1 = "select distinct(warehouse_ref) " _
'            & "from ecmsfile_Reference " _
'            & "where datepart(yyyy,date_archived) = '" & Trim(txt_filingyrmth.Text) & "' " _
'            & "and datepart(mm,date_archived) = '" & Trim(txt_filingmth.Text) & "' " _
'            & "and File_ref like '%" & user_country & "%' "
'
'        End If
'
'    Else
'
'        If optyr.Value = True Then
'            Adodc1.RecordSource = "select Module as Department, File_Ref as Filereference, warehouse_ref as Warehouse_Ref, " _
'            & "convert(varchar(20),date_archived,103) as Date_Archived, convert(varchar(20),date_destroyed,103) as Date_Destroyed, Remarks " _
'            & "from ecmsfile_Reference " _
'            & "where datepart(yyyy,date_archived) = '" & Trim(txt_filingyr.Text) & "' " _
'            & "and Module = '" & StrDeptVal & "' " _
'            & "and File_ref like '%" & user_country & "%' order by file_ref "
'
'            strYear = Trim(txt_filingyr.Text)
'            strMod = StrDeptVal
'            strMonth = ""
'
'            strQry1 = "select distinct(warehouse_ref) " _
'            & "from ecmsfile_Reference " _
'            & "where datepart(yyyy,date_archived) = '" & Trim(txt_filingyr.Text) & "' " _
'            & "and Module = '" & StrDeptVal & "' " _
'            & "and File_ref like '%" & user_country & "%'  "
'        Else
'            Adodc1.RecordSource = "select Module as Department, File_Ref as Filereference, warehouse_ref as Warehouse_Ref, " _
'            & "convert(varchar(20),date_archived,103) as Date_Archived, convert(varchar(20),date_destroyed,103) as Date_Destroyed, Remarks " _
'            & "from ecmsfile_Reference " _
'            & "where datepart(yyyy,date_archived) = '" & Trim(txt_filingyrmth.Text) & "' " _
'            & "and datepart(mm,date_archived) = '" & Trim(txt_filingmth.Text) & "' " _
'            & "and Module = '" & StrDeptVal & "' " _
'            & "and File_ref like '%" & user_country & "%'  order by file_ref "
'
'            strYear = Trim(txt_filingyrmth.Text)
'            strMod = StrDeptVal
'            strMonth = Trim(txt_filingmth.Text)
'
'            strQry1 = "select distinct(warehouse_ref)  " _
'            & "from ecmsfile_Reference " _
'            & "where datepart(yyyy,date_archived) = '" & Trim(txt_filingyrmth.Text) & "' " _
'            & "and datepart(mm,date_archived) = '" & Trim(txt_filingmth.Text) & "' " _
'            & "and Module = '" & StrDeptVal & "' " _
'            & "and File_ref like '%" & user_country & "%' "
'
'        End If
'
'    End If
'
'ElseIf fm_Whouseref.Opt_dtdes.Value = True Then
'
'    If opt_All.Value = True Then
'
'        If optyr.Value = True Then
'                Adodc1.RecordSource = "select Module as Department, File_Ref as Filereference, warehouse_ref as Warehouse_Ref, " _
'                & "convert(varchar(20),date_archived,103) as Date_Archived, convert(varchar(20),date_destroyed,103) as Date_Destroyed, Remarks " _
'                & "from ecmsfile_Reference " _
'                & "where datepart(yyyy,date_destroyed) = '" & Trim(txt_filingyr.Text) & "' " _
'                & "and File_ref like '%" & user_country & "%' order by module, file_ref "
'
'                strYear = Trim(txt_filingyr.Text)
'                strMod = ""
'                strMonth = ""
'
'                strQry1 = "select distinct(warehouse_ref) " _
'                & "from ecmsfile_Reference " _
'                & "where datepart(yyyy,date_destroyed) = '" & Trim(txt_filingyr.Text) & "' " _
'                & "and File_ref like '%" & user_country & "%'  "
'        Else
'            Adodc1.RecordSource = "select Module as Department, File_Ref as Filereference, warehouse_ref as Warehouse_Ref, " _
'            & "convert(varchar(20),date_archived,103) as Date_Archived, convert(varchar(20),date_destroyed,103) as Date_Destroyed, Remarks " _
'            & "from ecmsfile_Reference " _
'            & "where datepart(yyyy,date_destroyed) = '" & Trim(txt_filingyrmth.Text) & "' " _
'            & "and datepart(mm,date_destroyed) = '" & Trim(txt_filingmth.Text) & "' " _
'            & "and File_ref like '%" & user_country & "%' order by module, file_ref "
'
'            strYear = Trim(txt_filingyrmth.Text)
'            strMod = ""
'            strMonth = Trim(txt_filingmth.Text)
'
'            strQry1 = "select distinct(warehouse_ref) " _
'            & "from ecmsfile_Reference " _
'            & "where datepart(yyyy,date_destroyed) = '" & Trim(txt_filingyrmth.Text) & "' " _
'            & "and datepart(mm,date_destroyed) = '" & Trim(txt_filingmth.Text) & "' " _
'            & "and File_ref like '%" & user_country & "%' "
'
'        End If
'
'    Else
'
'        If optyr.Value = True Then
'            Adodc1.RecordSource = "select Module as Department, File_Ref as Filereference, warehouse_ref as Warehouse_Ref, " _
'            & "convert(varchar(20),date_archived,103) as Date_Archived, convert(varchar(20),date_destroyed,103) as Date_Destroyed, Remarks " _
'            & "from ecmsfile_Reference " _
'            & "where datepart(yyyy,date_destroyed) = '" & Trim(txt_filingyr.Text) & "' " _
'            & "and Module = '" & StrDeptVal & "' " _
'            & "and File_ref like '%" & user_country & "%' order by file_ref "
'
'            strYear = Trim(txt_filingyr.Text)
'            strMod = StrDeptVal
'            strMonth = ""
'
'            strQry1 = "select distinct(warehouse_ref) " _
'            & "from ecmsfile_Reference " _
'            & "where datepart(yyyy,date_destroyed) = '" & Trim(txt_filingyr.Text) & "' " _
'            & "and Module = '" & StrDeptVal & "' " _
'            & "and File_ref like '%" & user_country & "%' "
'
'        Else
'            Adodc1.RecordSource = "select Module as Department, File_Ref as Filereference, warehouse_ref as Warehouse_Ref, " _
'            & "convert(varchar(20),date_archived,103) as Date_Archived, convert(varchar(20),date_destroyed,103) as Date_Destroyed, Remarks " _
'            & "from ecmsfile_Reference " _
'            & "where datepart(yyyy,date_destroyed) = '" & Trim(txt_filingyrmth.Text) & "' " _
'            & "and datepart(mm,date_destroyed) = '" & Trim(txt_filingmth.Text) & "' " _
'            & "and Module = '" & StrDeptVal & "' " _
'            & "and File_ref like '%" & user_country & "%'  order by file_ref "
'
'            strYear = Trim(txt_filingyrmth.Text)
'            strMod = StrDeptVal
'            strMonth = Trim(txt_filingmth.Text)
'
'            strQry1 = "select distinct(warehouse_ref) " _
'            & "from ecmsfile_Reference " _
'            & "where datepart(yyyy,date_destroyed) = '" & Trim(txt_filingyrmth.Text) & "' " _
'            & "and datepart(mm,date_destroyed) = '" & Trim(txt_filingmth.Text) & "' " _
'            & "and Module = '" & StrDeptVal & "' " _
'            & "and File_ref like '%" & user_country & "%'  "
'        End If
'
'    End If
'End If
'
'Debug.Print Adodc1.RecordSource
'Adodc1.Refresh
'StatusBar1.Visible = True
'
'If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
'        DataGrid_warehouse.Refresh
'        StatusBar1.Panels.Item(1).Text = "Records: " & Adodc1.Recordset.RecordCount
'Else
'        StatusBar1.Panels.Item(1).Text = "Records: 0"
'End If
'
'Exit Sub
'DataRefresh:
' MsgBox Err.Description & " #" & Err.Number
'End Sub
'
'Private Sub UpDownyrmth_Change()
'On Error GoTo ErrUpDnMth
'
'txt_filingyrmth.Text = UpDownyrmth.Value
'If Trim(txt_filingyrmth.Text) <> "" And txt_filingmth.Text <> "" Then
'    Call data_refresh
'    If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
'        DataGrid_warehouse.Refresh
'        StatusBar1.Panels.Item(1).Text = "Records: " & Adodc1.Recordset.RecordCount
'    End If
'    Exit Sub
'End If
'
'Exit Sub
'ErrUpDnMth:
' MsgBox Err.Description & " #" & Err.Number
'End Sub
'
'Private Sub selsearch()
'
'If Len(Trim(txt_filingmth.Text)) = 1 Then
'    txt_filingmth.Text = "0" & txt_filingmth.Text
'Else
'End If
'
'        Adodc1.RecordSource = "select Module as Department, File_Ref as Filereference, warehouse_ref as Warehouse_Ref, " _
'        & "convert(varchar(20),date_archived,103) as Date_Archived, convert(varchar(20),date_destroyed,103) as Date_Destroyed, Remarks " _
'        & "from ecmsfile_Reference " _
'        & "where File_Ref = '" & strFileref_LF & "' and File_Ref like '%" & user_country & "%' "
'
'    StatusBar1.Panels.Item(1).Text = ""
'    Debug.Print Adodc1.RecordSource
'    Adodc1.Refresh
'
'    If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
'        DataGrid_warehouse.Refresh
'        StatusBar1.Panels.Item(1).Text = "Records: " & Adodc1.Recordset.RecordCount
'    End If
'    Exit Sub
'
'    cmdGoback.Visible = True
'End Sub
'
'Private Sub DataGrid_Billing_Click()
''StatusBar1.Panels.Item(4).Text = (DE.rsCommand1.AbsolutePosition) & "/" & (DE.rsCommand1.RecordCount)
'StatusBar1.Panels.Item(4).Text = (RSConnect1.AbsolutePosition) & "/" & (RSConnect1.RecordCount)
'End Sub

Private Sub cmd_close_Click()
Unload Me
End Sub

'Private Sub DataGrid_Billing_DblClick()
'MsgBox DataGrid_Billing.Columns(Index).Caption
'End Sub

'''Private Sub DataGrid_Billing_Click()
'''Dim intI As Integer
'''If blnRecStat = True Then
'''    If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
''''        StatusBar1.Panels.Item(1).Text = "Records:  " & (Adodc1.Recordset.AbsolutePosition) & "/" & (Adodc1.Recordset.RecordCount)
'''        For intI = 0 To 7
'''            DataGrid_Billing.Columns(intI).Locked = True
'''            DataGrid_Billing.AllowUpdate = False
'''        Next intI
'''        For intI = 8 To 14
'''            DataGrid_Billing.AllowUpdate = True
'''        Next intI
'''    End If
'''End If
'''
'''End Sub
'''
'''Private Sub DataGrid_Billing_ColEdit(ByVal ColIndex As Integer)
'''Dim intI As Integer
'''If blnRecStat = True Then
'''    If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
'''        StatusBar1.Panels.Item(1).Text = "Records:  " & (Adodc1.Recordset.AbsolutePosition) & "/" & (Adodc1.Recordset.RecordCount)
'''        For intI = 0 To 7
'''            DataGrid_Billing.Columns(intI).Locked = True
'''            DataGrid_Billing.AllowUpdate = False
'''        Next intI
'''        For intI = 8 To 14
'''            DataGrid_Billing.AllowUpdate = True
'''        Next intI
'''    End If
'''End If
'''End Sub
'''
'Private Sub DataGrid_Billing_HeadClick(ByVal ColIndex As Integer)
'Dim RSDgrid As ADODB.Recordset
''Dim sortField As String
''Dim sortString As String
'
'If IntSel = 1 Then
'    Set RSDgrid = RSConnect1
'ElseIf IntSel = 2 Then
'    Set RSDgrid = RSConnect2
'ElseIf IntSel = 3 Then
'    Set RSDgrid = RSConnect3
'End If
'
'sortField = DataGrid_Billing.Columns(ColIndex).Caption
'
'If InStr(1, sortField, "Remarks") = 0 Then
'
'    If InStr(RSDgrid.Sort, "Asc") Then
'        sortString = sortField & " Desc"
'    Else
'        sortString = sortField & " Asc"
'    End If
'    RSDgrid.Sort = sortString
'
'Else
'    MsgBox "Cannot sort by this Column ", vbInformation, "Sort"
'    Exit Sub
'End If
'End Sub
Private Sub cmd_from_Click()
IntStatFlag = 1
MonthView1.Visible = True
MonthView1.Left = txt_DateFrom.Left + 6500
MonthView1.Top = txt_DateFrom.Top + 100

End Sub

Private Sub cmd_to_Click()
IntStatFlag = 2
MonthView1.Visible = True
MonthView1.Left = txt_DateTo.Left + 6500
MonthView1.Top = txt_DateTo.Top + 100
End Sub

Private Sub Form_Activate()
    IntStatFlag = 0
    IntStat = 1
End Sub

Private Sub Form_Click()
MonthView1.Visible = False
MonthView1.Value = Date
End Sub

Public Sub animationstart()
    Animation1.Visible = True
    With Animation1
        .Open App.Path & "\globe.avi"
        .Play
    End With
        
    Me.MousePointer = 11
    StatusBar1.Panels.Item(2).Text = "Progressing..."
    
End Sub

Public Sub animationend()
    Animation1.Stop
    Animation1.Visible = False
    
    MousePointer = 1
    StatusBar1.Panels.Item(2).Text = ""
End Sub

Private Sub data_refresh()
On Error GoTo DataRefresh

DataGrid_Billing.ClearFields

'If IntSel <> 0 Then
'    If IntSel = 1 Then
'        animationstart

        Set DataGrid_Billing.DataSource = RSConnect1
        DataGrid_Billing.BackColor = RGB(109, 71, 86)

'    ElseIf IntSel = 2 Then
'        Set DataGrid_Billing.DataSource = RSConnect2
'        DataGrid_Billing.BackColor = RGB(109, 71, 86)
'    ElseIf IntSel = 3 Then
'        Set DataGrid_Billing.DataSource = RSConnect3
'        DataGrid_Billing.BackColor = RGB(15, 95, 114)
'    End If
''        DataGrid_Billing.ReBind
        DataGrid_Billing.Refresh

'End If

'    RSConnect.Close
'    Set RSConnect = Nothing

Exit Sub
DataRefresh:
 MsgBox Err.Description & " #" & Err.Number

End Sub

Private Sub Form_Load()
    IntStatFlag = 0
    IntStat = 1
    Call connectdb
'   Call GetData

'Call Stored procedure to insert records from BillTransactionImport to BillTransactionDetails
'Convert Filereference field from to Longform And Shortform

Call FileRef_SF_LF
End Sub

Private Sub frame_Print_Click()
MonthView1.Visible = False
End Sub

Private Sub Frame1_Click()
MonthView1.Visible = False
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
If IntStatFlag = 1 Then
    txt_DateFrom.Text = MonthView1.Value
ElseIf IntStatFlag = 2 Then
    txt_DateTo.Text = MonthView1.Value
End If
MonthView1.Visible = False
'''If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
'''    If Opt_MInv.Value = True Then
'''        If (Month(CDate(txt_DateFrom.Text)) = Month(CDate(txt_DateTo.Text))) And (Year(CDate(txt_DateFrom.Text)) = Year(CDate(txt_DateTo.Text))) Then
'''            Call GetData
'''            Call data_refresh
'''        Else
'''            MsgBox "Date not valid for Monthly Report", vbInformation, "Validate Date"
'''            DataGrid_Billing.ClearFields
'''            DataGrid_Billing.Refresh
'''        End If
'''    Else
'''        Call GetData
'''        Call data_refresh
'''    End If
'''End If
End Sub

Private Sub cmd_clear_Click()
Refobj.Text = Empty
End Sub

Private Sub MonthView1_LostFocus()
If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
    If Opt_MInv.Value = True Then
        If (Month(CDate(txt_DateFrom.Text)) = Month(CDate(txt_DateTo.Text))) And (Year(CDate(txt_DateFrom.Text)) = Year(CDate(txt_DateTo.Text))) Then
''            Call GetData
''            Call data_refresh
        Else
            MsgBox "Date not valid for Monthly Report", vbInformation, "Validate Date"
            DataGrid_Billing.ClearFields
            DataGrid_Billing.Refresh
        End If
        Exit Sub
    Else
        If CDate(txt_DateFrom.Text) <= CDate(txt_DateTo.Text) Then
''            Call GetData
''            Call data_refresh
        Else
            MsgBox "Not a valid Date", vbInformation, "Validate Date"
            DataGrid_Billing.ClearFields
            DataGrid_Billing.Refresh
        End If
        Exit Sub
    End If
Else
    Exit Sub
End If

End Sub

Private Sub Opt_MInv_Click()
IntStat = 1
End Sub

Private Sub Opt_MInv_LostFocus()
''If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
''        Call GetData
''        Call data_refresh
''        Exit Sub
''End If

If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
    If Opt_MInv.Value = True Then
        If (Month(CDate(txt_DateFrom.Text)) = Month(CDate(txt_DateTo.Text))) And (Year(CDate(txt_DateFrom.Text)) = Year(CDate(txt_DateTo.Text))) Then
''            Call GetData
''            Call data_refresh
        Else
            MsgBox "Date not valid for Monthly Report", vbInformation, "Validate Date"
            DataGrid_Billing.ClearFields
            DataGrid_Billing.Refresh
        End If
        Exit Sub
    Else
        If CDate(txt_DateFrom.Text) <= CDate(txt_DateTo.Text) Then
''            Call GetData
''            Call data_refresh
        Else
            MsgBox "Not a valid Date", vbInformation, "Validate Date"
            DataGrid_Billing.ClearFields
            DataGrid_Billing.Refresh
        End If
        Exit Sub
    End If
Else
    Exit Sub
End If
End Sub

Private Sub Opt_SInv_Click()
IntStat = 2
End Sub

Private Sub Opt_SInv_LostFocus()
'''If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
'''    Call GetData
'''    Call data_refresh
'''    Exit Sub
'''End If
If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
    If Opt_MInv.Value = True Then
        If (Month(CDate(txt_DateFrom.Text)) = Month(CDate(txt_DateTo.Text))) And (Year(CDate(txt_DateFrom.Text)) = Year(CDate(txt_DateTo.Text))) Then
''            Call GetData
''            Call data_refresh
        Else
            MsgBox "Date not valid for Monthly Report", vbInformation, "Validate Date"
            DataGrid_Billing.ClearFields
            DataGrid_Billing.Refresh
        End If
        Exit Sub
    Else
        If CDate(txt_DateFrom.Text) <= CDate(txt_DateTo.Text) Then
''            Call GetData
''            Call data_refresh
        Else
            MsgBox "Not a valid Date", vbInformation, "Validate Date"
            DataGrid_Billing.ClearFields
            DataGrid_Billing.Refresh
        End If
        Exit Sub
    End If
Else
    Exit Sub
End If
End Sub

Private Sub txt_ClientCode_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
    If txt_ClientCode.Text = Empty Or txt_ClientCode.Text = "" Then
        MsgBox "Please enter the client code", vbInformation, "Enter Client Code"
        Exit Sub
    Exit Sub
    End If
Refobj.Text = txt_ClientCode.Text
txt_ClientCode.Text = Refobj.fivechar
txt_ClientCode.Refresh
End If

End Sub

Private Sub txt_ClientCode_LostFocus()
If txt_ClientCode.Text = Empty Or txt_ClientCode.Text = "" Then
MsgBox "Please Enter the Client code", vbInformation, "Client Code"
Else
    Refobj.Text = txt_ClientCode.Text
    txt_ClientCode.Text = Refobj.fivechar
    txt_ClientCode.Refresh
''''    If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" Then
''''        If CDate(txt_DateFrom.Text) < CDate(txt_DateTo.Text) Then
''''            Call CallProc
''''            Call GetData
''''            Call data_refresh
''''            Exit Sub
''''        End If
''''    End If
End If

End Sub

