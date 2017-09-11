VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form frm_Billing 
   Caption         =   "Billing"
   ClientHeight    =   3195
   ClientLeft      =   -20025
   ClientTop       =   8580
   ClientWidth     =   4680
   Icon            =   "BillingReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdError 
      Caption         =   "Error Report"
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdExec 
      Height          =   495
      Left            =   11040
      Picture         =   "BillingReport.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Import"
      Top             =   120
      Width           =   720
   End
   Begin VB.TextBox txt_DNoteNo 
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
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   8160
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   86310913
      CurrentDate     =   38154
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "01481"
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   0
      TabIndex        =   19
      Top             =   960
      Width           =   11775
      Begin MSDataGridLib.DataGrid DataGrid_Billing 
         Bindings        =   "BillingReport.frx":0614
         Height          =   5805
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   10239
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
   Begin VB.CommandButton cmdrep 
      Height          =   495
      Left            =   10320
      Picture         =   "BillingReport.frx":0629
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Export"
      Top             =   7200
      Width           =   720
   End
   Begin VB.CommandButton cmd_close 
      Height          =   495
      Left            =   11040
      Picture         =   "BillingReport.frx":0933
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Exit"
      Top             =   7200
      Width           =   720
   End
   Begin VB.Frame frame_search 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   5160
      TabIndex        =   7
      Top             =   0
      Width           =   3495
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
         TabIndex        =   11
         Top             =   480
         Width           =   1335
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
         TabIndex        =   10
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmd_from 
         Caption         =   "..."
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmd_to 
         Caption         =   "..."
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "(dd/mm/yyyy)"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   480
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "(dd/mm/yyyy)"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   1305
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
         TabIndex        =   14
         Top             =   600
         Width           =   375
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
         TabIndex        =   12
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Frame frame_Print 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2760
      TabIndex        =   4
      Top             =   0
      Width           =   2175
      Begin VB.OptionButton Option1 
         Caption         =   "Summary Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   26
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton Opt_SInv 
         Caption         =   "Individual Invoice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton Opt_MInv 
         Caption         =   "Monthly Invoice Chart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   345
      Left            =   12000
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   609
      _Version        =   327681
      BackColor       =   12632256
      FullWidth       =   65
      FullHeight      =   23
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4200
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
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   4620
      TabIndex        =   21
      Top             =   2580
      Width           =   4680
   End
   Begin VB.Label lblDNoteNo 
      Caption         =   "Debit Note No"
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
      Left            =   120
      TabIndex        =   23
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lbl_fileref 
      Caption         =   "Client Code"
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
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frm_Billing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Refobj As New five_char
Dim xlApp As Excel.Application
Dim xlwb As Excel.Workbook
Dim file_obj As New Scripting.FileSystemObject

Private Sub cmdError_Click()
IntStat = 3
ReportForm.Show
End Sub

Private Sub cmdExec_Click()
If txt_ClientCode.Text = "" Then
    MsgBox "Please Enter Client Code", vbInformation, "Enter Client Code"
    Exit Sub
End If

If txt_DNoteNo.Text = "" Then
    MsgBox "Please Enter Debit Note No", vbInformation, "Enter Debit Note No"
    Exit Sub
End If

If txt_DateFrom.Text = "" And txt_DateTo.Text = "" Then
    MsgBox "Please Enter From and To date ", vbInformation, "Enter Debit Note No"
    Exit Sub
End If

If txt_ClientCode.Text <> "" And txt_DNoteNo.Text <> "" And txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" Then
        animationstart
        If IntStat = 1 Or IntStat = 4 Then
            Call LastDayOfMonth
        End If
        Call InsRec
'        Call FileRef_SF_LF
        Call Curwords
        Call CallProc
        Call CallSumProc
        Call GetData
        Call data_refresh
        animationend
        Exit Sub
End If
End Sub
Private Sub InsRec()
Dim strDel1, strIns1 As String
Dim strDateMon, strDateYr, strPeriod As String

strDateYr = DatePart("YYYY", Format(CDate(txt_DateFrom.Text), "mm/dd/yyyy 00:00:00"))
strDateMon = DatePart("m", Format(CDate(txt_DateFrom.Text), "mm/dd/yyyy 23:59:59"))

If Len(strDateMon) = 1 Then
    strPeriod = strDateYr & "0" & strDateMon
Else
    strPeriod = strDateYr & strDateMon
End If


strDel1 = "delete from [Reports].[dbo].BillTransactionImport "
DBConnect.Execute strDel1

strDel1 = "delete from [Reports].[dbo].BillTransactionDetails " _
        & " where ChargedDate Between '" & CStr(Format(CDate(txt_DateFrom.Text), "mm/dd/yyyy 00:00:00")) & "' And  '" & CStr(Format(CDate(txt_DateTo.Text), "mm/dd/yyyy  23:59:59")) & "' "
DBConnect.Execute strDel1

'--Old Query
'''strIns1 = "Insert into BillTransactionImport select  distinct(cy.Code), t.InvDate, t.InvNo, c.FileRef, Null, " _
'''         & " e.Code,  g.Code , 'Amount' = (select sum(v.ToBillAmt + v.AdjAmt + v.DistAmt) " _
'''         & " from [SPlegal].[dbo].tblARSalesTrx u, [SPlegal].[dbo].tblInvoiceDet v, " _
'''         & " [SPlegal].[dbo].tblCase w, [SPlegal].[dbo].tblCurrency x, " _
'''         & " [SPlegal].[dbo].tblEmployee y, [SPlegal].[dbo].tblChargeCode z, [SPlegal].[dbo].tblCustomer cl " _
'''         & " where t.InvDate between  '" & CStr(Format(CDate(txt_DateFrom.Text), "mm/dd/yyyy")) & "' And  '" & CStr(Format(CDate(txt_DateTo.Text), "mm/dd/yyyy")) & "' " _
'''         & " and u.DocType <> 'B' and u.RecStatus <> 'D' and u.CaseID = w.IDNo " _
'''         & " and u.CurrencyID = x.IDNo and u.EmployeeID = y.IDNo and u.IDNo = v.TrxID " _
'''         & " and v.LineType in ('C','D') and v.ChargeCodeID = z.IDNo " _
'''         & " and cl.IDNo = c.CustomerID and cl.code = '01481' " _
'''         & " and v.ToBillAmt + v.AdjAmt + v.DistAmt <> 0 " _
'''         & " and t.InvNo = u.InvNo and g.Code = z.Code " _
'''         & " group by u.InvNo, z.Code) " _
'''         & " from [SPlegal].[dbo].tblARSalesTrx t, [SPlegal].[dbo].tblInvoiceDet d, " _
'''         & " [SPlegal].[dbo].tblCase c, [SPlegal].[dbo].tblCurrency cy, " _
'''         & " [SPlegal].[dbo].tblEmployee e, [SPlegal].[dbo].tblChargeCode g, [SPlegal].[dbo].tblCustomer cl " _
'''         & " where t.InvDate between '" & CStr(Format(CDate(txt_DateFrom.Text), "mm/dd/yyyy")) & "' And  '" & CStr(Format(CDate(txt_DateTo.Text), "mm/dd/yyyy")) & "' " _
'''         & " and t.DocType <> 'B' and t.RecStatus <> 'D' and t.CaseID = c.IDNo " _
'''         & " and t.CurrencyID = cy.IDNo and t.EmployeeID = e.IDNo and t.IDNo = d.TrxID " _
'''         & " and d.LineType in ('C','D') and d.ChargeCodeID = g.IDNo " _
'''         & " and cl.IDNo = c.CustomerID and cl.code = '01481' " _
'''         & " and d.ToBillAmt + d.AdjAmt + d.DistAmt <> 0 " _
'''         & " group by cy.Code, t.InvDate, t.InvNo, c.FileRef, " _
'''         & " e.Code , g.Code, c.CustomerID "

'--Credit note not included

'strIns1 = "Insert into BillTransactionImport " _
'        & "select  distinct(cy.Code), t.InvDate, t.InvNo, c.FileRef, Null,  e.Code,  g.Code , " _
'        & "'Amount' = (select sum(v.ToBillAmt + v.AdjAmt + v.DistAmt)  from " _
'        & "[SPlegal].[dbo].tblARSalesTrx u, [SPlegal].[dbo].tblInvoiceDet v, " _
'        & "[SPlegal].[dbo].tblCase w, [SPlegal].[dbo].tblCurrency x, " _
'        & "[SPlegal].[dbo].tblEmployee y, [SPlegal].[dbo].tblChargeCode z, " _
'        & "[SPlegal].[dbo].tblCustomer cl  where u.InvDate Between '" & CStr(Format(CDate(txt_DateFrom.Text), "mm/dd/yyyy")) & "' And  '" & CStr(Format(CDate(txt_DateTo.Text), "mm/dd/yyyy")) & "' " _
'        & "and u.DocType <> 'B' and u.RecStatus <> 'D' and u.CaseID = w.IDNo " _
'        & "and u.CurrencyID = x.IDNo and u.EmployeeID = y.IDNo and u.IDNo = v.TrxID " _
'        & "and v.LineType in ('C','D') and v.ChargeCodeID = z.IDNo  and cl.IDNo = c.CustomerID " _
'        & "and cl.code = '01481' and v.ToBillAmt + v.AdjAmt + v.DistAmt <> 0 " _
'        & "and t.InvNo = u.InvNo and g.Code = z.Code group by u.InvNo, z.Code) " _
'        & "from [SPlegal].[dbo].tblARSalesTrx t, [SPlegal].[dbo].tblInvoiceDet d, " _
'        & "[SPlegal].[dbo].tblCase c, [SPlegal].[dbo].tblCurrency cy, " _
'        & "[SPlegal].[dbo].tblEmployee e, [SPlegal].[dbo].tblChargeCode g, " _
'        & "[SPlegal].[dbo].tblCustomer cu  where t.InvDate Between '" & CStr(Format(CDate(txt_DateFrom.Text), "mm/dd/yyyy")) & "' And  '" & CStr(Format(CDate(txt_DateTo.Text), "mm/dd/yyyy")) & "' " _
'        & "and t.DocType <> 'B' and t.RecStatus <> 'D' and t.CaseID = c.IDNo " _
'        & "and t.CurrencyID = cy.IDNo and t.EmployeeID = e.IDNo and t.IDNo = d.TrxID " _
'        & "and d.LineType in ('C','D') and d.ChargeCodeID = g.IDNo  and cu.IDNo = c.CustomerID " _
'        & "and cu.code = '01481'  and d.ToBillAmt + d.AdjAmt + d.DistAmt <> 0 " _
'        & "and t.IDNo not in (Select distinct(a.InvTrxID) from [SPlegal].[dbo].tblarreceiptappln a, [SPlegal].[dbo].tblarreceiptappln b " _
'        & "where a.Type = 'AR01' and b.Type = 'AR01' and a.TrxID = b.TrxID " _
'        & "and a.period = '" & strPeriod & "' and a.InvTrxID is not Null) group by cy.Code, t.InvDate, t.InvNo, c.FileRef, " _
'        & "e.Code , g.Code, c.CustomerID "
    
'--Credit included
'---Error - Billine (WIPcode Not captured)
'strIns1 = "Insert into BillTransactionImport " _
'        & "select  distinct(Case when Isnull(O.Currency,'') = '' then 'SGD' else O.currency end) as Currency, O.ITEMDATE as InvDate, O.OpenitemNo as InvNo, C.IRN, " _
'        & "(Select Reports.dbo.fnFileRef_LF(officialNumber) from ECSF.DBO.OfficialNumbers R where CaseID = C.CaseID and NumberType = 'Z') as FileRef, " _
'        & "(Select officialNumber from ECSF.DBO.OfficialNumbers R where CaseID = C.CaseID and NumberType = 'Z')  as FileRef_SF, " _
'        & "(Select AbbreviatedName from ECSF.DBO.Employee where Employeeno = E.EmployeeNo) as EmpCode , B.WIPCODE as ChargeCode,  " _
'        & "(SELECT Narrativecode from ECSF.DBO.NARRATIVE where NARRATIVENO = B.NarrativeNo) as NarrativeCode,  " _
'        & "(Case when ltrim(rtrim(IsNull(b.shortnarrative,''))) ='' then ltrim(rtrim(Convert(varchar(8000),na.NarrativeText))) else ltrim(rtrim(IsNull(b.shortnarrative,''))) end ) as NarrativeText, B.FOREIGNVALUE as Amount,Null as doctype    from ECSF.DBO.OPENITEM O  " _
'        & "join ECSF.DBO.BILLLINE B on (B.ITEMENTITYNO=O.ITEMENTITYNO  and B.ITEMTRANSNO =O.ITEMTRANSNO) join ECSF.DBO.CASES C on (C.IRN=B.IRN) " _
'        & "left join ECSF.DBO.WORKHISTORY wh2 on (b.ITEMTRANSNO = wh2.reftransno  and b.ITEMLINENO = wh2.billlineno)  " _
'        & "LEFT JOIN ECSF.DBO.WORKHISTORY WH1 ON (WH1.REFTRANSNO = b.ITEMTRANSNO AND WH1.BILLLINENO = b.ITEMLINENO)  " _
'        & "LEFT JOIN ECSF.DBO.WORKHISTORY WH3 ON (WH3.REFTRANSNO = WH1.REFTRANSNO AND WH3.TRANSNO = WH1.TRANSNO  AND WH3.DISCOUNTFLAG = 1)  " _
'        & "LEFT JOIN ECSF.DBO.BILLLINE b1 on (b1.ITEMTRANSNO = wh3.reftransno AND b1.ITEMLINENO = WH3.BILLLINENO)  " _
'        & "left join ECSF.DBO.CASENAME EMP on (EMP.CASEID=C.CASEID and EMP.NAMETYPE='ATT' and EMP.EXPIRYDATE is null) left join ECSF.DBO.WIPCATEGORY W on (W.CATEGORYCODE=B.CATEGORYCODE) " _
'        & "left join ECSF.DBO.EMPLOYEE E   on (E.EMPLOYEENO=isnull((select min(WH.EMPLOYEENO) from ECSF.DBO.WORKHISTORY WH Where WH.REFENTITYNO=O.ITEMENTITYNO and WH.REFTRANSNO =O.ITEMTRANSNO and WH.BILLLINENO =B.ITEMLINENO and WH.EMPLOYEENO is not null),EMP.NAMENO))  " _
'        & "left join ECSF.DBO.TABLECODES T on (T.TABLECODE=E.STAFFCLASS)  " _
'        & "left join ECSF.DBO.ASSOCIATEDNAME AN on (AN.NAMENO=O.ACCTDEBTORNO and AN.RELATIONSHIP='RES' and AN.JOBROLE=101347) left join ECSF.DBO.NAME N   on (N.NAMENO=AN.RELATEDNAME)   " _
'        & "left join ECSF.DBO.COUNTRY CN   on (CN.COUNTRYCODE=C.COUNTRYCODE) left join ECSF.DBO.CASETYPE CT  on (CT.CASETYPE=C.CASETYPE)  " _
'        & "left join ECSF.DBO.VALIDPROPERTY VP on (VP.PROPERTYTYPE=C.PROPERTYTYPE and VP.COUNTRYCODE=(Select min(VP1.COUNTRYCODE) from ECSF.DBO.VALIDPROPERTY VP1 where VP1.PROPERTYTYPE=VP.PROPERTYTYPE and VP1.COUNTRYCODE in (C.COUNTRYCODE, 'ZZZ')))  " _
'        & "left join ECSF.DBO.CASENAME REF on (REF.CASEID=C.CASEID and REF.NAMETYPE=CASE WHEN(O.RENEWALDEBTORFLAG=1) THEN 'Z' ELSE 'D' END and REF.NAMENO=O.ACCTDEBTORNO and REF.EXPIRYDATE is null)  " _
'        & "LEFT JOIN ECSF.DBO.FEESCALCULATION FC ON (O.ACCTDEBTORNO = FC.DEBTOR AND B.WIPCODE = FC.DISBWIPCODE AND B.NARRATIVENO = FC.DISBNARRATIVE  AND FC.DISBWIPCODE = 'D0009' AND FC.DISBNARRATIVE = 712 )  " _
'        & "LEFT JOIN ECSF.DBO.OFFICIALNUMBERS ONU ON (ONU.CASEID = C.CASEID AND ONU.NUMBERTYPE = 'A') LEFT JOIN ECSF.DBO.NARRATIVE NA ON (NA.NARRATIVENO = b.NARRATIVENO) " _
'        & "LEFT JOIN ECSF.DBO.CASETEXT CX ON (CX.CASEID = C.CASEID AND CX.TEXTTYPE = 'BR') " _
'        & "Where O.ACCTDEBTORNO in (-56054,-22869) and o.Status = 1 " _
'        & "and O.ITEMDATE Between '" & CStr(Format(CDate(txt_DateFrom.Text), "mm/dd/yyyy 00:00:00")) & "' And  '" & CStr(Format(CDate(txt_DateTo.Text), "mm/dd/yyyy 23:59:59")) & "' "

'------(4 Spe 2017)---Updated After Inprotech v1 go live to cpature WIPcode from Workhistory when the WIPs are merged---------

strIns1 = "Insert into BillTransactionImport " _
        & "select  distinct(Case when Isnull(O.Currency,'') = '' then 'SGD' else O.currency end) as Currency, O.ITEMDATE as InvDate, O.OpenitemNo as InvNo, C.IRN, " _
        & "(Select Reports.dbo.fnFileRef_LF(officialNumber) from ECSF.DBO.OfficialNumbers R where CaseID = C.CaseID and NumberType = 'Z') as FileRef, " _
        & "(Select officialNumber from ECSF.DBO.OfficialNumbers R where CaseID = C.CaseID and NumberType = 'Z')  as FileRef_SF, " _
        & "(Select AbbreviatedName from ECSF.DBO.Employee where Employeeno = E.EmployeeNo) as EmpCode , " _
        & "(Select Top 1 WX.WIPcode from ECSF.DBO.WORKHISTORY WX where WX.REFTRANSNO = b.ITEMTRANSNO AND WX.BILLLINENO = b.ITEMLINENO and wX.MOVEMENTCLASS = 2 and wX.status <> 0) as ChargeCode, " _
        & "(SELECT Narrativecode from ECSF.DBO.NARRATIVE where NARRATIVENO = B.NarrativeNo) as NarrativeCode,  " _
        & "(Case when ltrim(rtrim(IsNull(b.shortnarrative,''))) ='' then ltrim(rtrim(Convert(varchar(8000),na.NarrativeText))) else ltrim(rtrim(IsNull(b.shortnarrative,''))) end ) as NarrativeText, B.FOREIGNVALUE as Amount,Null as doctype    from ECSF.DBO.OPENITEM O  " _
        & "join ECSF.DBO.BILLLINE B on (B.ITEMENTITYNO=O.ITEMENTITYNO  and B.ITEMTRANSNO =O.ITEMTRANSNO) join ECSF.DBO.CASES C on (C.IRN=B.IRN) " _
        & "Join ECSF.DBO.WORKHISTORY wh2 on (b.ITEMTRANSNO = wh2.reftransno  and b.ITEMLINENO = wh2.billlineno)  " _
        & "JOIN ECSF.DBO.WORKHISTORY WH1 ON (WH1.REFTRANSNO = b.ITEMTRANSNO AND WH1.BILLLINENO = b.ITEMLINENO)  " _
        & "LEFT JOIN ECSF.DBO.WORKHISTORY WH3 ON (WH3.REFTRANSNO = WH1.REFTRANSNO AND WH3.TRANSNO = WH1.TRANSNO  AND WH3.DISCOUNTFLAG = 1)  " _
        & "LEFT JOIN ECSF.DBO.BILLLINE b1 on (b1.ITEMTRANSNO = wh3.reftransno AND b1.ITEMLINENO = WH3.BILLLINENO)  " _
        & "left join ECSF.DBO.CASENAME EMP on (EMP.CASEID=C.CASEID and EMP.NAMETYPE='ATT' and EMP.EXPIRYDATE is null) left join ECSF.DBO.WIPCATEGORY W on (W.CATEGORYCODE=B.CATEGORYCODE) " _
        & "left join ECSF.DBO.EMPLOYEE E   on (E.EMPLOYEENO=isnull((select min(WH.EMPLOYEENO) from ECSF.DBO.WORKHISTORY WH Where WH.REFENTITYNO=O.ITEMENTITYNO and WH.REFTRANSNO =O.ITEMTRANSNO and WH.BILLLINENO =B.ITEMLINENO and WH.EMPLOYEENO is not null),EMP.NAMENO))  " _
        & "left join ECSF.DBO.TABLECODES T on (T.TABLECODE=E.STAFFCLASS)  " _
        & "left join ECSF.DBO.ASSOCIATEDNAME AN on (AN.NAMENO=O.ACCTDEBTORNO and AN.RELATIONSHIP='RES' and AN.JOBROLE=101347) left join ECSF.DBO.NAME N   on (N.NAMENO=AN.RELATEDNAME)   " _
        & "left join ECSF.DBO.COUNTRY CN   on (CN.COUNTRYCODE=C.COUNTRYCODE) left join ECSF.DBO.CASETYPE CT  on (CT.CASETYPE=C.CASETYPE)  " _
        & "left join ECSF.DBO.VALIDPROPERTY VP on (VP.PROPERTYTYPE=C.PROPERTYTYPE and VP.COUNTRYCODE=(Select min(VP1.COUNTRYCODE) from ECSF.DBO.VALIDPROPERTY VP1 where VP1.PROPERTYTYPE=VP.PROPERTYTYPE and VP1.COUNTRYCODE in (C.COUNTRYCODE, 'ZZZ')))  " _
        & "left join ECSF.DBO.CASENAME REF on (REF.CASEID=C.CASEID and REF.NAMETYPE=CASE WHEN(O.RENEWALDEBTORFLAG=1) THEN 'Z' ELSE 'D' END and REF.NAMENO=O.ACCTDEBTORNO and REF.EXPIRYDATE is null)  " _
        & "LEFT JOIN ECSF.DBO.FEESCALCULATION FC ON (O.ACCTDEBTORNO = FC.DEBTOR AND B.WIPCODE = FC.DISBWIPCODE AND B.NARRATIVENO = FC.DISBNARRATIVE  AND FC.DISBWIPCODE = 'D0009' AND FC.DISBNARRATIVE = 712 )  " _
        & "LEFT JOIN ECSF.DBO.OFFICIALNUMBERS ONU ON (ONU.CASEID = C.CASEID AND ONU.NUMBERTYPE = 'A') LEFT JOIN ECSF.DBO.NARRATIVE NA ON (NA.NARRATIVENO = b.NARRATIVENO) " _
        & "LEFT JOIN ECSF.DBO.CASETEXT CX ON (CX.CASEID = C.CASEID AND CX.TEXTTYPE = 'BR') " _
        & "Where O.ACCTDEBTORNO in (-56054,-22869) and o.Status = 1 and wh2.MOVEMENTCLASS = 2 and wh2.Status <> 0" _
        & "and O.ITEMDATE Between '" & CStr(Format(CDate(txt_DateFrom.Text), "mm/dd/yyyy 00:00:00")) & "' And  '" & CStr(Format(CDate(txt_DateTo.Text), "mm/dd/yyyy 23:59:59")) & "' "
'-------------------------------------------------------------------------------

'& "Where O.ACCTDEBTORNO = -56054 and o.Status = 1 " _
'strIns1 = "Insert into BillTransactionImport " _
'        & "select  distinct(cy.Code), t.InvDate, t.InvNo, c.FileRef, Null,  e.Code,  g.Code , " _
'        & "'Amount' = (select Case when u.Doctype = 'C' then -1 * sum(v.ToBillAmt + v.AdjAmt + v.DistAmt) else " _
'        & "sum(v.ToBillAmt + v.AdjAmt + v.DistAmt) end  from " _
'        & "[SPlegal].[dbo].tblARSalesTrx u, [SPlegal].[dbo].tblInvoiceDet v, " _
'        & "[SPlegal].[dbo].tblCase w, [SPlegal].[dbo].tblCurrency x, " _
'        & "[SPlegal].[dbo].tblEmployee y, [SPlegal].[dbo].tblChargeCode z, " _
'        & "[SPlegal].[dbo].tblCustomer cl  where u.InvDate Between '" & CStr(Format(CDate(txt_DateFrom.Text), "mm/dd/yyyy")) & "' And  '" & CStr(Format(CDate(txt_DateTo.Text), "mm/dd/yyyy")) & "' " _
'        & "and u.DocType <> 'B' and u.RecStatus <> 'D' and u.CaseID = w.IDNo " _
'        & "and u.CurrencyID = x.IDNo and u.EmployeeID = y.IDNo and u.IDNo = v.TrxID " _
'        & "and v.LineType in ('C','D') and v.ChargeCodeID = z.IDNo  and cl.IDNo = c.CustomerID " _
'        & "and c.CustomerID = '538' and v.ToBillAmt + v.AdjAmt + v.DistAmt <> 0 " _
'        & "and t.InvNo = u.InvNo and g.Code = z.Code group by u.InvNo, z.Code, u.doctype), doctype  " _
'        & "from [SPlegal].[dbo].tblARSalesTrx t, [SPlegal].[dbo].tblInvoiceDet d, " _
'        & "[SPlegal].[dbo].tblCase c, [SPlegal].[dbo].tblCurrency cy, " _
'        & "[SPlegal].[dbo].tblEmployee e, [SPlegal].[dbo].tblChargeCode g, " _
'        & "[SPlegal].[dbo].tblCustomer cu  where t.InvDate Between '" & CStr(Format(CDate(txt_DateFrom.Text), "mm/dd/yyyy")) & "' And  '" & CStr(Format(CDate(txt_DateTo.Text), "mm/dd/yyyy")) & "' " _
'        & "and t.DocType <> 'B' and t.RecStatus <> 'D' and t.CaseID = c.IDNo " _
'        & "and t.CurrencyID = cy.IDNo and t.EmployeeID = e.IDNo and t.IDNo = d.TrxID " _
'        & "and d.LineType in ('C','D') and d.ChargeCodeID = g.IDNo  and cu.IDNo = c.CustomerID " _
'        & "and c.CustomerID = '538'  " _
'        & "and d.ToBillAmt + d.AdjAmt + d.DistAmt <> 0 " _
'        & "group by cy.Code, doctype, t.InvDate, t.InvNo, c.FileRef, " _
'        & "e.Code , g.Code, c.CustomerID "
'
     
''' strIns1 = strIns1 & " UNION " _
'''        & "select  distinct(cy.Code), t.InvDate, t.InvNo, c.FileRef, Null,  e.Code,  g.Code , " _
'''        & "'Amount' = (select Case when u.Doctype = 'C' then -1 * sum(v.ToBillAmt + v.AdjAmt + v.DistAmt) else " _
'''        & "sum(v.ToBillAmt + v.AdjAmt + v.DistAmt) end  from " _
'''        & "[SPlegal].[dbo].tblARSalesTrx u, [SPlegal].[dbo].tblInvoiceDet v, " _
'''        & "[SPlegal].[dbo].tblCase w, [SPlegal].[dbo].tblCurrency x, " _
'''        & "[SPlegal].[dbo].tblEmployee y, [SPlegal].[dbo].tblChargeCode z, " _
'''        & "[SPlegal].[dbo].tblCustomer cl  where u.InvDate between '12/01/2009' And  '12/31/2009' and t.InvNo = 'SG26273/09' " _
'''        & "and u.DocType <> 'B' and u.RecStatus <> 'D' and u.CaseID = w.IDNo " _
'''        & "and u.CurrencyID = x.IDNo and u.EmployeeID = y.IDNo and u.IDNo = v.TrxID " _
'''        & "and v.LineType in ('C','D') and v.ChargeCodeID = z.IDNo  and cl.IDNo = c.CustomerID " _
'''        & "and c.CustomerID = '538' and v.ToBillAmt + v.AdjAmt + v.DistAmt <> 0 " _
'''        & "and t.InvNo = u.InvNo and g.Code = z.Code group by u.InvNo, z.Code, u.doctype), doctype  " _
'''        & "from [SPlegal].[dbo].tblARSalesTrx t, [SPlegal].[dbo].tblInvoiceDet d, " _
'''        & "[SPlegal].[dbo].tblCase c, [SPlegal].[dbo].tblCurrency cy, " _
'''        & "[SPlegal].[dbo].tblEmployee e, [SPlegal].[dbo].tblChargeCode g, " _
'''        & "[SPlegal].[dbo].tblCustomer cu  where t.InvDate  between '12/01/2009' And  '12/31/2009' and t.InvNo = 'SG26273/09'  " _
'''        & "and t.DocType <> 'B' and t.RecStatus <> 'D' and t.CaseID = c.IDNo " _
'''        & "and t.CurrencyID = cy.IDNo and t.EmployeeID = e.IDNo and t.IDNo = d.TrxID " _
'''        & "and d.LineType in ('C','D') and d.ChargeCodeID = g.IDNo  and cu.IDNo = c.CustomerID " _
'''        & "and c.CustomerID = '538'  " _
'''        & "and d.ToBillAmt + d.AdjAmt + d.DistAmt <> 0 " _
'''        & "group by cy.Code, doctype, t.InvDate, t.InvNo, c.FileRef, " _
'''        & "e.Code , g.Code, c.CustomerID "
        
         
        '& "and cu.code = '01481'  "
        
        
'        & "and t.IDNo not in (Select distinct(a.InvTrxID) from [SPlegal].[dbo].tblarreceiptappln a, [SPlegal].[dbo].tblarreceiptappln b " _
'        & "where a.Type = 'AR01' and b.Type = 'AR01' and a.TrxID = b.TrxID " _
'        & "and a.period = '" & strPeriod & "' and a.InvTrxID is not Null) group by cy.Code, t.InvDate, t.InvNo, c.FileRef, " _
         & "e.Code , g.Code, c.CustomerID "
         
  DBConnect.CommandTimeout = 36000
  DBConnect.Execute strIns1
  
End Sub
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

If IntStat = 1 Then
    
    If RSConnect1.RecordCount > 0 Then
        
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

        fldCount = RSConnect1.Fields.Count
    
        For iCol = 1 To fldCount
                xlWs.Cells(6, iCol).Value = RSConnect1.Fields(iCol - 1).Name
        Next
        
    '   Check version of Excel
        If Val(Mid(xlApp.Version, 1, InStr(1, xlApp.Version, ".") - 1)) > 8 Then
            'EXCEL 2000 or 2002: Use CopyFromRecordset
         
            ' Copy the recordset to the worksheet, starting in cell A2
            xlWs.Cells(6, 1).CopyFromRecordset RSConnect1
       
            'Note: CopyFromRecordset will fail if the recordset
            'contains an OLE object field or array data such
            'as hierarchical recordsets
        
        Else
            ' Copy recordset to an array
            recArray = RSConnect1.GetRows
            
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
                    .Font.Size = 8
                End With
            Next
        Next

    '   Auto-fit the column widths and row heights
'        xlApp.Selection.CurrentRegion.Columns.AutoFit
'        xlApp.Selection.CurrentRegion.Rows.AutoFit

'        xlwb.Close
'        xlApp.Quit

    '   Release Excel references
'       Set xlWs = Nothing
        xlwb.SaveAs App.Path & "\MonthlyInvoice_" & strMName & strYear & ".xls "
        Set xlwb = Nothing
        Set xlApp = Nothing
        
        RSConnect1.Close
        Set RSConnect1 = Nothing
    End If

End If
       
If IntStat = 2 Or IntStat = 4 Then
'    If RSConnect4.RecordCount > 0 Then
        ReportForm.Show
'    End If
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

Private Sub cmd_close_Click()
Unload Me
End Sub

Private Sub cmd_from_Click()
IntStatFlag = 1
MonthView1.Visible = True
MonthView1.Left = txt_DateFrom.Left + 6500
MonthView1.Top = txt_DateFrom.Top

End Sub

Private Sub cmd_to_Click()
IntStatFlag = 2
MonthView1.Visible = True
MonthView1.Left = txt_DateTo.Left + 6500
MonthView1.Top = txt_DateTo.Top
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
'    StatusBar1.Panels.Item(2).Text = "Progressing..."
    
End Sub

Public Sub animationend()
    Animation1.Stop
    Animation1.Visible = False
    
    MousePointer = 1
'    StatusBar1.Panels.Item(2).Text = ""
End Sub

Private Sub data_refresh()
On Error GoTo DataRefresh

DataGrid_Billing.ClearFields

    Set DataGrid_Billing.DataSource = RSConnect1
    DataGrid_Billing.BackColor = RGB(109, 71, 86)
        
    Set DataGrid_Billing.DataSource = RSConnect1
    DataGrid_Billing.BackColor = RGB(109, 71, 86)
        
Exit Sub
DataRefresh:
 MsgBox Err.Description & " #" & Err.Number

End Sub

Private Sub Form_Load()
    IntStatFlag = 0
    IntStat = 1
    Call create_odbc
    Call connectdb
''    Call FileRef_SF_LF
    MonthView1.Value = Date
End Sub

Private Sub frame_Print_Click()
MonthView1.Visible = False
End Sub

Private Sub Frame1_Click()
MonthView1.Visible = False
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
'MonthView1.ShowToday = True
If IntStatFlag = 1 Then
    txt_DateFrom.Text = MonthView1.Value
ElseIf IntStatFlag = 2 Then
    txt_DateTo.Text = MonthView1.Value
End If
MonthView1.Visible = False
End Sub

Private Sub cmd_clear_Click()
Refobj.Text = Empty
End Sub

Private Sub MonthView1_LostFocus()
If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
    If Opt_MInv.Value = True Then
        If (Month(CDate(txt_DateFrom.Text)) = Month(CDate(txt_DateTo.Text))) And (Year(CDate(txt_DateFrom.Text)) = Year(CDate(txt_DateTo.Text))) Then
            Call GetData
            Call data_refresh
        Else
            MsgBox "Date not valid for Monthly Report", vbInformation, "Validate Date"
            DataGrid_Billing.ClearFields
            DataGrid_Billing.Refresh
        End If
        Exit Sub
    Else
        If CDate(txt_DateFrom.Text) <= CDate(txt_DateTo.Text) Then
            Call GetData
            Call data_refresh
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
If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
    If (Month(CDate(txt_DateFrom.Text)) = Month(CDate(txt_DateTo.Text))) And (Year(CDate(txt_DateFrom.Text)) = Year(CDate(txt_DateTo.Text))) Then
        Call GetData
        Call data_refresh
        Exit Sub
    Else
        MsgBox "Not a valid Date", vbInformation, "Check Date"
    End If
End If
End Sub

Private Sub Opt_MInv_LostFocus()
''If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
''        Call GetData
''        Call data_refresh
''        Exit Sub
''End If

'''If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
'''    If Opt_MInv.Value = True Then
'''        If (Month(CDate(txt_DateFrom.Text)) = Month(CDate(txt_DateTo.Text))) And (Year(CDate(txt_DateFrom.Text)) = Year(CDate(txt_DateTo.Text))) Then
'''''            Call GetData
'''''            Call data_refresh
'''        Else
'''            MsgBox "Date not valid for Monthly Report", vbInformation, "Validate Date"
'''            DataGrid_Billing.ClearFields
'''            DataGrid_Billing.Refresh
'''        End If
'''        Exit Sub
'''    Else
'''        If CDate(txt_DateFrom.Text) <= CDate(txt_DateTo.Text) Then
'''''            Call GetData
'''''            Call data_refresh
'''        Else
'''            MsgBox "Not a valid Date", vbInformation, "Validate Date"
'''            DataGrid_Billing.ClearFields
'''            DataGrid_Billing.Refresh
'''        End If
'''        Exit Sub
'''    End If
'''Else
'''    Exit Sub
'''End If
End Sub

Private Sub Opt_SInv_Click()
IntStat = 2
If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
    If CDate(txt_DateFrom.Text) <= CDate(txt_DateTo.Text) Then
        Call GetData
        Call data_refresh
        Exit Sub
    Else
        MsgBox "Not a valid Date", vbInformation, "Check Date"
    End If
End If
End Sub

Private Sub Opt_SInv_LostFocus()
''If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
''    Call GetData
''    Call data_refresh
''    Exit Sub
''End If
'''If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
'''    If Opt_MInv.Value = True Then
'''        If (Month(CDate(txt_DateFrom.Text)) = Month(CDate(txt_DateTo.Text))) And (Year(CDate(txt_DateFrom.Text)) = Year(CDate(txt_DateTo.Text))) Then
'''''            Call GetData
'''''            Call data_refresh
'''        Else
'''            MsgBox "Date not valid for Monthly Report", vbInformation, "Validate Date"
'''            DataGrid_Billing.ClearFields
'''            DataGrid_Billing.Refresh
'''        End If
'''        Exit Sub
'''    Else
'''        If CDate(txt_DateFrom.Text) <= CDate(txt_DateTo.Text) Then
'''''            Call GetData
'''''            Call data_refresh
'''        Else
'''            MsgBox "Not a valid Date", vbInformation, "Validate Date"
'''            DataGrid_Billing.ClearFields
'''            DataGrid_Billing.Refresh
'''        End If
'''        Exit Sub
'''    End If
'''Else
'''    Exit Sub
'''End If
End Sub

Private Sub Option1_Click()
IntStat = 4
If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" And txt_ClientCode.Text <> "" Then
    If CDate(txt_DateFrom.Text) <= CDate(txt_DateTo.Text) Then
        Call GetData
        Call data_refresh
        Exit Sub
    Else
        MsgBox "Not a valid Date", vbInformation, "Check Date"
    End If
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
Else
    Refobj.Text = txt_ClientCode.Text
    txt_ClientCode.Text = Refobj.fivechar
    txt_ClientCode.Refresh
    If txt_DateFrom.Text <> "" And txt_DateTo.Text <> "" Then
        If CDate(txt_DateFrom.Text) < CDate(txt_DateTo.Text) Then
            Call GetData
            Call data_refresh
            Exit Sub
        End If
    End If
End If

End Sub

