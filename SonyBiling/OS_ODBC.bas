Attribute VB_Name = "OS_ODBC"
Public DBConnect As New ADODB.Connection
Public RSConnect1 As New ADODB.Recordset
Public RSConnect2 As New ADODB.Recordset
Public RSConnect3 As New ADODB.Recordset
Public RSConnect4 As New ADODB.Recordset
Public IntX1, IntX2, IntX3 As Integer
Public StrConnect As String

Public IntRenref As Integer
Public IntRepref As Integer
Public IntStatFlag As Integer

Public refconvert As New FileReference

Public prm As ADODB.Parameter
Public cmdProc As New ADODB.Command
Public intProcRef As Integer

Public prmSum As ADODB.Parameter
Public cmdSumProc As New ADODB.Command
Public intProcSumRef As Integer

Public Const REG_SZ = 1    'Constant for a string variable type.
Public Const HKEY_LOCAL_MACHINE = &H80000002

Public Declare Function RegCreateKey Lib "advapi32.dll" Alias _
"RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
phkResult As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal _
cbData As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long

Public Const VER_PLATFORM_WIN32s As Long = 0 ' Win32s on Windows 3.1.
Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1 'Win32 on Windows 95.
Public Const VER_PLATFORM_WIN32_NT As Long = 2
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Type OSVERSIONINFO
OSVSize     As Long
dwVerMajor   As Long
dwVerMinor   As Long
dwBuildNumber  As Long
PlatformID   As Long
szCSDVersion  As String * 128
End Type

Public idwidth, cwidth, datewidth, eventwidth As Integer
Public StrSearch As String
Public StrFormNo As Integer

Global Const CB_ERR = -1
Global Const CB_FINDSTRING = &H14C
Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Any) As Long

Public Forms(3) As Form
Public SqlContact, SqlContactCl, SqlContactClTM, SqlContactClPAT As String
Public SqlContactNonCl, SqlPotContact, SqlMergedContact As String

Public IntSel As Integer
Public IntStat As Integer
Public n As Integer

Public RSContact As ADODB.Recordset
Public CmdContact As ADODB.Command

Public strval As String

Public Sub create_odbc()
   Dim DataSourceName As String
   Dim DatabaseName As String
   Dim Description As String
   Dim DriverPath As String
   Dim DriverName As String
   Dim LastUser As String
   Dim Regional As String
   Dim Server As String

   Dim lResult As Long
   Dim hKeyHandle As Long
   
   '*****How to find OS
   Dim ostext As String
   
#If Win32 Then
Dim OSV As OSVERSIONINFO
OSV.OSVSize = Len(OSV)
If GetVersionEx(OSV) = 1 Then
Select Case OSV.PlatformID
Case VER_PLATFORM_WIN32_WINDOWS
  ostext = "Windows 95/98"
Case VER_PLATFORM_WIN32_NT
  ostext = "Windows NT/2000"
End Select
End If
#End If
'******************
   'Specify the DSN parameters.
'**************dsn="ECSFBilling"
   DataSourceName = "ECSFBilling"
   DatabaseName = "Reports"
   Description = "SonyBilling"
   If ostext = "Windows 95/98" Then
   DriverPath = "C:\Windows\System\"
   ElseIf ostext = "Windows NT/2000" Then
   DriverPath = "C:\WINNT\SYSTEM32\"
   End If
   LastUser = "ECSFBilling"
''   Server = "DBServer1"
   Server = "sql2008"
''   Server = "sqltest"
''   Server = "mails"
   DriverName = "SQL Server"

   'Create the new DSN key.

   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & _
        DataSourceName, hKeyHandle)

   'Set the values of the new DSN key.

   lResult = RegSetValueEx(hKeyHandle, "Database", 0&, REG_SZ, _
      ByVal DatabaseName, Len(DatabaseName))
   lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, _
      ByVal Description, Len(Description))
   lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, _
      ByVal DriverPath, Len(DriverPath))
   lResult = RegSetValueEx(hKeyHandle, "LastUser", 0&, REG_SZ, _
      ByVal LastUser, Len(LastUser))
   lResult = RegSetValueEx(hKeyHandle, "Server", 0&, REG_SZ, _
      ByVal Server, Len(Server))

   'Close the new DSN key.

   lResult = RegCloseKey(hKeyHandle)

   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, _
      "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
   lResult = RegSetValueEx(hKeyHandle, DataSourceName, 0&, REG_SZ, _
      ByVal DriverName, Len(DriverName))
   lResult = RegCloseKey(hKeyHandle)

'**********************
End Sub

Public Sub connectdb()
StrConnect = "DSN=ECSFBilling;UID=ECSFBilling;PWD=ECSFBilling;"
Set DBConnect = New ADODB.Connection
DBConnect.ConnectionString = StrConnect
DBConnect.CursorLocation = adUseClient
DBConnect.Open

intProcRef = 0
intProcSumRef = 0

End Sub

Public Sub GetData()

Dim StrDelTemp As String

Dim dtFromDate As Date
Dim dtToDate As Date

''''SqlContact = "Select Country, LawFirm , " _
''''                & "CurrencyCode , ChargedDate , DebitNoteNo , BillTransactionDetails.InvoiceNo, ClientRefNo, ClientCode, " _
''''                & "FileRefNo , AttorneyName , HourlyRate, BillTransactionDetails.BillCode, ApplicationNo, PatentNo, " _
''''                & "BillingCode, serviceFee, OfficialFee, Others, TotalAmount, Description, CurWords " _
''''                & "from BillTransactionDetails, BillingCode, Curwords " _
''''                & "Where Curwords.InvoiceNo = BillTransactionDetails.InvoiceNo"
                
dtFromDate = CDate(frm_Billing.txt_DateFrom.Text)
dtToDate = CDate(frm_Billing.txt_DateTo.Text)

If IntStat = 1 Then
    
'''    SqlContact = SqlContact & " and datepart(year, chargeddate) = " & Year(dtFromDate) & "" _
'''                 & "and datepart(mm, chargeddate) = " & Month(dtToDate) & "  " _
'''                 & "and ClientBillingCode = BillingCode order by BillTransactionDetails.InvoiceNo, billingcode "
    SqlContact = "Select Country, LawFirm , " _
                & "CurrencyCode , MonInvDate , DebitNoteNo , InvoiceNo, ClientRefNo, " _
                & "FileRefNo , AttorneyName , HourlyRate, BillTransactionDetails.BillCode, ApplicationNo, PatentNo, " _
                & "BillingCode, serviceFee, OfficialFee, Others, TotalAmount, Description " _
                & "from BillTransactionDetails, BillingCode " _
                & "where datepart(year, chargeddate) = " & Year(dtFromDate) & " " _
                & "and datepart(mm, chargeddate) = " & Month(dtToDate) & "  " _
                & "and BillTransactionDetails.WIPCode = BillingCode.WIPCode " _
                & "and BillTransactionDetails.NarrativeCode = BillingCode.NarrativeCode " _
                & "and ClientBillingCode = BillingCode and TotalAmount <> 0.00 order by InvoiceNo, billingcode "
                 
                Set RSConnect1 = New ADODB.Recordset
                RSConnect1.Open SqlContact, DBConnect, adOpenKeyset, adLockOptimistic, adCmdText
    
ElseIf IntStat = 2 Then
''    SqlContact = SqlContact & " and chargeddate between '" & Format(frm_Billing.txt_DateFrom.Text, "MM/DD/YYYY") & "' and '" & Format(frm_Billing.txt_DateTo.Text, "MM/DD/YYYY") & "' " _
''    & " and ClientBillingCode = BillingCode order by BillTransactionDetails.InvoiceNo, billingcode "
    
    SqlContact = "Select Country, LawFirm , CurrencyCode , ChargedDate , DebitNoteNo , BillTransactionDetails.InvoiceNo, " _
                & "ClientRefNo, ClientCode, FileRefNo , AttorneyName , HourlyRate, BillTransactionDetails.BillCode, " _
                & "ApplicationNo, PatentNo, BillingCode, serviceFee, OfficialFee, Others, TotalAmount, " _
                & "BillingCode.Description, Curwords, DocType, Remarks from BillTransactionDetails, BillingCode , Billcode, CurWords " _
                & "where chargeddate between '" & Format(frm_Billing.txt_DateFrom.Text, "MM/DD/YYYY 00:00:00") & "' and '" & Format(frm_Billing.txt_DateTo.Text, "MM/DD/YYYY 23:59:59") & "' and " _
                & "ClientBillingCode = BillingCode And BillTransactionDetails.BillCode = BillCode.BillCode " _
                & "and BillTransactionDetails.WIPCode = BillingCode.WIPCode " _
                & "and BillTransactionDetails.NarrativeCode = BillingCode.NarrativeCode " _
                & "and BillTransactionDetails.InvoiceNo = CurWords.InvoiceNo and TotalAmount <> 0.00  " _
                & "order by BillTransactionDetails.InvoiceNo, billingcode "
''
''                SqlTemp = "Select distinct(InvoiceNo) " _
''                & "from BillTransactionDetails " _
''                & "where chargeddate between '" & Format(frm_Billing.txt_DateFrom.Text, "MM/DD/YYYY") & "' and '" & Format(frm_Billing.txt_DateTo.Text, "MM/DD/YYYY") & "' "
               
                Set RSConnect1 = New ADODB.Recordset
                RSConnect1.Open SqlContact, DBConnect, adOpenKeyset, adLockOptimistic, adCmdText
                
'''                StrDelTemp = "Delete from InvoiceTemp"
'''                              DBConnect.Execute StrDelTemp
'''
'''                Set RSConnect4 = New ADODB.Recordset
'''                RSConnect4.Open SqlTemp, DBConnect, adOpenKeyset, adLockOptimistic, adCmdText
ElseIf IntStat = 4 Then

    SqlContact = "select BillTransactionSummaryRep.BillCode, BillCode.Description, Invoiceno, sum(ServiceFee), " _
    & "sum(OfficialFee), sum(Others), " _
    & "Sum (TotalAmount), DebitNoteNo " _
    & "From BillTransactionSummaryRep, BillCode " _
    & "Where BillCode.BillCode = BillTransactionSummaryRep.BillCode " _
    & "group by BillTransactionSummaryRep.BillCode, Description, Invoiceno, DebitNoteNo " _
    & "order by BillTransactionSummaryRep.BillCode "
               
                Set RSConnect4 = New ADODB.Recordset
                RSConnect4.Open SqlContact, DBConnect, adOpenKeyset, adLockOptimistic, adCmdText
                
End If

End Sub

'Public Sub FileRef_SF_LF()
Public Sub Curwords()
Dim SqlProc As String
Dim strIns As String
Dim strDel As String
Dim StrInvNo As String
Dim CurTAmt As String

'SqlProc = "Select InvoiceNo,  " _
'            & "FileRefNo , FileRefNo_SF, BillingCode, TotalAmount  " _
'            & "from BillTransactionImport"
'
''Insert FileRef_Sf to Table BillTransactionImport
'
'Set RSConnect2 = New ADODB.Recordset
'RSConnect2.Open SqlProc, DBConnect, adOpenKeyset, adLockOptimistic, adCmdText
'
'If RSConnect2.BOF And RSConnect2.EOF Then
'Else
'    RSConnect2.MoveFirst
'    Do While Not RSConnect2.EOF
'        If RSConnect2!FileRefNo <> Null Or RSConnect2!FileRefNo <> "" Then
'            refconvert.Text = RSConnect2!FileRefNo
'            RSConnect2!FileRefNo = refconvert.LongForm
'            RSConnect2!FileRefNo_SF = refconvert.ShortForm
''''            If RSConnect2!TotalAmount <> "" Then
''''                Call CurToWords(CCur(RSConnect2!TotalAmount))
'''''                MsgBox CurrencyToWords
''''                RSConnect2!CurWords = CurrencyToWords
''                MsgBox RSConnect2!FileRefNo
''                MsgBox RSConnect2!FileRefNo_SF
'
''''            End If
'        End If
'        RSConnect2.MoveNext
'    Loop
'End If
'
'RSConnect2.Close
'Set RSConnect2 = Nothing

'Convert CurToWords
'SqlProc = "select invoiceno, sum(convert(numeric(9,2),totalamount)) as TotalAmount from BillTransactionImport " _
'           & " group by invoiceno "

SqlProc = "Select invoiceno, (Case when DocType = 'C' then (-1 * sum(convert(numeric(9,2),totalamount))) " _
        & " else sum(convert(numeric(9,2),totalamount)) end ) as TotalAmount from BillTransactionImport " _
        & " group by invoiceno, DocType "
           
                
'Insert CurToWords to Table CurWords
    
Set RSConnect3 = New ADODB.Recordset
RSConnect3.Open SqlProc, DBConnect, adOpenKeyset, adLockOptimistic, adCmdText

strDel = "Delete from CurWords "
          DBConnect.Execute strDel

If RSConnect3.BOF And RSConnect3.EOF Then
Else
    RSConnect3.MoveFirst
    Do While Not RSConnect3.EOF
        If RSConnect3!InvoiceNo <> Null Or RSConnect3!InvoiceNo <> "" Then
            StrInvNo = RSConnect3!InvoiceNo
            CurTAmt = Format(RSConnect3!totalamount, "#####.00")
                Call CurToWords(CCur(CurTAmt))
''                MsgBox CurrencyToWords
''                RSConnect3!CurWords = CurrencyToWords
                strIns = "Insert CurWords values('" & StrInvNo & "', '" & CurTAmt & "','" & CurrencyToWords & "' ) "
                DBConnect.Execute strIns
        End If
        RSConnect3.MoveNext
    Loop
End If
    
RSConnect3.Close
Set RSConnect3 = Nothing

End Sub

Public Sub CallProc()
''Call Stored Procedure
On Error GoTo errproc

If intProcRef = 0 Then
    cmdProc.CommandText = "ECSFBilling_PROC"
    cmdProc.CommandType = adCmdStoredProc
    cmdProc.CommandTimeout = 36000
    cmdProc.Name = "Bill"
    
    Set prm = cmdProc.CreateParameter("p_ClientCode", adVarChar, adParamInput, 5)
    cmdProc.Parameters.Append prm
    Set prm = cmdProc.CreateParameter("p_DNoteCode", adVarChar, adParamInput, 20)
    cmdProc.Parameters.Append prm
    Set prm = cmdProc.CreateParameter("p_MEndDate", adVarChar, adParamInput, 12)
    cmdProc.Parameters.Append prm
'    Set prm = cmdProc.CreateParameter("p_FromDate", adVarChar, adParamInput, 12)
'    cmdProc.Parameters.Append prm
'    Set prm = cmdProc.CreateParameter("p_ToDate", adVarChar, adParamInput, 12)
'    cmdProc.Parameters.Append prm
    intProcRef = intProcRef + 1
End If

Set cmdProc.ActiveConnection = DBConnect
'DBConnect.Bill frm_Billing.txt_ClientCode.Text, frm_Billing.txt_DNoteNo.Text, CStr(Format(CDate(MEndDate), "yyyy/mm/dd")), CStr(Format(CDate(frm_Billing.txt_DateFrom.Text), "mm/dd/yyyy")), CStr(Format(CDate(frm_Billing.txt_DateTo.Text), "mm/dd/yyyy"))
DBConnect.Bill frm_Billing.txt_ClientCode.Text, frm_Billing.txt_DNoteNo.Text, CStr(Format(CDate(MEndDate), "yyyy/mm/dd"))

Exit Sub
errproc:
MsgBox Err.Description, vbInformation
End Sub

Public Sub CallSumProc()
''Call Stored Procedure
Dim StrFDate As String
Dim StrTDate As String

On Error GoTo errproc

If intProcSumRef = 0 Then
    cmdSumProc.CommandText = "ECSFBilling_Summary_PROC"
    cmdSumProc.CommandType = adCmdStoredProc
    cmdSumProc.CommandTimeout = 36000
    cmdSumProc.Name = "BillSum"
    
    Set prmSum = cmdSumProc.CreateParameter("p_FDate", adVarChar, adParamInput, 12)
    cmdSumProc.Parameters.Append prmSum
    Set prmSum = cmdSumProc.CreateParameter("p_TDate", adVarChar, adParamInput, 12)
    cmdSumProc.Parameters.Append prmSum
'    Format(CDate(frm_Billing.txt_DateTo.Text), "dd/mm/yyyy")
    intProcSumRef = intProcSumRef + 1
End If

'StrFDate = CStr(Format(CDate(frm_Billing.txt_DateFrom.Text), "mm/dd/yyyy 00:00:00"))
'StrTDate = CStr(Format(CDate(frm_Billing.txt_DateTo.Text), "mm/dd/yyyy 23:59:59"))
'StrFDate = CStr(Format(CDate(frm_Billing.txt_DateFrom.Text), "mm/dd/yyyy"))
'StrTDate = CStr(Format(CDate(frm_Billing.txt_DateTo.Text), "mm/dd/yyyy"))

StrFDate = CStr(Format(CDate(frm_Billing.txt_DateFrom.Text), "yyyy/mm/dd"))
StrTDate = CStr(Format(CDate(frm_Billing.txt_DateTo.Text), "yyyy/mm/dd"))

Set cmdSumProc.ActiveConnection = DBConnect
DBConnect.BillSum StrFDate, StrTDate

Exit Sub

errproc:
MsgBox Err.Description, vbInformation
End Sub





