VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const CLIENTCODE_MASK As String = "00000"
Private Const FILENUMBER_MASK As String = "00000"

Private m_sRefType As String
Private m_sText As String
Private m_sClientCode As String
Private m_sFileNumber As String
Private m_sTail As String
Private m_sSeparator As String
Private m_sShortForm As String
Private m_sLongForm As String
Private m_sStaff As String
Private m_bIsValidRef As Boolean
Private m_iSeparatorPos As Integer
Public Property Get RefType() As String
RefType = m_sRefType
End Property
Public Property Get Text() As String
Text = m_sText
End Property
Public Property Get ClientCode() As String
ClientCode = m_sClientCode
End Property
Public Property Get FileNumber() As String
FileNumber = m_sFileNumber
End Property
Public Property Get Tail() As String
Tail = m_sTail
End Property
Public Property Get Separator() As String
Separator = m_sSeparator
End Property
Public Property Get ShortForm() As String
ShortForm = m_sShortForm
End Property
Public Property Get LongForm() As String
LongForm = m_sLongForm
End Property

Public Property Get Staff() As String
Staff = m_sStaff
End Property

Public Property Get IsValidRef() As Boolean
IsValidRef = m_bIsValidRef
End Property

Public Property Let Text(sNewRef As String)
    PopulateProperties sNewRef
End Property

'Public Function InDocument(Optional bAdopt As Boolean = True) _
'    As Boolean
'Dim iColonPos As Integer, iBackSlashPos As Integer
'Dim sFileRef As String, iCounter As Integer
'With ActiveDocument
 '   'Check Bookmarks first
'    If SharedItems.BookmarkExists("FileRef") Then
'        sFileRef = .Bookmarks("FileRef").Range.Text
'
'    'Remove non-alphanumeric end characters
'        For iCounter = Len(sFileRef) To 1 Step -1
'            If Mid(sFileRef, iCounter, 1) >= "0" Then
'                sFileRef = Left(sFileRef, iCounter)
'                Exit For
'            End If
'        Next iCounter
'    End If
'
' '   'Check Document Custom Fields second
'    If sFileRef = "" Then
'        If DocPropertyExists("FileRef") Then
'            sFileRef = _
'                .CustomDocumentProperties.Item("FileRef").Value
'            InDocument = True
'        End If
'    End If
'End With
'
'If Len(sFileRef) > 3 Then
'    InDocument = True
'    If bAdopt Then PopulateProperties sFileRef
'Else
'    InDocument = False
'    sFileRef = ""
'End If
        
'End Function
Public Function MaskClientCode(ByVal sClientCode As String, _
Optional sShortForm As Boolean = True) As String
Dim sBuffer As String, iClientCodeLen As Integer
Dim sClientNumber As String, sClientAlphabet As String

If sClientCode = "" Then Exit Function
iClientCodeLen = Len(sClientCode)

Select Case m_sRefType
    Case "ECGM"
        sClientAlphabet = Left(sClientCode, 1)
        sClientNumber = Trim(Str(Val(Right(sClientCode, _
            iClientCodeLen - 1))))
        If sShortForm Then
            sBuffer = ""
        Else
            sBuffer = Right(CLIENTCODE_MASK, Len(CLIENTCODE_MASK) _
                - Len(sClientAlphabet) - Len(sClientNumber))
        End If
    Case "ECMS"
        sClientAlphabet = ""
        sClientNumber = Trim(Str(Val(sClientCode)))
        If sShortForm Then
            sBuffer = ""
        Else
            sBuffer = Right(CLIENTCODE_MASK, Len(CLIENTCODE_MASK) _
                - Len(sClientNumber))
        End If
    Case "ENQ"
        sClientAlphabet = Left(sClientCode, 4)
        sClientNumber = Right(sClientCode, 4)
        sBuffer = ""
    Case Else
        Exit Function
End Select

If Val(sClientNumber) > 0 Then
    MaskClientCode = sClientAlphabet & sBuffer & sClientNumber
End If
End Function
Public Function MaskFileNumber(ByVal sFileNumber As String, _
    Optional sShortForm As Boolean = True) As String
Dim sMask As String, iMaskLen As Integer
Dim iFileNumberLen As Integer

sFileNumber = Trim(Str(Val(sFileNumber)))
iFileNumberLen = Len(sFileNumber)
sMask = FILENUMBER_MASK
iMaskLen = Len(sMask)

If sShortForm Then
    MaskFileNumber = sFileNumber
Else
    MaskFileNumber = Left(sMask, iMaskLen - iFileNumberLen) & _
    sFileNumber
End If

End Function
Public Sub Clear(Optional bClearAll As Boolean = False)
m_sRefType = ""
m_sClientCode = ""
m_sFileNumber = ""
m_sTail = ""
m_sSeparator = ""
m_sShortForm = ""
m_sLongForm = ""
m_sStaff = ""
m_bIsValidRef = False
m_iSeparatorPos = 0
If bClearAll Then m_sText = ""
End Sub
Private Sub PopulateProperties(sNewRef As String)
Dim iColonPos As Integer, iNewRefLen As Integer
Dim sClientCode As String, sFileNumber As String

sNewRef = UCase(sNewRef)
iNewRefLen = Len(sNewRef)

m_bIsValidRef = False

'Separate file reference from staff initials
iColonPos = InStr(1, sNewRef, ":", vbTextCompare)
If iColonPos > 0 Then
    m_sStaff = Mid(sNewRef, iColonPos + 1, iNewRefLen - iColonPos)
    m_sText = Left(sNewRef, iColonPos - 1)
Else
    m_sStaff = ""
    m_sText = sNewRef
End If

m_sRefType = GetRefType()   'Gets RefType, Separator & SeparatorPos

If m_sRefType <> "" Then
    sClientCode = Left(m_sText, m_iSeparatorPos - 1)    'Get Client Code
    If sClientCode <> "" Then
        sFileNumber = GetFileNumber()   'Gets FileNumber & Tail
        If sFileNumber <> "" Then
            m_sFileNumber = MaskFileNumber(sFileNumber, True)
            m_sClientCode = MaskClientCode(sClientCode, True)
            m_sLongForm = MaskClientCode(m_sClientCode, False) & _
                m_sSeparator & MaskFileNumber(m_sFileNumber, False) _
                & m_sTail
            m_sShortForm = m_sClientCode & m_sSeparator & _
                m_sFileNumber & m_sTail
            m_bIsValidRef = True
        Else: Clear
        End If
    
    Else: Clear
    End If
Else: Clear
End If

End Sub
Private Function GetRefType() As String
Dim iHyphenPos As Integer, iNewRefLen As Integer
Dim sChar() As String, iCounter As Integer

GetRefType = ""
iHyphenPos = InStr(1, m_sText, "-", 1)
iNewRefLen = Len(m_sText)
If iNewRefLen > 0 Then ReDim sChar(1 To iNewRefLen) Else _
    Exit Function

For iCounter = 1 To iNewRefLen
    sChar(iCounter) = Mid(m_sText, iCounter, 1)
Next iCounter

Select Case iHyphenPos
    Case 0      'ECMS Reference
        If Not IsNumber(sChar(1)) Then Exit Function
        For iCounter = 2 To iNewRefLen
            If IsLetter(sChar(iCounter)) Then
                If iCounter < 7 Then
                    If iNewRefLen >= iCounter + 2 Then
                        If IsLetter(sChar(iCounter + 1)) And _
                            IsNumber(sChar(iCounter + 2)) Then
                            m_iSeparatorPos = iCounter
                            m_sSeparator = Mid(m_sText, _
                            iCounter, 2)
                            GetRefType = "ECMS"
                        End If
                        Exit Function
                    End If
                Else: Exit Function
                End If
            Else: If Not IsNumber(sChar(iCounter)) Then _
                Exit Function
            End If
        Next iCounter

    Case 3 To 6     'ECGM Reference
        If IsLetter(sChar(1)) Then
            If IsNumber(sChar(2)) Then
                If iHyphenPos >= 3 Then
                    For iCounter = 3 To iHyphenPos - 1
                        If Not IsNumber(sChar(iCounter)) Then
                            Exit Function
                        End If
                    Next iCounter
                    m_iSeparatorPos = iHyphenPos
                    m_sSeparator = "-"
                    GetRefType = "ECGM"
                Else: Exit Function
                End If
            Else: Exit Function
            End If
        Else: Exit Function
        End If
        
    Case 9      'ENQ Reference
        For iCounter = 1 To iHyphenPos - 5
            If Not IsLetter(sChar(iCounter)) Then _
                Exit Function
        Next iCounter
        For iCounter = iHyphenPos - 4 To iHyphenPos - 1
            If Not IsNumber(sChar(iCounter)) Then _
                Exit Function
        Next iCounter
        m_iSeparatorPos = iHyphenPos
        m_sSeparator = "-"
        GetRefType = "ENQ"
    
    Case Else
        Exit Function

End Select
End Function
Private Function GetFileNumber() As String
Dim sPostSeparator As String, iPostSeparatorLen As Integer
Dim sChar As String, sFileNumber As String
Dim sMask As String, iMaskLen As Integer
Dim iFileNumberLen As Integer, iCounter As Integer

iPostSeparatorLen = Len(m_sText) - m_iSeparatorPos - _
    Len(m_sSeparator) + 1
sPostSeparator = Right(m_sText, iPostSeparatorLen)

'Extract File Number
For iCounter = 1 To iPostSeparatorLen
    sChar = Mid(sPostSeparator, iCounter, 1)
    If IsNumber(sChar) Then
        sFileNumber = sFileNumber & sChar
    Else: Exit For
    End If
Next iCounter
iFileNumberLen = Len(sFileNumber)
If iFileNumberLen > Len(FILENUMBER_MASK) Or _
    Val(sFileNumber) < 1 Then
    GetFileNumber = ""
    Exit Function
End If
GetFileNumber = sFileNumber


'Extract Tail
If iFileNumberLen <= iPostSeparatorLen Then
    m_sTail = Mid(sPostSeparator, iFileNumberLen + 1, _
        Len(sPostSeparator) - iFileNumberLen + 1)
Else: m_sTail = ""
End If

End Function
Private Function IsLetter(sChar As String) As Boolean
If Len(sChar) = 1 Then
    Select Case Asc(sChar)
        Case 65 To 90
            IsLetter = True
        Case Else
            IsLetter = False
    End Select
Else: IsLetter = False
End If
End Function
Private Function IsNumber(sChar As String) As Boolean
If Len(sChar) = 1 Then
    Select Case Asc(sChar)
        Case 48 To 57
            IsNumber = True
        Case Else
            IsNumber = False
    End Select
Else: IsNumber = False
End If
End Function

