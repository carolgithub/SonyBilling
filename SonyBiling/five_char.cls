VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "five_char"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private codeok As String
Public Property Get fivechar() As String
fivechar = codeok
End Property

Public Property Let Text(refstring As String)
convert5char refstring
End Property
Private Sub convert5char(refstring As String)
Dim strlen As Integer
strlen = Len(refstring)
If strlen = 5 Then
codeok = refstring
Else
    If strlen = 4 Then
        codeok = "0" & refstring
    Else
        If strlen = 3 Then
            codeok = "00" & refstring
        Else
            If strlen = 2 Then
                codeok = "000" & refstring
            Else
                If strlen = 1 Then
                    codeok = "0000" & refstring
                Else
                End If
            End If
        End If
    End If
End If

End Sub
