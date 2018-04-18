Attribute VB_Name = "Module2"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const CB_ERR = (-1)
Const CB_FINDSTRING = &H14C

Const CB_SHOWDROPDOWN = &H14F
'File exist's constant declaration
Const OFS_MAXPATHNAME = 128

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
'API used for checking whether a file exists
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

'Structure declaration for checking whether a file exists
Private Type OFSTRUCT
CBytes As Byte
fFixedDisk As Byte
nErrCode As Integer
Reserved1 As Integer
Reserved2 As Integer
szPathName(OFS_MAXPATHNAME) As Byte
End Type
'Making the structure a instant
Private typOfStruct As OFSTRUCT
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public vPubConnectString As String
Public Sub cboBox_GotFocus(cboBox As ComboBox)
    Dim i As Long
    i = SendMessage(cboBox.hwnd, CB_SHOWDROPDOWN, True, 0&)
End Sub

Public Sub cboBox_KeyPress(KeyAscii As Integer, cboBox As ComboBox)
 On Error Resume Next
    Dim cb As Long
    Dim FindString As String
    
    If KeyAscii < 32 Or KeyAscii > 127 Then Exit Sub
    
    If cboBox.SelLength = 0 Then
        FindString = cboBox.Text & Chr$(KeyAscii)
    Else
        FindString = Left$(cboBox.Text, cboBox.SelStart) & Chr$(KeyAscii)
    End If
    
    cb = SendMessage(cboBox.hwnd, CB_FINDSTRING, -1, ByVal FindString)
    
    If cb <> CB_ERR Then
        cboBox.ListIndex = cb
        cboBox.SelStart = Len(FindString)
        cboBox.SelLength = Len(cboBox.Text) - cboBox.SelStart
            KeyAscii = 0
    End If

End Sub
Public Sub Load2Combo(obj As ComboBox, ByVal qry As String, Optional ArgCount As Integer = 1)
'On Error GoTo Xit
    Dim con As New ADODB.Connection
    con.Open vPubFasConnectString
    Dim RS As New ADODB.Recordset
    Dim sql As String
    RS.Open qry, con, adOpenKeyset, adLockReadOnly
    While Not RS.EOF
        If ArgCount = 1 Then obj.AddItem RS(0) Else obj.AddItem RS(0) & "--" & RS(1)
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    con.Close
    Set con = Nothing
Xit:  Err.Clear
End Sub


Public Sub GetConnectionString()
'****************************************
'Get Connection Function Gets Properties of Connectionstring from
'ServerSetup.ini
'****************************************
'On Error GoTo Errhandler


Dim ValueFrom As String
Dim iret As Long
     'Server Name
     ValueFrom = String(256, 0)
     iret = GetPrivateProfileString("DB", "Server", "", ValueFrom, Len(ValueFrom), App.Path & "\ServerSetup.ini")
     If iret > 0 Then
          ValueFrom = Left(ValueFrom, iret)
     Else
          ValueFrom = ""
     End If
     vPubConnectString = ValueFrom
    
   
    
Exit Sub
Errhandler:
MsgBox "Change Database Connection Path in ServerSetup.ini", vbInformation
Err.Clear
End Sub
Public Sub GetCompanyInformation()
Dim RS As New ADODB.Recordset
RS.Open "Select * from Company ", cn, 1, 1
If Not RS.EOF Then
        vPubCompanyName = RS!CompanyName
        vPubAddress = RS!Address
        vPubRegno = RS!regno
        vPubPhone = RS!phone
        vPubFax = RS!fax
        vPubEmail = RS!email
        pubFlashMsg = RS("flashmsg")
End If

RS.Close
End Sub


Public Function TxtAcceptMoney(obj As Object, ByVal KeyAscii As Integer) As Integer
    If KeyAscii = 8 Then
        If Len(obj.Text) >= 1 Then
            TxtAcceptMoney = KeyAscii
            Exit Function
        Else
            KeyAscii = 0
        End If
    End If
    
    If KeyAscii = 13 Then SendKeys "{TAB}": TxtAcceptMoney = 0: Exit Function
    
    If KeyAscii = 46 Then
        If InStr(obj.Text, ".") Then
            TxtAcceptMoney = 0
        Else
            TxtAcceptMoney = KeyAscii: Exit Function
        End If
    End If
    
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        TxtAcceptMoney = KeyAscii
        Exit Function
    Else
        TxtAcceptMoney = 0
    End If
End Function



Public Function TxtAcceptNumeric(obj As Object, ByVal KeyAscii As Integer) As Integer
    If KeyAscii = 8 Then
            If Len(obj.Text) >= 1 Then
                    TxtAcceptNumeric = KeyAscii
                    Exit Function
            Else
                    TxtAcceptNumeric = 0
            End If
    End If
    If KeyAscii = 13 Then SendKeys "{TAB}": TxtAcceptNumeric = 0: Exit Function
    If KeyAscii >= 48 And KeyAscii <= 57 Then TxtAcceptNumeric = KeyAscii: Exit Function
    TxtAcceptNumeric = 0
End Function
'Function used to check whether the file exists or not
'Output will be either TRUE or FALSE
'If the file exist then it will be TRUE else FALSE
Public Function exists(ByVal sFilename As String) As Boolean
Dim typOfStruct As OFSTRUCT
'On Error Resume Next
If Len(sFilename) > 0 Then
OpenFile sFilename, typOfStruct, OF_EXIST
exists = typOfStruct.nErrCode <> 2
End If
End Function
Public Sub flxAcceptNumeric(obj As Object, KeyAscii As Integer)
  ' On Error Resume Next
  
    If KeyAscii = 8 Then obj.Text = Mid(obj.Text, 1, Len(obj.Text) - 1): Exit Sub
    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Then
        obj.Text = obj.Text & Chr(KeyAscii)
        Exit Sub
    End If
    
End Sub
Public Sub SafeExit()

On Error Resume Next

'Method for ensuring that all forms are unloaded from memory
'before exiting the program
Dim Form As Form

   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form
   
End Sub

