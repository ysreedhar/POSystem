VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmNewUser 
   Caption         =   "User Management"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   639
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   779
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "User Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4800
      Left            =   720
      TabIndex        =   15
      Top             =   360
      Width           =   7335
      Begin VB.OptionButton optAdministrator 
         Caption         =   "Administrator"
         Height          =   495
         Left            =   2280
         TabIndex        =   5
         Top             =   2760
         Width           =   1575
      End
      Begin VB.OptionButton optGeneralUser 
         Caption         =   "General User"
         Height          =   495
         Left            =   3960
         TabIndex        =   6
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox txtDepartment 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000080&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2265
         MousePointer    =   3  'I-Beam
         TabIndex        =   2
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtUName 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000080&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2265
         MousePointer    =   3  'I-Beam
         TabIndex        =   1
         Top             =   840
         Width           =   3135
      End
      Begin VB.ListBox lstsecurity 
         BackColor       =   &H00FFFFFF&
         Height          =   1185
         ItemData        =   "frmNewUser.frx":0000
         Left            =   2265
         List            =   "frmNewUser.frx":000D
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   3300
         Width           =   3060
      End
      Begin VB.TextBox txtConfirmPassword 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000080&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2265
         MousePointer    =   3  'I-Beam
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2340
         Width           =   3135
      End
      Begin VB.TextBox txtNewPassword 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000080&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2265
         MousePointer    =   3  'I-Beam
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1860
         Width           =   3135
      End
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000080&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2265
         MousePointer    =   3  'I-Beam
         TabIndex        =   0
         Top             =   420
         Width           =   3135
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   315
         TabIndex        =   22
         Top             =   1380
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   315
         TabIndex        =   21
         Top             =   900
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access Rights"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   19
         Top             =   3240
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm  Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   315
         TabIndex        =   18
         Top             =   2400
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   17
         Top             =   1920
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   315
         TabIndex        =   16
         Top             =   480
         Width           =   630
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flxUsers 
      Height          =   2295
      Left            =   720
      TabIndex        =   8
      Top             =   5760
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   12648447
      BackColorFixed  =   65535
      ForeColorSel    =   14723990
      BackColorBkg    =   -2147483633
      GridColor       =   14723990
      GridColorFixed  =   14723990
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      ToolTipText     =   "Click to save Contents (Alt+S)"
      Top             =   8280
      Width           =   1230
      Caption         =   "Clear"
      PicturePosition =   131072
      Size            =   "2170;873"
      MousePointer    =   99
      Accelerator     =   115
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of Existing Users"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   20
      Top             =   5520
      Width           =   1740
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Left            =   3960
      TabIndex        =   10
      ToolTipText     =   "Click to save Contents (Alt+S)"
      Top             =   8265
      Width           =   1230
      Caption         =   "Save"
      PicturePosition =   131072
      Size            =   "2170;873"
      MousePointer    =   99
      Accelerator     =   115
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdDelete 
      Height          =   495
      Left            =   5400
      TabIndex        =   11
      ToolTipText     =   "Click to Clear Contents (Alt+A)"
      Top             =   8280
      Width           =   1230
      Caption         =   "Delete"
      PicturePosition =   131072
      Size            =   "2170;873"
      MousePointer    =   99
      Accelerator     =   97
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdExit 
      Height          =   495
      Left            =   6780
      TabIndex        =   12
      ToolTipText     =   "Click to Delete Contents (Alt+D)"
      Top             =   8265
      Width           =   1230
      Caption         =   "Close"
      PicturePosition =   131072
      Size            =   "2170;873"
      MousePointer    =   99
      Accelerator     =   100
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   495
      Left            =   3480
      TabIndex        =   14
      ToolTipText     =   "Click to Close (Alt+C)"
      Top             =   11865
      Width           =   495
      PicturePosition =   131072
      Size            =   "873;873"
      MousePointer    =   99
      Accelerator     =   99
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   495
      Left            =   2280
      TabIndex        =   13
      ToolTipText     =   "Click to Save Contents (Alt+S)"
      Top             =   11865
      Width           =   495
      PicturePosition =   131072
      Size            =   "873;873"
      MousePointer    =   99
      Accelerator     =   115
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lflag As String
Dim Uid As Long
Dim MaxNo As Long
Dim verify As Boolean
Dim RS As ADODB.Recordset
Dim listval As String
Private Function validate_fields() As Boolean
    verify = False
    If txtUsername.Text = "" Then
        MsgBox "Please enter the username", vbExclamation
        txtUsername.SetFocus
        validate_fields = False
    ElseIf txtNewPassword = "" Then
        MsgBox "Please enter the password", vbExclamation
        txtNewPassword.SetFocus
        validate_fields = False
    ElseIf txtConfirmPassword.Text = "" Then
        MsgBox "Please re-enter the new password", vbExclamation
        txtConfirmPassword.SetFocus
        validate_fields = False
    ElseIf txtConfirmPassword.Text <> txtNewPassword.Text Then
        MsgBox "New Password and confirm Password should be the same", vbExclamation
        txtConfirmPassword.SetFocus
        validate_fields = False
    Else
        validate_fields = True
    End If
End Function
Private Sub cmdcancel_Click()
Frame1.Visible = False
End Sub

Private Sub cmdClear_Click()
            txtUName.Text = ""
            txtDepartment.Text = ""
            txtUsername.Text = ""
            txtNewPassword.Text = ""
            txtConfirmPassword.Text = ""
            For i = 0 To lstsecurity.ListCount - 1
                lstsecurity.Selected(i) = False
            Next
End Sub

Private Sub cmdDelete_Click()
'On Error GoTo Xit
cn.Execute "Delete from tblusers  where Name='" & txtUsername.Text & "' and Password='" & txtNewPassword.Text & "'"
MsgBox "User Deleted", vbInformation, App.Title
Xit:
Err.Clear
End Sub
Private Sub cmdExit_Click()
    'On Error Resume Next
    Unload Me
End Sub
Private Sub CommandButton1_Click()
    Dim Msg, Style, response
    Dim Pass As String
    Dim Typ As String
If Lflag = "Delete" Then
    Msg = MsgBox("do you really want to delete the user", vbYesNo + vbInformation)
    If Msg = vbYes Then
    cn.Execute "delete from tblusers where id=" & Uid
    MsgBox "User name has been deleted", vbInformation, App.Title
    End If
Else
    Msg = "Do you really want to update this record ?"    ' Define message.
    Style = vbYesNo + vbCritical ' Define buttons.
    response = MsgBox(Msg, Style)
    If response = vbYes Then   ' User chose Yes.
        verify = validate_fields
       ' verify1 = CheckName
        If optAdministrator.Value = True Then
            Typ = "Yes"
        Else
            Typ = "No"
        End If
        If verify = True Then
            Pass = UCase(txtConfirmPassword.Text)
            If RS.State Then RS.Close
                RS.Open "select max(id) from tblusers"
                If RS.RecordCount > 0 Then
                    If Not IsNull(RS.Fields(0)) Then
                                MaxNo = RS.Fields(0) + 1
                    Else
                            MaxNo = 1
                    End If
                  Else
                    MaxNo = 1
                End If
            If RS.State Then RS.Close
            RS.Open "SELECT * from tblusers where name ='" & txtUsername.Text & "'", cn, 3, 2
             Call security
             If RS.EOF Then
             RS.AddNew
             'rs.Fields(0) = MaxNo
             Msg = "New user has been created sucessfully."
            Else
            Msg = " User Information has been Modified sucessfully."
            End If
            RS.Fields("UserID") = txtUsername.Text
            RS.Fields("Name") = txtUName.Text
            RS.Fields("Department") = txtDepartment.Text
            RS.Fields("Password") = txtNewPassword.Text
            RS.Fields("Type") = Typ
            RS.Fields("Accessrights") = listval
            RS.Update
            MsgBox Msg, vbExclamation
            cmdClear_Click
        End If
    End If
End If

Lflag = ""
End Sub
Function RepaintFlexGrid()
' Reset the backcolor
For ch = 1 To flxUsers.Rows - 1
For flxcls = 0 To flxUsers.Cols - 1
    flxUsers.Row = ch
    flxUsers.Col = flxcls
If flxUsers.CellBackColor = vbYellow Then flxUsers.CellBackColor = vbWhite
Next flxcls
Next ch
End Function

Private Sub CommandButton2_Click()

End Sub

Private Sub flxUsers_Click()
cmdClear_Click
current = flxUsers.Row
RepaintFlexGrid
'Current  row
flxUsers.Row = current
For i = 1 To flxUsers.Cols - 1
flxUsers.Col = i
flxUsers.CellBackColor = vbYellow
Next
LoadUserDetails
vprev = flxUsers.Row
flxUsers.Col = 1
End Sub
Sub LoadUserDetails()
cmdClear_Click
txtUName.Text = flxUsers.TextMatrix(flxUsers.Row, 1)
txtUsername.Text = flxUsers.TextMatrix(flxUsers.Row, 2)
txtDepartment.Text = flxUsers.TextMatrix(flxUsers.Row, 3)
txtNewPassword.Text = flxUsers.TextMatrix(flxUsers.Row, 4)
txtConfirmPassword.Text = flxUsers.TextMatrix(flxUsers.Row, 4)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Call UPPER_CASE(KeyAscii)
If KeyAscii = "27" Then Unload Me
End Sub

Private Sub Form_Load()
Set RS = New ADODB.Recordset
RS.ActiveConnection = cn
RS.CursorLocation = adUseClient
RS.CursorType = adOpenDynamic
RS.LockType = adLockOptimistic
LoadFlexTitles
LoadUserData
End Sub
Private Sub LoadFlexTitles()
On Error Resume Next
    With flxUsers
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
    .TextMatrix(0, 1) = "User Name"
    .ColWidth(1) = 1500
     .TextMatrix(0, 2) = "UserID"
     .ColWidth(2) = 1000
     .TextMatrix(0, 3) = "Department"
     .ColWidth(3) = 2000
     .TextMatrix(0, 4) = "Password"
     .ColWidth(4) = 1200
     .TextMatrix(0, 5) = "Type"
     .ColWidth(5) = 1000
     .TextMatrix(0, 6) = "AccessRights"
     .ColWidth(6) = 0
    End With
End Sub

Private Sub LoadUserData()
Dim fldata3 As New ADODB.Recordset
If fldata3.State Then fldata3.Close
fldata3.Open "select * from tblusers", cn, 3, 2
With flxUsers
.Rows = 1
    While Not fldata3.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata3!id
        .TextMatrix(.Rows - 1, 1) = fldata3!Name
        .TextMatrix(.Rows - 1, 2) = fldata3!UserID
        .TextMatrix(.Rows - 1, 3) = fldata3!Department
        .TextMatrix(.Rows - 1, 4) = fldata3!Password
        .TextMatrix(.Rows - 1, 5) = fldata3!Type
        .TextMatrix(.Rows - 1, 6) = fldata3!accessrights
 fldata3.MoveNext
 Wend
 End With
 If fldata3.State Then fldata3.Close
End Sub

Private Sub MSFlexGrid1_Click()
'On Error Resume Next
txtUsername.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
If MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1) <> "" Then Uid = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 1)
'Frame1.Visible = False
End Sub

Private Sub txtConfirmPassword_KeyPress(KeyAscii As Integer)
KeyAscii = TxtAcceptString(Me.txtConfirmPassword, KeyAscii)
End Sub

Private Sub txtNewPassword_KeyPress(KeyAscii As Integer)
KeyAscii = TxtAcceptString(Me.txtNewPassword, KeyAscii)
End Sub

Private Sub txtUsername_Change()
Dim j As Integer
j = 0
If RS.State Then RS.Close
RS.Open "select * from tblusers where name = '" & txtUsername.Text & "'", cn, 3, 2
If Not RS.EOF Then
    txtUsername.Text = RS("name")
    txtNewPassword.Text = RS("password")
    Select Case RS("type")
    Case "Yes"
         optAdministrator.Value = True
    Case "No"
         optGeneralUser.Value = True
    End Select
   For i = 1 To lstsecurity.ListCount - 1
         lstsecurity.Selected(i) = False
    Next
    While j < Len(RS("Accessrights"))
        j = j + 1
        If (Mid(RS("Accessrights"), j, 1)) = "1" Then
            lstsecurity.Selected(j - 1) = True
        End If
    Wend
End If
End Sub

Private Sub txtUsername_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    SearchString "select name from tblusers", Me.txtUsername
End If
End Sub

Public Sub security()
listval = ""
  For i = 0 To lstsecurity.ListCount - 1
  If lstsecurity.Selected(i) Then
     listval = listval & "1"
  Else
     listval = listval & "0"
  End If
  Next i
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
KeyAscii = TxtAcceptString(Me.txtUsername, KeyAscii)
End Sub

