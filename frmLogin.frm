VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PO System Login"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
   FillColor       =   &H8000000F&
   ForeColor       =   &H8000000F&
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   154
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   343
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   285
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox cmbName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Top             =   435
      Width           =   2655
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H8000000A&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3390
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1740
      Width           =   1215
   End
   Begin VB.CommandButton cmdContinue 
      BackColor       =   &H8000000A&
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2115
      MaskColor       =   &H0080C0FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1740
      Width           =   1215
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1935
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   750
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   750
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Private Sub cmbName_KeyPress(KeyAscii As Integer)
KeyAscii = TxtAcceptString(Me.cmbName, KeyAscii)
End Sub

Private Sub cmdContinue_click()
    If (cmbName.Text = "") Or (txtPassword.Text = "") Then
        MsgBox ("You cannot leave username or password as blank"), vbCritical
        If cmbName.Text = "" Then
            cmbName.SetFocus
        Else
            txtPassword.SetFocus
        End If
    Else
    If rs1.State = 1 Then rs1.Close

    rs1.Open "Select * from tblusers where UserID='" & cmbName.Text & "' and password='" & txtPassword.Text & "'", cn, 1, 1
       If Not rs1.EOF Then
            
            If rs1("type") = "Yes" Then
                Type_doctor = True
            End If
           
            Username = cmbName.Text
            Pass = txtPassword.Text
            pubUserName = rs1("Name")
              pubUserDept = rs1("Department")
              
            '  MDIMain.Administration_User.Visible = UCase(Username) = "ADMIN"
            MDIMain.mnuAdministration.Visible = Mid(rs1("accessrights"), 1, 1) = "1"
            MDIMain.mnuTransactions.Visible = Mid(rs1("accessrights"), 2, 1) = "1"
            MDIMain.mnuInvoiceTransactions.Visible = Mid(rs1("accessrights"), 3, 1) = "1"
                       
            varAccessStr = "logged In"
            Unload Me
            
            MDIMain.Show
        Else
            MsgBox ("You have entered wrong password"), vbCritical
            txtPassword.Text = ""
            txtPassword.SetFocus
        End If
    End If

'Patient Payables
strRetrieve = "FALSE"
End Sub

Private Sub cmdExit_Click()
'On Error Resume Next
    'Unload frmLogin
    End
End Sub

Private Sub Form_activate()
Me.Left = (Screen.Width - Me.ScaleWidth) \ 2 - 2000
Me.Top = (Screen.Height - Me.ScaleHeight) \ 2 - 2200
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = "27" Then Call cmdExit_Click
End Sub

Private Sub Form_Load()
GetConnectionString
If Trim(vPubConnectString) = "" Then
cdlg.Filter = "Microsoft Access Files(*.Mdb)|*.Mdb"
cdlg.ShowOpen
vPubConnectString = cdlg.FileName
If vPubConnectString = "" Then End: Exit Sub
Dim SectionName As String
Dim VariableName As String
Dim ValueFrom As String
Dim iret As Long
ValueFrom = String(256, 0)
ValueFrom = vPubConnectString
WritePrivateProfileString "DB", "Server", ValueFrom, App.Path & "\ServerSetup.Ini"
End If
DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=" & vPubConnectString
Call ModSoho01.connect
Dim RS As New ADODB.Recordset
'rs.Open "Select flashmsg from company ", cn, 1, 1
'If rs.EOF Then pubFlashMsg = "RanteQ Technology" Else pubFlashMsg = rs(0)
'rs.Close
GetCompanyInformation
Set rs1 = New ADODB.Recordset
Set RS = New ADODB.Recordset
RS.ActiveConnection = cn
RS.CursorLocation = adUseClient
RS.CursorType = adOpenDynamic
RS.LockType = adLockOptimistic
rs1.ActiveConnection = cn
rs1.CursorLocation = adUseClient
rs1.CursorType = adOpenDynamic
rs1.LockType = adLockOptimistic
End Sub
Public Sub Password()
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select password, type, Name,Department from username where UserID='" & cmbName.Text & "'"
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call cmdContinue_click: Exit Sub
KeyAscii = TxtAcceptString(Me.txtPassword, KeyAscii)
End Sub
