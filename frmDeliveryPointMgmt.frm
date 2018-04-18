VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDeliveryPointMgmt 
   Caption         =   "Delivery Point Management"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6675
   ScaleWidth      =   10050
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   7320
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   7320
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   7320
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtDPAddress 
      Height          =   1095
      Left            =   2160
      TabIndex        =   2
      Text            =   " "
      Top             =   960
      Width           =   4215
   End
   Begin VB.TextBox txtDPName 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   4215
   End
   Begin MSFlexGridLib.MSFlexGrid flxDeliveryPoints 
      Height          =   3615
      Left            =   480
      TabIndex        =   0
      Top             =   2640
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6376
      _Version        =   393216
      Cols            =   3
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Existing Delivery Points"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   2280
      Width           =   1980
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Delivery Point"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   1185
   End
End
Attribute VB_Name = "frmDeliveryPointMgmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flxDPRowID As Integer

Private Sub cmdClear_Click()
flxDPRowID = 0
txtDPName.Text = ""
txtDPAddress.Text = ""
End Sub

Private Sub cmdDelete_Click()
response = MsgBox("Are you sure you want to delete this Delivery Point?", vbOKCancel, App.Title)
If response = 1 Then
Dim rsDeliveryPoint As New ADODB.Recordset
rsDeliveryPoint.Open "Select * from DeliveryPoints Where DeliveryPointID=" & flxDPRowID, cn, adOpenKeyset, adLockReadOnly
If Not rsDeliveryPoint.EOF Then
    vsqlstr = "Delete from deliverypoints WHERE DeliveryPointID= " & flxDPRowID
    cn.Execute vsqlstr
    MsgBox "Record Deleted", vbInformation, App.Title
LoadDeliveryPoints
cmdClear_Click
End If
End If
End Sub

Private Sub cmdSave_Click()
If Trim(txtDPName.Text) = "" Then
MsgBox "Choose a Name for Delivery Point!", vbInformation, App.Title
Exit Sub
End If
If Trim(txtDPAddress.Text) = "" Then
MsgBox "Enter Address for Delivery Point", vbInformation, App.Title
Exit Sub
End If
response = MsgBox("Are you sure you want to add this Delivery Point?", vbOKCancel, App.Title)
If response = 1 Then
Dim rsDeliveryPoint As New ADODB.Recordset
rsDeliveryPoint.Open "Select * from DeliveryPoints Where DeliveryPointID=" & flxDPRowID, cn, adOpenKeyset, adLockReadOnly
If rsDeliveryPoint.EOF Then
'AddNew
    vsqlstr = "Insert Into DeliveryPoints (" & _
    "DeliveryPointName," & _
    "DeliveryPointAddress)" & _
    " Values (" & _
    "'" & Trim(txtDPName.Text) & "', '" & _
    Trim(txtDPAddress.Text) & "')"
    
    cn.Execute vsqlstr
    MsgBox "Record Added", vbInformation, App.Title
Else
'Update
vsqlstr = "Update DeliveryPoints SET " & _
    "DeliveryPointName= '" & Trim(txtDPName.Text) & _
    "', DeliveryPointAddress= '" & Trim(txtDPAddress.Text) & "' WHERE DeliveryPointID= " & flxDPRowID
    cn.Execute vsqlstr
    MsgBox "Record Updated", vbInformation, App.Title
End If

If rsDeliveryPoint.State Then rsDeliveryPoint.Close
Set rsDeliveryPoint = Nothing
LoadDeliveryPoints
cmdClear_Click
End If
End Sub

Private Sub flxDeliveryPoints_Click()
current = flxDeliveryPoints.Row
RepaintFlexGrid
'Current  row
flxDeliveryPoints.Row = current
For i = 1 To flxDeliveryPoints.Cols - 1
flxDeliveryPoints.Col = i
flxDeliveryPoints.CellBackColor = vbYellow
Next
flxDPRowID = flxDeliveryPoints.TextMatrix(flxDeliveryPoints.Row, 0)
txtDPName.Text = flxDeliveryPoints.TextMatrix(flxDeliveryPoints.Row, 1)
txtDPAddress.Text = flxDeliveryPoints.TextMatrix(flxDeliveryPoints.Row, 2)
vprev = flxDeliveryPoints.Row
flxDeliveryPoints.Col = 1
End Sub
Function RepaintFlexGrid()
' Reset the backcolor
For ch = 1 To flxDeliveryPoints.Rows - 1
For flxcls = 0 To flxDeliveryPoints.Cols - 1
    flxDeliveryPoints.Row = ch
    flxDeliveryPoints.Col = flxcls
If flxDeliveryPoints.CellBackColor = vbYellow Then flxDeliveryPoints.CellBackColor = vbWhite
Next flxcls
Next ch
End Function
Private Sub Form_Load()
LoadFlexTitles
LoadDeliveryPoints
End Sub

Private Sub LoadDeliveryPoints()
Dim fldata3 As New ADODB.Recordset
If fldata3.State Then fldata3.Close
fldata3.Open "select * from DeliveryPoints", cn, 3, 2
With flxDeliveryPoints
.Rows = 1
    While Not fldata3.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata3!DeliveryPointID
        .TextMatrix(.Rows - 1, 1) = fldata3!DeliveryPointName
        .TextMatrix(.Rows - 1, 2) = fldata3!DeliveryPointAddress
 fldata3.MoveNext
 Wend
 End With
 If fldata3.State Then fldata3.Close
End Sub
Private Sub LoadFlexTitles()
On Error Resume Next
    With flxDeliveryPoints
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
    .TextMatrix(0, 1) = "Name"
    .ColWidth(1) = 3500
     .TextMatrix(0, 2) = "Address"
     .ColWidth(2) = 5000
    End With
End Sub
