VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPOEntry 
   Caption         =   "Enter Purchase Order"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDataEntry.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   10905
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6765
      TabIndex        =   25
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtEditFlexMatrix 
      Height          =   495
      Left            =   4680
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flxPODetails 
      Height          =   2655
      Left            =   525
      TabIndex        =   23
      Top             =   5640
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   4683
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9645
      TabIndex        =   22
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8205
      TabIndex        =   21
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Purchase Order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   525
      TabIndex        =   0
      Top             =   240
      Width           =   10335
      Begin VB.ComboBox cboChargeType 
         Height          =   315
         ItemData        =   "frmDataEntry.frx":0442
         Left            =   7080
         List            =   "frmDataEntry.frx":0455
         TabIndex        =   31
         Text            =   "NA"
         Top             =   3840
         Width           =   735
      End
      Begin VB.ComboBox cboDeliveryPoint 
         Height          =   315
         Left            =   7080
         TabIndex        =   29
         Top             =   3300
         Width           =   3135
      End
      Begin VB.ComboBox cboVendorName 
         Height          =   315
         Left            =   1200
         TabIndex        =   28
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox txtCurrency 
         Height          =   375
         Left            =   7080
         TabIndex        =   26
         Text            =   "RM"
         Top             =   2640
         Width           =   495
      End
      Begin MSComCtl2.DTPicker dtpPODate 
         Height          =   375
         Left            =   7080
         TabIndex        =   19
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16384001
         CurrentDate     =   39191
      End
      Begin VB.TextBox txtDeliveryTerms 
         Height          =   1095
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   3240
         Width           =   4575
      End
      Begin VB.TextBox txtRemarks 
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Top             =   2760
         Width           =   4575
      End
      Begin VB.TextBox txtApprovedBy 
         Height          =   375
         Left            =   7080
         TabIndex        =   14
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox txtOrderedBy 
         Height          =   375
         Left            =   7080
         TabIndex        =   12
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtCostCenter 
         Height          =   375
         Left            =   7080
         TabIndex        =   10
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtReqOrderNo 
         Height          =   375
         Left            =   7080
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtPONumber 
         Height          =   375
         Left            =   1200
         TabIndex        =   5
         Top             =   2160
         Width           =   4575
      End
      Begin VB.TextBox txtAddress 
         Height          =   1215
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Charge Type"
         Height          =   195
         Left            =   6000
         TabIndex        =   32
         Top             =   3900
         Width           =   930
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Delivery Point"
         Height          =   195
         Left            =   6000
         TabIndex        =   30
         Top             =   3360
         Width           =   990
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Currency"
         Height          =   195
         Left            =   6000
         TabIndex        =   27
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Delivery Terms"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   3360
         Width           =   1065
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Approved By"
         Height          =   195
         Left            =   6000
         TabIndex        =   13
         Top             =   2280
         Width           =   930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Ordered By"
         Height          =   195
         Left            =   6000
         TabIndex        =   11
         Top             =   1800
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cost Center"
         Height          =   195
         Left            =   6000
         TabIndex        =   9
         Top             =   1320
         Width           =   870
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Req. Order # "
         Height          =   195
         Left            =   6000
         TabIndex        =   7
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         Height          =   195
         Left            =   6000
         TabIndex        =   6
         Top             =   360
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PO. Number"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vendor Name"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.Label lblPOtotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10650
      TabIndex        =   33
      Top             =   8400
      Width           =   90
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Purchase Order Details"
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
      Left            =   525
      TabIndex        =   20
      Top             =   5280
      Width           =   1935
   End
End
Attribute VB_Name = "frmPOEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bookmark As Double
Dim addflag As String
Dim i, newRecordID As Integer
Dim rsVendors As New ADODB.Recordset
Dim rsDeliveryPoints As New ADODB.Recordset
Dim dbltotal As Double
Dim verify As Boolean
Private Sub LoadVendors()
cboVendorName.Clear
If rsVendors.State Then RS.Close
rsVendors.Open "Select v_name from Vendor order by v_name", cn, 3, 2
While Not rsVendors.EOF
cboVendorName.AddItem rsVendors(0)
rsVendors.MoveNext
Wend
If rsVendors.State Then rsVendors.Close
End Sub
Private Sub LoadDeliveryPoints()
cboDeliveryPoint.Clear
If rsDeliveryPoints.State Then RS.Close
rsDeliveryPoints.Open "Select DeliveryPointName from DeliveryPoints ", cn, 3, 2
While Not rsDeliveryPoints.EOF
cboDeliveryPoint.AddItem rsDeliveryPoints(0)
rsDeliveryPoints.MoveNext
Wend
If rsDeliveryPoints.State Then rsDeliveryPoints.Close
End Sub
Function GetVendorAddress()
If Len(cboVendorName.Text) > 0 Then
If rsVendors.State Then RS.Close
rsVendors.Open "Select v_address,v_phone,v_fax,v_contactperson   from Vendor where v_name = '" & cboVendorName.Text & "'", cn, 3, 2
If Not rsVendors.EOF Then
txtAddress.Text = rsVendors(0) & vbNewLine & "TEL.:" & rsVendors(1) & vbNewLine & "FAX:" & rsVendors(2) & vbNewLine & "ATTN:" & rsVendors(3)
End If
End If
If rsVendors.State Then rsVendors.Close
End Function
Function GetDeliveryPointInfo()
If Len(cboDeliveryPoint.Text) > 0 Then
If rsDeliveryPoints.State Then RS.Close
rsDeliveryPoints.Open "Select DeliveryPointAddress  from DeliveryPoints where DeliveryPointName = '" & cboDeliveryPoint.Text & "'", cn, 3, 2
If Not rsDeliveryPoints.EOF Then
txtDeliveryTerms.Text = cboDeliveryPoint.Text & vbNewLine & "Address.:" & rsDeliveryPoints(0)
End If
End If
If rsDeliveryPoints.State Then rsDeliveryPoints.Close
End Function
Private Sub cboDeliveryPoint_Change()
GetDeliveryPointInfo
End Sub
Private Sub cboDeliveryPoint_Click()
GetDeliveryPointInfo
End Sub
Private Sub cboVendorName_Change()
txtAddress.Text = ""
GetVendorAddress
End Sub
Private Sub cboVendorName_Click()
txtAddress.Text = ""
GetVendorAddress
End Sub
Private Function ex_validate_fields() As Boolean
    verify = False
If cboVendorName.Text = "" Then
    MsgBox "Please Choose the Vendor's Name", vbExclamation
        ex_validate_fields = False
ElseIf txtAddress.Text = "" Then
    MsgBox "Please Enter the Vendor's Address!", vbExclamation
        ex_validate_fields = False
ElseIf txtPONumber.Text = "" Then
    MsgBox "Please Enter the PO Number!", vbExclamation
        ex_validate_fields = False
ElseIf txtRemarks.Text = "" Then
    MsgBox "Please Enter the Remarks / Project Information!", vbExclamation
        ex_validate_fields = False
ElseIf txtDeliveryTerms.Text = "" Then
    MsgBox "Please Enter the Delivery Terms Information!", vbExclamation
        ex_validate_fields = False
ElseIf txtReqOrderNo.Text = "" Then
    MsgBox "Please Enter the Request Order Number!", vbExclamation
        ex_validate_fields = False
ElseIf txtCostCenter.Text = "" Then
    MsgBox "Please Enter the Cost Center!", vbExclamation
        ex_validate_fields = False
ElseIf txtOrderedBy.Text = "" Then
    MsgBox "Please Enter Ordered by!", vbExclamation
        ex_validate_fields = False
ElseIf txtApprovedBy.Text = "" Then
    MsgBox "Please Enter the Approving Authority!", vbExclamation
        ex_validate_fields = False
ElseIf txtCurrency.Text = "" Then
    MsgBox "Please Enter the Currency!", vbExclamation
        ex_validate_fields = False
ElseIf cboDeliveryPoint.Text = "" Then
    MsgBox "Please Choose the Delivery Point Information!", vbExclamation
        ex_validate_fields = False
ElseIf flxPODetails.TextMatrix(1, 0) = "" Then
    MsgBox "Please Enter atleast one Item to Enter!", vbExclamation
        ex_validate_fields = False
    Else
        ex_validate_fields = True
    End If
End Function
Private Sub cmdSave_Click()
On Error Resume Next
If ex_validate_fields = False Then Exit Sub
Dim alert As String
alert = MsgBox("Do you want to Save this record?", vbOKCancel)
If alert = vbOK Then
    Dim RsAdd As New ADODB.Recordset
    If RsAdd.State Then RsAdd.Close
    RsAdd.Open "Select * from PO_Header", cn, 3, 2
        RsAdd.AddNew
            RsAdd!VendorName = cboVendorName.Text
            RsAdd!VendorAddress = txtAddress.Text
            RsAdd!PO_OrderNo = txtPONumber.Text
            RsAdd!PO_Date = Format(dtpPODate.Value, "MM/dd/yyyy")
            RsAdd!Requisition_OrderNo = txtReqOrderNo.Text
            RsAdd!CostCenter = txtCostCenter.Text
            RsAdd!Orderedby = txtOrderedBy.Text
            RsAdd!Approvedby = txtApprovedBy.Text
            RsAdd!Remarks = txtRemarks.Text
            RsAdd!DeliveryTerms = txtDeliveryTerms.Text
            RsAdd!Currency = txtCurrency.Text
            RsAdd.Update
            Bookmark = RsAdd.AbsolutePosition  ' First, store the location of the cursor
            RsAdd.Requery
            RsAdd.AbsolutePosition = Bookmark
            newRecordID = RsAdd("POID")
            If RsAdd.State Then RsAdd.Close
            'Save the Detail
            For i = 1 To flxPODetails.Rows - 1
            If flxPODetails.TextMatrix(i, 0) <> "" Then
            RsAdd.Open "Select * from PO_Details", cn, 3, 2
            RsAdd.AddNew
            RsAdd!POID = newRecordID
            RsAdd!ItemDescription = flxPODetails.TextMatrix(i, 0)
            RsAdd!quantity = flxPODetails.TextMatrix(i, 1)
            RsAdd!UOM = flxPODetails.TextMatrix(i, 2)
            RsAdd!AccountCode = flxPODetails.TextMatrix(i, 3)
            RsAdd!UnitPrice = flxPODetails.TextMatrix(i, 4)
            RsAdd.Update '
            If RsAdd.State Then RsAdd.Close
            End If
            Next i
            MsgBox "Purchase Order Saved Successfully", vbInformation, App.Title
            'cmdReset_Click
End If
Xit:
Err.Clear
End Sub
Private Sub Form_Load()
Dim r As Integer
Dim c As Integer
Dim max_len As Single
Dim new_len As Single
dtpPODate.Value = Format(Date, "dd/mm/yyyy")
    ' Use no border.
    txtEditFlexMatrix.BorderStyle = vbBSNone
    ' Match the grid's font.
    txtEditFlexMatrix.FontName = flxPODetails.FontName
    txtEditFlexMatrix.FontSize = flxPODetails.FontSize
    txtEditFlexMatrix.Visible = False
    ' Create some data.
    flxPODetails.FixedCols = 0
    flxPODetails.Cols = 6
    flxPODetails.FixedRows = 1
    flxPODetails.Rows = 10
    flxPODetails.TextMatrix(0, 0) = "Item Description"
    flxPODetails.ColWidth(0) = 3500
    flxPODetails.TextMatrix(0, 1) = "Quantity"
    flxPODetails.ColWidth(1) = 1000
    flxPODetails.TextMatrix(0, 2) = "UOM"
    flxPODetails.ColWidth(2) = 600
    flxPODetails.TextMatrix(0, 3) = "Account Code"
    flxPODetails.ColWidth(3) = 1200
    flxPODetails.TextMatrix(0, 4) = "Unit Price"
    flxPODetails.ColWidth(4) = 1300
    flxPODetails.TextMatrix(0, 5) = "Total"
    flxPODetails.ColWidth(4) = 1300
LoadDeliveryPoints
LoadVendors
End Sub
Private Sub GridEdit(KeyAscii As Integer)
    ' Position the TextBox over the cell.
    txtEditFlexMatrix.Left = flxPODetails.CellLeft + flxPODetails.Left
    txtEditFlexMatrix.Top = flxPODetails.CellTop + flxPODetails.Top
    txtEditFlexMatrix.Width = flxPODetails.CellWidth
    txtEditFlexMatrix.Height = flxPODetails.CellHeight
    txtEditFlexMatrix.Visible = True
    txtEditFlexMatrix.SetFocus
    Select Case KeyAscii
        Case 0 To Asc(" ")
            txtEditFlexMatrix.Text = flxPODetails.Text
            txtEditFlexMatrix.SelStart = Len(txtEditFlexMatrix.Text)
        Case Else
            txtEditFlexMatrix.Text = Chr$(KeyAscii)
            txtEditFlexMatrix.SelStart = 1
    End Select
End Sub
Private Sub txtEditFlexMatrix_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            txtEditFlexMatrix.Visible = False
            flxPODetails.SetFocus

        Case vbKeyReturn
            ' Finish editing.
            flxPODetails.SetFocus
        Case vbKeyDown
            ' Move down 1 row.
            flxPODetails.SetFocus
            DoEvents
            If flxPODetails.Row < flxPODetails.Rows - 1 Then
                flxPODetails.Row = flxPODetails.Row + 1
            End If
        Case vbKeyUp
            ' Move up 1 row.
            flxPODetails.SetFocus
            DoEvents
            If flxPODetails.Row > flxPODetails.FixedRows Then
                flxPODetails.Row = flxPODetails.Row - 1
            End If
            If flxPODetails.Col = 4 And flxPODetails.TextMatrix(flxPODetails.Row, 1) <> "" Then flxPODetails.TextMatrix(flxPODetails.Row, 5) = flxPODetails.TextMatrix(flxPODetails.Row, 1) * flxPODetails.TextMatrix(flxPODetails.Row, 4)
            
    End Select
End Sub
Private Function ComputeTotal()
dbltotal = 0
For i = 1 To flxPODetails.Rows - 1
        dbltotal = IIf(flxPODetails.TextMatrix(i, 5) = "", 0, flxPODetails.TextMatrix(i, 5)) + dbltotal
Next i
If dbltotal > 0 Then lblPOtotal.Caption = "Total Value of Purchase Order = " & txtCurrency.Text & ". " & Format(dbltotal, "###,###,##0.00") Else lblPOtotal.Caption = ""
End Function
' Do not beep on Return or Escape.
Private Sub txtEditFlexMatrix_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
                  Select Case flxPODetails.Col
                  Case 0, 2, 3:
            KeyAscii = TxtAcceptString(Me.txtEditFlexMatrix, KeyAscii)
                  Case 1, 4:
            KeyAscii = TxtAcceptMoney(Me.txtEditFlexMatrix, KeyAscii)
            If flxPODetails.Col = 4 And flxPODetails.TextMatrix(flxPODetails.Row, 1) <> "" And flxPODetails.TextMatrix(flxPODetails.Row, 4) <> "" Then
            flxPODetails.TextMatrix(flxPODetails.Row, 5) = Format(flxPODetails.TextMatrix(flxPODetails.Row, 1) * flxPODetails.TextMatrix(flxPODetails.Row, 4), "###,###,##0.00")
            ComputeTotal
            Else
            flxPODetails.TextMatrix(flxPODetails.Row, 5) = 0
            ComputeTotal
            End If
            Case Else
            KeyAscii = 0
            End Select
End Sub
Private Sub flxPODetails_DblClick()
    GridEdit Asc(" ")
End Sub
Private Sub flxPODetails_KeyPress(KeyAscii As Integer)
    GridEdit KeyAscii
End Sub
Private Sub flxPODetails_LeaveCell()
    If txtEditFlexMatrix.Visible Then
        flxPODetails.Text = txtEditFlexMatrix.Text
        txtEditFlexMatrix.Visible = False
    End If
End Sub
Private Sub flxPODetails_GotFocus()
    If txtEditFlexMatrix.Visible Then
        flxPODetails.Text = txtEditFlexMatrix.Text
        txtEditFlexMatrix.Visible = False
    End If
End Sub


