VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmInvoiceLog 
   Caption         =   "Invoice Log"
   ClientHeight    =   8370
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
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
   ScaleHeight     =   8370
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtEditFlexMatrix 
      Height          =   495
      Left            =   960
      TabIndex        =   15
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpInvoiceDate 
      Height          =   315
      Left            =   2520
      TabIndex        =   10
      Top             =   1440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   39210
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogInvoice 
      Caption         =   "Log Invoice"
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtInvoiceNumber 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   2535
   End
   Begin VB.ComboBox cboPurchaseOrders 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   300
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid flxPODetails 
      Height          =   3255
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      AllowUserResizing=   1
   End
   Begin MSComCtl2.DTPicker dtpReceiptDate 
      Height          =   315
      Left            =   7200
      TabIndex        =   12
      Top             =   1440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   39210
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "PO Value"
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
      Left            =   5280
      TabIndex        =   18
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblPOValue 
      AutoSize        =   -1  'True
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
      Left            =   7200
      TabIndex        =   17
      Top             =   840
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Modify Inv. Quantity to match the delivery as per Invoice."
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   4200
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Vendor Name"
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
      Left            =   5280
      TabIndex        =   14
      Top             =   360
      Width           =   1125
   End
   Begin VB.Label lblVendorName 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   7200
      TabIndex        =   13
      Top             =   360
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Invoice Receipt Date"
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
      Left            =   5280
      TabIndex        =   11
      Top             =   1500
      Width           =   1770
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Invoice Date"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1500
      Width           =   1080
   End
   Begin VB.Label lblInvoiceValue 
      AutoSize        =   -1  'True
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
      Left            =   5280
      TabIndex        =   6
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Match Items on Purchase Order"
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
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   2685
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Invoice Number"
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
      Left            =   240
      TabIndex        =   2
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Purchase Order Number"
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
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2010
   End
End
Attribute VB_Name = "FrmInvoiceLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strChecked = "þ"
Const strUnChecked = "q"
Dim strInvoiceCurrency As String
Dim rsLogInvoice As New ADODB.Recordset
Dim Finaltotal As Double
Private Sub cboPurchaseOrders_Change()
LoadPurchaseOrderDetails
End Sub
Private Sub cboPurchaseOrders_Click()
LoadPurchaseOrderDetails
End Sub
Private Sub cmdLogInvoice_Click()
response = MsgBox("Are you sure you want Log this Invoice?", vbYesNo, App.Title)
If response = vbYes Then
If rsLogInvoice.State Then rsLogInvoice.Close
 rsLogInvoice.Open "Select * from tblInvoices", cn, 3, 2
            rsLogInvoice.AddNew
            rsLogInvoice!InvoiceNumber = txtInvoiceNumber.Text
            rsLogInvoice!Invoice_Value = Finaltotal
            rsLogInvoice!PO_Number = cboPurchaseOrders.Text
            rsLogInvoice!InvoiceDate = Format(dtpInvoiceDate.Value, "dd/MM/yyyy")
            rsLogInvoice!ReceiptDate = Format(dtpReceiptDate.Value, "dd/MM/yyyy")
            rsLogInvoice.Update '
            If rsLogInvoice.State Then rsLogInvoice.Close
MsgBox "Invoice Logged Successfully!", vbInformation, App.Title
cmdReset_Click
End If
End Sub
Private Sub cmdReset_Click()
flxPODetails.Rows = 1
cboPurchaseOrders.Text = ""
txtInvoiceNumber.Text = ""
lblInvoiceValue.Caption = ""
dtpInvoiceDate.Value = Date
dtpReceiptDate.Value = Date
End Sub
Private Sub Form_Load()
dtpInvoiceDate.Value = Date
dtpReceiptDate.Value = Date
LoadPos
    With flxPODetails
        .Row = 0:    .Col = 0
        .ColWidth(0) = 100
        .TextMatrix(0, 1) = "Select"
        .ColWidth(1) = 500
    flxPODetails.TextMatrix(0, 2) = "Item Description"
    flxPODetails.ColWidth(2) = 3500
    flxPODetails.TextMatrix(0, 3) = "Req. Quantity"
    flxPODetails.ColWidth(3) = 1000
    flxPODetails.TextMatrix(0, 4) = "UOM"
    flxPODetails.ColWidth(4) = 600
    flxPODetails.TextMatrix(0, 5) = "Unit Price"
    flxPODetails.ColWidth(5) = 1300
    flxPODetails.TextMatrix(0, 6) = "Inv. Quantity"
    flxPODetails.ColWidth(6) = 1300
    End With
End Sub
Private Sub LoadPos()
Dim rsPos As New ADODB.Recordset
Dim strCommand As String
cboPurchaseOrders.Clear
strCommand = "Select distinct(PO_OrderNo) from PO_header where status = 1 order by po_OrderNo asc"
If rsPos.State Then rsPos.Close
rsPos.Open strCommand, cn, 3, 2
While Not rsPos.EOF
cboPurchaseOrders.AddItem rsPos(0)
rsPos.MoveNext
Wend
End Sub
Private Sub LoadPurchaseOrderDetails()
flxPODetails.Rows = 1
Dim fldata3 As New ADODB.Recordset
If fldata3.State Then fldata3.Close
fldata3.Open "select * from PO_Details where POID in (select POID from PO_header where PO_OrderNo = '" & cboPurchaseOrders.Text & "')", cn, 3, 2
With flxPODetails
.Rows = 1
    While Not fldata3.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = .RowPos(.Rows - 1)
        .TextMatrix(.Rows - 1, 2) = fldata3!ItemDescription
        .TextMatrix(.Rows - 1, 3) = fldata3!quantity
        .TextMatrix(.Rows - 1, 4) = fldata3!UOM
        .TextMatrix(.Rows - 1, 5) = fldata3!UnitPrice
        .TextMatrix(.Rows - 1, 6) = fldata3!quantity
        rowtotal = fldata3!quantity * fldata3!UnitPrice + rowtotal
 fldata3.MoveNext
            'define fields as checkbox
            For Y = 1 To .Rows - 1
                    .Row = Y
                    .Col = 1
                    .CellFontName = "Wingdings"
                    .CellFontSize = 14
                    .CellAlignment = flexAlignCenterCenter
                    .Text = strChecked
            Next Y
 Wend
 End With
 
 If fldata3.State Then fldata3.Close
fldata3.Open "select vendorName, Currency from PO_header where PO_OrderNo = '" & cboPurchaseOrders.Text & "'", cn, 3, 2
If Not fldata3.EOF Then
lblVendorName.Caption = fldata3(0)
strInvoiceCurrency = fldata3(1)
 If fldata3.State Then fldata3.Close
 ComputeTotal
 End If
End Sub

Private Sub txtEditFlexMatrix_KeyDown(KeyCode As Integer, Shift As Integer)
            If flxPODetails.Col < 6 Then Exit Sub
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
    
    End Select
End Sub
' Do not beep on Return or Escape.
Private Sub txtEditFlexMatrix_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
                  Select Case flxPODetails.Col
                  Case 6:
            KeyAscii = TxtAcceptMoney(Me.txtEditFlexMatrix, KeyAscii)
'          If flxPODetails.Col = 6 And flxPODetails.TextMatrix(flxPODetails.Row, 1) <> "" And flxPODetails.TextMatrix(flxPODetails.Row, 4) <> "" Then flxPODetails.TextMatrix(flxPODetails.Row, 5) = Format(flxPODetails.TextMatrix(flxPODetails.Row, 1) * flxPODetails.TextMatrix(flxPODetails.Row, 4), "###,###,##0.00") Else flxPODetails.TextMatrix(flxPODetails.Row, 5) = 0
            ComputeTotal
            Case Else
            KeyAscii = 0
            End Select
End Sub
Private Function ComputeTotal()
dbltotal = 0
For i = 1 To flxPODetails.Rows - 1
        dbltotal = IIf(flxPODetails.TextMatrix(i, 6) = "", 0, (flxPODetails.TextMatrix(i, 5) * flxPODetails.TextMatrix(i, 6))) + dbltotal
Next i
If dbltotal > 0 Then
lblInvoiceValue.Caption = "Total Value of Invoice = " & strInvoiceCurrency & ". " & Format(dbltotal, "###,###,##0.00")
Finaltotal = dbltotal
Else
lblInvoiceValue.Caption = ""
Finaltotal = 0
End If
End Function
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

