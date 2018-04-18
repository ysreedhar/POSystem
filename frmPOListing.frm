VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPOListing 
   Caption         =   "Purchase Order Listing"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
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
   ScaleHeight     =   8910
   ScaleWidth      =   11850
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkShowCancelled 
      Caption         =   "Show Cancelled Purchase Orders"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid flxPODetails 
      Height          =   3255
      Left            =   480
      TabIndex        =   1
      Top             =   5280
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid flxPurchaseOrders 
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   1
      Cols            =   13
      AllowUserResizing=   1
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   15
      Left            =   7080
      TabIndex        =   6
      Top             =   5280
      Width           =   1215
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
      Left            =   14760
      TabIndex        =   5
      Top             =   4800
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Purchase Order Details"
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
      Left            =   480
      TabIndex        =   3
      Top             =   4800
      Width           =   3270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Purchase Orders"
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
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   2340
   End
End
Attribute VB_Name = "frmPOListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function RepaintFlexGrid()
' Reset the backcolor
For ch = 1 To flxPurchaseOrders.Rows - 1
For flxcls = 0 To flxPurchaseOrders.Cols - 1
    flxPurchaseOrders.Row = ch
    flxPurchaseOrders.Col = flxcls
If flxPurchaseOrders.CellBackColor = vbYellow Then flxPurchaseOrders.CellBackColor = vbWhite
Next flxcls

Next ch
End Function

Private Sub chkShowCancelled_Click()
If chkShowCancelled.Value Then
LoadPurchaseOrders (True)
End If
End Sub


Private Sub flxPurchaseOrders_Click()
current = flxPurchaseOrders.Row
RepaintFlexGrid
'Current  row
flxPurchaseOrders.Row = current
For i = 1 To flxPurchaseOrders.Cols - 1
flxPurchaseOrders.Col = i
flxPurchaseOrders.CellBackColor = vbYellow
Next
LoadPurchaseOrderDetails
vprev = flxPurchaseOrders.Row
flxPurchaseOrders.Col = 1
End Sub

Private Sub Form_Load()
LoadFlexTitles
LoadPurchaseOrders (False)
End Sub
Private Sub LoadFlexTitles()
On Error Resume Next
    With flxPurchaseOrders
        .Row = 0:    .Col = 0
        .ColWidth(0) = 400
       .TextMatrix(0, 0) = "POID"
        .ColWidth(1) = 500
       .TextMatrix(0, 1) = "VendorName"
        .ColWidth(2) = 3000
        .TextMatrix(0, 2) = "VendorAddress"
        .ColWidth(3) = 3300
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Order #"
        .ColWidth(3) = 1100
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Date"
        .ColWidth(5) = 1600
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "Req. Order#"
        .ColWidth(5) = 1000
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "Cost Center"
        .ColWidth(7) = 1000
        .TextMatrix(0, 7) = "Ordered by"
        .ColWidth(7) = 2000
        .TextMatrix(0, 8) = "Approved by"
        .ColWidth(8) = 2000
        .TextMatrix(0, 9) = "Remarks"
        .ColWidth(9) = 2650
        .TextMatrix(0, 10) = "Del. Terms"
        .ColWidth(11) = 1000
        .TextMatrix(0, 11) = "Currency"
        .ColWidth(11) = 600
        .TextMatrix(0, 12) = "Status"
        .ColWidth(12) = 600
    End With
    
    With flxPODetails
        .Row = 0:    .Col = 0
        .ColWidth(0) = 100
    flxPODetails.TextMatrix(0, 1) = "Item Description"
    flxPODetails.ColWidth(1) = 3500
    flxPODetails.TextMatrix(0, 2) = "Quantity"
    flxPODetails.ColWidth(2) = 1000
    flxPODetails.TextMatrix(0, 3) = "UOM"
    flxPODetails.ColWidth(3) = 600
    flxPODetails.TextMatrix(0, 4) = "Account Code"
    flxPODetails.ColWidth(4) = 1200
    flxPODetails.TextMatrix(0, 5) = "Unit Price"
    flxPODetails.ColWidth(5) = 1300
    flxPODetails.TextMatrix(0, 6) = "Total"
    flxPODetails.ColWidth(6) = 1300
    End With
End Sub

Private Sub LoadPurchaseOrders(boolShowCancelled As Boolean)
Dim fldata3 As New ADODB.Recordset
If fldata3.State Then fldata3.Close
If boolShowCancelled = True Then
fldata3.Open "select * from PO_header", cn, 3, 2
Else
fldata3.Open "select * from PO_header where status <> 6", cn, 3, 2
End If
With flxPurchaseOrders
.Rows = 1
    While Not fldata3.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata3!POID
        .TextMatrix(.Rows - 1, 1) = fldata3!VendorName
        .TextMatrix(.Rows - 1, 2) = Replace(fldata3!VendorAddress, vbNewLine, "")
        .TextMatrix(.Rows - 1, 3) = fldata3!PO_OrderNo
        .TextMatrix(.Rows - 1, 4) = fldata3!PO_Date
        .TextMatrix(.Rows - 1, 5) = fldata3!Requisition_OrderNo
        .TextMatrix(.Rows - 1, 6) = fldata3!CostCenter
        .TextMatrix(.Rows - 1, 7) = fldata3!Orderedby
        .TextMatrix(.Rows - 1, 8) = fldata3!Approvedby
        .TextMatrix(.Rows - 1, 9) = fldata3!Remarks
        .TextMatrix(.Rows - 1, 10) = fldata3!DeliveryTerms
        .TextMatrix(.Rows - 1, 11) = fldata3!Currency
        .TextMatrix(.Rows - 1, 12) = fldata3!Status
 fldata3.MoveNext
 Wend
 End With
 If fldata3.State Then fldata3.Close
End Sub

Private Sub LoadPurchaseOrderDetails()
Dim fldata3 As New ADODB.Recordset
If fldata3.State Then fldata3.Close
Dim intPOID As Integer
If flxPurchaseOrders.Row > 0 Then intPOID = flxPurchaseOrders.TextMatrix(flxPurchaseOrders.Row, 0)
fldata3.Open "select * from PO_Details where POID =" & intPOID, cn, 3, 2
With flxPODetails
.Rows = 1
    While Not fldata3.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = .RowPos(.Rows - 1)
        .TextMatrix(.Rows - 1, 1) = fldata3!ItemDescription
        .TextMatrix(.Rows - 1, 2) = fldata3!quantity
        .TextMatrix(.Rows - 1, 3) = fldata3!UOM
        .TextMatrix(.Rows - 1, 4) = fldata3!AccountCode
        .TextMatrix(.Rows - 1, 5) = fldata3!UnitPrice
        .TextMatrix(.Rows - 1, 6) = fldata3!quantity * fldata3!UnitPrice
        dbltotal = fldata3!quantity * fldata3!UnitPrice + dbltotal
 fldata3.MoveNext
 Wend
 End With
 If fldata3.State Then fldata3.Close
If dbltotal > 0 Then lblPOtotal.Caption = "Total Value of Purchase Order = " & flxPurchaseOrders.TextMatrix(flxPurchaseOrders.Row, 11) & "." & Format(dbltotal, "###,###,##0.00") Else lblPOtotal.Caption = ""
End Sub
