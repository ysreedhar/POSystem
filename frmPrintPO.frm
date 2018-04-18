VERSION 5.00
Begin VB.Form frmPrintPO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print PO"
   ClientHeight    =   3840
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrintPO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPrintDraft 
      Caption         =   "Print Draft"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   3615
   End
   Begin VB.CheckBox chkShowPrinted 
      Caption         =   "Show Printed Purchase Orders"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print PO"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox cboPurchaseOrders 
      Height          =   315
      Left            =   2520
      TabIndex        =   0
      Top             =   720
      Width           =   2535
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
      Top             =   750
      Width           =   2010
   End
End
Attribute VB_Name = "frmPrintPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As _
        String, ByVal lpParameters As String, ByVal lpDirectory As _
        String, ByVal nShowCmd As Long) As Long
        Dim i As Integer
        Dim ReportTotal As Double

Private Sub cmdExit_Click()
 Unload Me 'SafeExit
End Sub

Private Sub chkShowPrinted_Click()
If chkShowPrinted.Value Then LoadPos True
End Sub

Private Sub cmdPrint_Click()
GetShapedRS cboPurchaseOrders.Text, False
'rptResults.PrintReport True
'rptResults.Hide
End Sub

Private Sub Form_Load()
LoadPos False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmPOreport = Nothing
End Sub
Private Sub LoadPos(boolShowPrinted As Boolean)
Dim rsPos As New ADODB.Recordset
Dim strCommand As String
cboPurchaseOrders.Clear
If boolShowPrinted = True Then
strCommand = "Select distinct(PO_OrderNo) from PO_header where status in (1,2) order by PO_orderNo"
Else
strCommand = "Select distinct(PO_OrderNo) from PO_header where status = 1 order by po_OrderNo asc"
End If
If rsPos.State Then rsPos.Close
rsPos.Open strCommand, cn, 3, 2
While Not rsPos.EOF
cboPurchaseOrders.AddItem rsPos(0)
rsPos.MoveNext
Wend
End Sub

Public Function GetShapedRS(dblPOnumber As Double, boolPrintPONumber As Boolean)

'Instantiate the Recordset Object
Dim RS As Recordset
Set RS = New Recordset

'Build the Connection String
Dim strConnect As String
strConnect = "Data Provider=Microsoft.Jet.OLEDB.4.0; " & _
             "Provider=MSDataShape;Data Source=" & vPubConnectString & ";"
             
'Build the SQL String for a Shaped recordset
Dim strSQL As String

strSQL = "SHAPE {SELECT * FROM vHeader where PO_OrderNo = '" & dblPOnumber & "' } APPEND " _
         & "({SELECT * FROM vPODetails_report where POID = (select poID from vHeader where PO_OrderNo = '" & dblPOnumber & "' )} AS rsPOItems " _
         & "RELATE POID TO POID)"

    i = 0
With RS
    .ActiveConnection = strConnect
    .Open strSQL, strConnect, adOpenStatic, adLockBatchOptimistic
 
    'Display a populated data report
    Set rptResults.DataSource = RS
    'Display or Hide the PO Number
   
   If boolPrintPONumber = True Then rptResults.Sections.Item("GroupHeader").Controls.Item("txtPONumber").Visible = True
   rptResults.Sections.Item("GroupFooter").Controls.Item("lblTotal").Caption = RS("Currency") & ". " & Format(RS("GrandTotal"), "###,###,##0.00")
   rptResults.Sections.Item("GroupFooter").Controls.Item("lblAmountinwords").Caption = "Amount in Words : " & RS("Currency") & ". " & NumToWords(CDbl(RS("Grandtotal"))) & " Only"

    'Display the report
    rptResults.Show vbModal, frmPrintPO
        
    .Close
End With

'Cleanup
Set RS = Nothing

End Function

'Private Sub PrintTbl(RS, indent)
'
'Dim s As String
'Dim col As Field
'Dim rsChild As Recordset
'
' ' This routine distinguishes between columns in the recordset with
' ' data, i.e. type <> adChapter, and those which contain a child
' ' recordset, for example, type = adChapter.
'  Do While Not RS.EOF
'
'   s = Space(indent)
'   For Each col In RS.Fields
'     If col.Type <> adChapter Then
'       If Len(s) > indent Then s = s & " | "
'       s = s & col.Value
'     Else
'       ' Display data columns encountered so far (if any).
'       If Len(s) > indent Then Debug.Print Space(indent) & s
'       ' Recursively call printtbl to display child recordset.
'       Set rsChild = col.Value
'       PrintTbl rsChild, indent + 4
'       ' Reset in case there are further data columns.
'       s = Space(indent)
'     End If
'   Next
'
'   ' In case we have any data columns that have not been
'   ' displayed yet.
'   If Len(s) > indent Then Debug.Print s
'
'   RS.MoveNext
' Loop
'
'End Sub

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
