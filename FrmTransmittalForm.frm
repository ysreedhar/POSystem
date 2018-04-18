VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FrmTransmittalForm 
   Caption         =   "Transmittal Form"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
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
   ScaleWidth      =   9870
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPreview 
      Caption         =   "P&review"
      Height          =   495
      Left            =   8220
      TabIndex        =   11
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdPageSetup 
      Caption         =   "Page Set&up"
      Height          =   495
      Left            =   6360
      TabIndex        =   10
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtTransmitToDept 
      Height          =   315
      Left            =   1920
      TabIndex        =   7
      Top             =   840
      Width           =   4575
   End
   Begin SHDocVwCtl.WebBrowser rptViewer 
      Height          =   6015
      Left            =   480
      TabIndex        =   6
      Top             =   3480
      Width           =   14295
      ExtentX         =   25215
      ExtentY         =   10610
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   495
      Left            =   4500
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.ListBox lstInvoiceNos 
      BackColor       =   &H00FFFFFF&
      Height          =   1185
      ItemData        =   "FrmTransmittalForm.frx":0000
      Left            =   1920
      List            =   "FrmTransmittalForm.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1320
      Width           =   7620
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View Transmittal Form"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txtTransmitToName 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   360
      TabIndex        =   9
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   900
      Width           =   1035
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Invoice Nos."
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
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Transmit to:"
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "FrmTransmittalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intTransmittalNumber As Integer

Private Function GenerateTransmittalNumber()
intTransmittalNumber = 1
Dim rsTNumber As New ADODB.Recordset
rsTNumber.Open "select max(TransmittalNumber) from tblInvoices", cn, 3, 2
If Not rsTNumber.EOF Then
intTransmittalNumber = rsTNumber(0) + 1
Else
intTransmittalNumber = 1
End If
If rsTNumber.State Then rsTNumber.Close
End Function
Private Sub cmdView_Click()
If Trim(txtTransmitToName.Text) = "" Then
MsgBox "Please Enter Name of the Receiver", vbInformation, App.Title
Exit Sub
End If
If Trim(txtTransmitToDept.Text) = "" Then
MsgBox "Please Enter Department of the Receiver", vbInformation, App.Title
Exit Sub
End If
If lstInvoiceNos.SelCount = 0 Then
MsgBox "Please Choose the Invoices to be Transmitted", vbInformation, App.Title
Exit Sub
End If
Dim fso As Object
Dim WriteObj As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Set WriteObj = fso.CreateTextFile(App.Path & "\report.html", True)

WriteObj.WriteLine " <html>"
WriteObj.WriteLine " <head>"
WriteObj.WriteLine " <Style>"
WriteObj.WriteLine "    .BODY {FONT-FAMILY: Tahoma; FONT-SIZE: 8pt;}"
WriteObj.WriteLine "    .TableFont"
WriteObj.WriteLine "    {"
WriteObj.WriteLine "        COLOR: Black;"
WriteObj.WriteLine "        FONT-FAMILY: Tahoma;"
WriteObj.WriteLine "        FONT-SIZE: 8pt;"
WriteObj.WriteLine "        TEXT-TRANSFORM: capitalize;"
WriteObj.WriteLine "    }"
WriteObj.WriteLine "    .TrFont"
WriteObj.WriteLine "    {"
WriteObj.WriteLine "        COLOR: black;"
WriteObj.WriteLine "        FONT-FAMILY: Tahoma;"
WriteObj.WriteLine "        FONT-SIZE: 8pt;"
WriteObj.WriteLine "        TEXT-TRANSFORM: capitalize;"
WriteObj.WriteLine "        CURSOR:HAND;"
WriteObj.WriteLine "   }"
WriteObj.WriteLine "</style></head>"
WriteObj.WriteLine " <body class=TableFont>"
Dim RS As New ADODB.Recordset
WriteObj.WriteLine " <table border=0 cellspacing=1 cellpadding=2 BORDERCOLOR=GRAY width=95%>"
WriteObj.WriteLine "  <tr bgcolor=white  height=15>"
WriteObj.WriteLine "   <td  colspan=2 align=center valign=top >"
RS.Open "select * from company ", cn, 1, 1
If Not RS.EOF Then
        WriteObj.WriteLine "<FONT size=4> " & RS!CompanyName & "</font><FONT size=2> - [" & RS!regno & "] </FONT>"
        'WriteObj.WriteLine "<tr><TD  align=center><FONT size=2> " & rs!regno & "," & rs!Address & " </FONT></TD></tr>"
        'WriteObj.WriteLine "<tr><td Width=100% align=center><font size=2.5>Phone: " & rs!phone & "</font><font size=2.5>  Fax: " & rs!fax & "</font><font size=2.5> E-mail: " & rs!email & "</td></tr>"
End If
If RS.State Then RS.Close
WriteObj.WriteLine "   </td>"
WriteObj.WriteLine "  </tr>"
WriteObj.WriteLine "  <tr>"
WriteObj.WriteLine "   <td align=center colspan=2>Vendor Invoice  Transmittal Form"
WriteObj.WriteLine "   </td> </tr>"
WriteObj.WriteLine "   <tr><td colspan=2><hr></td></tr>"
WriteObj.WriteLine " <tr><td>To: </br>" & UCase(txtTransmitToName.Text) & " </br> " & UCase(txtTransmitToDept.Text) & "</td>"
WriteObj.WriteLine "   <td valign='top'>   Transmittal No.:" & Format(intTransmittalNumber, "0000000") & "   </td></tr>"
WriteObj.WriteLine " <tr><td valign=top>From:</br>" & UCase(pubUserName) & " </br>" & UCase(pubUserDept) & " </td>"
WriteObj.WriteLine "   <td  valign=top>Date: " & Format(Date, "dd/MM/yyyy") & " </td></tr>"
WriteObj.WriteLine "   <tr><td colspan=2><hr></td></tr>"
WriteObj.WriteLine "  <tr>  <td colspan=2 valign=top >Please receive the following Vendor Invoices:</br></td></tr>"
WriteObj.WriteLine "  <tr><td  colspan=2 valign=top>"
WriteObj.WriteLine "   <table border=1 cellspacing=1 cellpadding=2>"
WriteObj.WriteLine "    <tr><td align='center' nowrap >Sno.</td>"
WriteObj.WriteLine "     <td align='center' nowrap>Vendor</td>"
WriteObj.WriteLine "     <td align='center' nowrap>Invoice No.</td>"
WriteObj.WriteLine "    <td width=100 align='center' nowrap>Value</td>"
WriteObj.WriteLine " <td align='center' nowrap>Currency</td>"
WriteObj.WriteLine "     <td align='center' nowrap >Inv. Date</td>"
WriteObj.WriteLine "     <td align='center' nowrap>Receipt Date</td>"
WriteObj.WriteLine "    <td align='center' nowrap >PO / SO #</td></tr>"
Sno = 1
Dim l As Integer
l = 0
GenerateTransmittalNumber
For l = 0 To lstInvoiceNos.ListCount - 1
If lstInvoiceNos.Selected(l) = True Then
InvSelection = Split(lstInvoiceNos.List(l), " - ", Len(lstInvoiceNos.List(l)), vbTextCompare)
If RS.State Then RS.Close
RS.Open "select * from vInvoiceVendorValue where invoiceNumber = '" & InvSelection(0) & "'", cn, 1, 1
While Not RS.EOF
    WriteObj.WriteLine "    <tr class='trfont'><td>"
    WriteObj.WriteLine "     " & Sno & "</td>"
    WriteObj.WriteLine "     <td nowrap>" & RS("vendorName") & "</td>"
    WriteObj.WriteLine "     <td>" & RS("InvoiceNumber") & "</td>"
    WriteObj.WriteLine "    <td align = 'right'>" & Format(RS("Invoice_Value"), "###,###,##0.00") & "</td>"
    WriteObj.WriteLine " <td>" & RS("Currency") & "</td>"
    WriteObj.WriteLine "     <td>" & Format(RS("InvoiceDate"), "dd/MM/yyyy") & "</td>"
    WriteObj.WriteLine "     <td>" & Format(RS("ReceiptDate"), "dd/MM/yyyy") & "</td>"
    WriteObj.WriteLine "    <td valign=top>" & RS("PO_oRDERNO") & "</td></tr>"
    Sno = Sno + 1
RS.MoveNext
Wend
RS.Close
End If
Next l
WriteObj.WriteLine "     </table>"
WriteObj.WriteLine "     </td> </tr> "
WriteObj.WriteLine "   <tr><td colspan=2><hr></td></tr>"
WriteObj.WriteLine "    <tr><td>Received By:</br>NAME:</br>"
WriteObj.WriteLine "  DEPARTMENT:</td>"
WriteObj.WriteLine "  <td>Signature:</td></tr>"
WriteObj.WriteLine " <tr><td colspan=2>Received Date:</td></tr>"
WriteObj.WriteLine "</table>"
WriteObj.WriteLine "</body>"
WriteObj.WriteLine "</html>"

rptViewer.Navigate App.Path & "\report.html"
End Sub

Private Sub cmdPreview_Click()
'On Error GoTo Xit
rptViewer.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER
Xit:
End Sub
Private Sub cmdPrint_Click()
'On Error GoTo Xit
rptViewer.ExecWB 6, OLECMDEXECOPT_DODEFAULT
'rptViewer.ExecWB OLECMDID_PRINT
' Dim eQuery As OLECMDF       'return value type for QueryStatusWB
'        On Error Resume Next
'            eQuery = rptViewer.QueryStatusWB(OLECMDID_PRINT)  'get print command status
'            If Err.Number = 0 Then
'                    If eQuery And OLECMDF_ENABLED Then
'                        rptViewer.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER, "", ""    'Ok to Print?
'                    Else
'                        MsgBox "The Print command is currently disabled."
'                    End If
'            End If
'            If Err.Number <> 0 Then MsgBox "Print command Error: " & Err.Description
For l = 0 To lstInvoiceNos.ListCount - 1
If lstInvoiceNos.Selected(l) = True Then
InvSelection = Split(lstInvoiceNos.List(l), " - ", Len(lstInvoiceNos.List(l)), vbTextCompare)
cn.Execute "update tblinvoices set transmittalnumber = " & intTransmittalNumber & "  where invoiceNumber = '" & InvSelection(0) & "'"
End If
Next l
txtTransmitToName.Text = ""
txtTransmitToDept.Text = ""
Xit:
End Sub
Private Sub cmdPageSetup_Click()
'On Error GoTo Xit
rptViewer.ExecWB 8, 0
Xit:
End Sub


Private Sub Form_Load()
rptViewer.Navigate "about:blank"
LoadInvoices
End Sub

Private Function LoadInvoices()
Dim rsInvoices As New ADODB.Recordset
If rsInvoices.State Then RS.Close
rsInvoices.Open "Select ExprConcat from vInvoiceVendorValue where TransmittalNumber = 0", cn, 3, 2
While Not rsInvoices.EOF
lstInvoiceNos.AddItem rsInvoices(0)
rsInvoices.MoveNext
Wend
If rsInvoices.State Then rsInvoices.Close
End Function


