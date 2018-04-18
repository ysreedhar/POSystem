VERSION 5.00
Begin VB.Form frmPOreport 
   Caption         =   "Invoice Summary Report"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11055
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
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
      Left            =   4320
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
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
      Left            =   2880
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox cmbPONumbers 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PO Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "frmPOreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub cmdGenerate_Click()
    Dim rs2 As New ADODB.Recordset
     Dim RS As New ADODB.Recordset
     Dim Sum As Currency
    '*****************************************************************************
    Dim a As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set a = fso.CreateTextFile(App.Path & "\report.html", True)
    Sum = 0
    a.WriteLine ("<html>")
    a.WriteLine ("<head>")
    a.WriteLine ("<Style type=text/css>P {page-break-before:always}</Style>")
    a.WriteLine ("</head>")
    a.WriteLine "<BODY>"
    sql = "select * from PO_Header where PONumber = " & cmbPONumbers.Text
    rs2.Open sql, cn, 1, 1
    Sno = 0
    While Not rs2.EOF
    Sno = Sno + 1
    vLineCount = vLineCount + 1
    If vLineCount > 28 Then
    a.WriteLine "<P></p>"
    End If
    a.WriteLine "<table border=0 cellspacing=0 cellpadding=0>"
    a.WriteLine "<tr>"
    a.WriteLine "  <td width=295 colspan=3 valign=top>"
    a.WriteLine "  <p>To</p>"
    a.WriteLine "  <p></p>"
    a.WriteLine "  </td>"
    a.WriteLine "  <td width=197 colspan=2 valign=top>"
    a.WriteLine "  <p> </p>"
    a.WriteLine "  <p> </p>"
    a.WriteLine "  </td>"
    a.WriteLine "  <td width=98 valign=top >"
    a.WriteLine "  <p >" & Format(rs2!PODate, "dd-MMM-yy") & "</p>"
    a.WriteLine "  </td>"
    a.WriteLine " </tr>"
    a.WriteLine " <tr>"
    a.WriteLine "  <td width=295 colspan=3 valign=top  >"
    a.WriteLine "  <p ></p>"
    a.WriteLine "</td>"
    a.WriteLine "<td width=197 colspan=2 valign=top>"
    a.WriteLine "<p >" & rs2!Requisition_OrderNo & "</p>"
    a.WriteLine "</td>"
    a.WriteLine "<td width=98 valign=top >"
    a.WriteLine "<p >" & rs2!CostCenter & "</p>"
    a.WriteLine "</td>"
    a.WriteLine "</tr>"
    a.WriteLine "<tr>"
    a.WriteLine "<td width=295 colspan=3 valign=top>"
    a.WriteLine "<p ></p>"
    a.WriteLine "</td>"
    a.WriteLine "<td width=197 colspan=2 valign=top >"
    a.WriteLine "<p >" & rs2!Orderedby & "</p>"
    a.WriteLine "</td>"
    a.WriteLine "<td width=98 valign=top >"
    a.WriteLine "<p >" & rs2!Approvedby & "</p>"
    a.WriteLine "</td>"
    a.WriteLine "</tr>"
    a.WriteLine "<tr>"
    a.WriteLine "<td width=295 colspan=3 valign=top>"
    a.WriteLine "<p ></p>"
    a.WriteLine "</td>"
    a.WriteLine "<td width=295 colspan=3 valign=top >"
    a.WriteLine "<p >" & rs2!Remarks & "</p>"
    a.WriteLine "</td>"
    a.WriteLine "</tr>"
    a.WriteLine "<tr>"
    RS.Open "Select * from PO_details where POID = rs2!POID, cn, 1, 1"
    a.WriteLine "<td width=590 colspan=6 valign=top  >"
    a.WriteLine "<p ></p>"
    a.WriteLine "</td>"
    a.WriteLine "</tr>"
    a.WriteLine "<tr>"
    a.WriteLine "<td width=98 valign=top   >"
    a.WriteLine "<p >" & Sno & "</p>"
    a.WriteLine "</td>"
    a.WriteLine "<td width=98 valign=top  >"
    a.WriteLine "<p >" & RS!Quantity & "</p>"
    a.WriteLine "</td>"
    a.WriteLine "<td width=98 valign=top >"
    a.WriteLine "<p >" & RS!ItemDescription & "</p>"
    a.WriteLine "</td>"
    a.WriteLine "<td width=98 valign=top  >"
    a.WriteLine "<p >" & RS!AccountCode & "</p>"
    a.WriteLine "</td>"
    a.WriteLine "<td width=98 valign=top >"
    a.WriteLine "<p >" & RS!UnitPrice & "</p>"
    a.WriteLine "</td>"
    a.WriteLine "<td width=98 valign=top  >"
    a.WriteLine "<p >" & RS!Quantity * RS!UnitPrice & "</p>"
    GrandTotal = GrandTotal + (RS!Quantity * RS!UnitPrice)
    a.WriteLine "</td>"
    a.WriteLine "</tr>"
    a.WriteLine "<tr>"
    a.WriteLine "<td width=98 valign=top>"
    a.WriteLine "<p ></p>"
    a.WriteLine "</td>"
    a.WriteLine "<td width=98 valign=top >"
    a.WriteLine "</td>"
    a.WriteLine "<td width=98 valign=top >"
    a.WriteLine "<p ></p>"
    a.WriteLine " </td>"
    a.WriteLine "<td width=98 valign=top >"
    a.WriteLine "  <p ></p>"
    a.WriteLine "  </td>"
    a.WriteLine "<td width=98 valign=top>"
    a.WriteLine "<p ></p>"
    a.WriteLine "  </td>"
    a.WriteLine "<td width=98 valign=top  >"
    a.WriteLine " </td>"
    a.WriteLine " </tr>"
    RS.Close
    rs2.MoveNext
    Wend
    
    While vLineCount <> 28
    a.WriteLine "<tr><td COLSPAN=5><br></td></tr>"
    vLineCount = vLineCount + 1
    Wend
    a.WriteLine "</table>"
    
    a.WriteLine "</body>"
    a.WriteLine "</html>"
    a.Close
    WebBrowser1.Navigate App.Path & "\report.html"

End Sub

Private Sub Form_Load()
WebBrowser1.Navigate "about:blank"
LoadPONumbers
End Sub
Private Sub LoadPONumbers()
cmbPONumbers.Clear
Dim rsPoNumbers As New ADODB.Recordset
If rsPoNumbers.State Then RS.Close
rsPoNumbers.Open "select PO_OrderNo from PO_header  ", cn, 3, 2
While Not rsPoNumbers
cmbPONumbers.AddItem rsPoNumbers(0)
Wend
If rsPoNumbers.State Then rsPoNumbers.Close
cn.Close
End Sub
