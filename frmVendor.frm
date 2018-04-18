VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmVendor 
   Caption         =   "Vendor Management"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   795
   ClientWidth     =   11235
   ClipControls    =   0   'False
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
   ScaleHeight     =   566
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   749
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5790
      Left            =   652
      TabIndex        =   14
      Top             =   660
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   10213
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   529
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Create New Vendor"
      TabPicture(0)   =   "frmVendor.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "List of Vendors/ Modify Vendor"
      TabPicture(1)   =   "frmVendor.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame1 
         Caption         =   "  New Vendor Details  "
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
         Height          =   5400
         Left            =   -74985
         TabIndex        =   32
         Top             =   360
         Width           =   9825
         Begin VB.CheckBox chkAVL 
            Caption         =   "AVL"
            Height          =   495
            Left            =   5205
            TabIndex        =   63
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtVendorCode 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   6840
            MaxLength       =   30
            TabIndex        =   61
            Top             =   720
            Width           =   1575
         End
         Begin VB.Frame Frame2 
            Caption         =   "  Options  "
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
            Height          =   975
            Left            =   7215
            TabIndex        =   45
            Top             =   3990
            Width           =   2265
            Begin MSForms.CommandButton cmdclosevendor 
               Height          =   495
               Left            =   1200
               TabIndex        =   47
               ToolTipText     =   "Click to Close Contents (Alt+S)"
               Top             =   240
               Width           =   975
               Caption         =   "Close"
               PicturePosition =   131072
               Size            =   "1720;873"
               MousePointer    =   99
               Accelerator     =   99
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
            Begin MSForms.CommandButton cmdSaveVendor 
               Height          =   495
               Left            =   120
               TabIndex        =   46
               ToolTipText     =   "Click to Save Contents (Alt+S)"
               Top             =   240
               Width           =   975
               Caption         =   "Save"
               PicturePosition =   131072
               Size            =   "1720;873"
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
         End
         Begin VB.TextBox TxtRegNo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2055
            MaxLength       =   30
            TabIndex        =   44
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txtVendorCountry 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2055
            MaxLength       =   25
            TabIndex        =   43
            Top             =   2925
            Width           =   2895
         End
         Begin VB.TextBox txtVendorCity 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2055
            MaxLength       =   25
            TabIndex        =   42
            Top             =   2010
            Width           =   2895
         End
         Begin VB.TextBox txtVendorName 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2055
            MaxLength       =   75
            TabIndex        =   41
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox txtVendorEmail 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2055
            MaxLength       =   35
            TabIndex        =   40
            Top             =   4755
            Width           =   2880
         End
         Begin VB.TextBox txtVendorFax 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2055
            MaxLength       =   20
            TabIndex        =   39
            Top             =   4290
            Width           =   2895
         End
         Begin VB.TextBox txtVendorZip 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2055
            MaxLength       =   15
            TabIndex        =   38
            Top             =   2490
            Width           =   2895
         End
         Begin VB.TextBox txtVendorMob 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2040
            MaxLength       =   20
            TabIndex        =   37
            Top             =   3810
            Width           =   2895
         End
         Begin VB.TextBox txtVendorPh 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2055
            MaxLength       =   20
            TabIndex        =   36
            Top             =   3405
            Width           =   2895
         End
         Begin VB.TextBox txtVendorAdd 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2055
            MaxLength       =   255
            TabIndex        =   35
            Top             =   1545
            Width           =   2880
         End
         Begin VB.TextBox txtVendorFirmName 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2055
            MaxLength       =   75
            TabIndex        =   34
            Top             =   210
            Width           =   2880
         End
         Begin VB.TextBox txtVendorId 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   9165
            TabIndex        =   33
            Top             =   300
            Visible         =   0   'False
            Width           =   375
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   6840
            TabIndex        =   48
            Top             =   240
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   582
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
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   16515075
            CurrentDate     =   37897
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Code"
            Height          =   195
            Left            =   5205
            TabIndex        =   62
            Top             =   840
            Width           =   930
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier Since"
            Height          =   255
            Left            =   5205
            TabIndex        =   60
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Biz. Registration No."
            Height          =   195
            Left            =   135
            TabIndex        =   59
            Top             =   600
            Width           =   1470
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
            Height          =   255
            Left            =   135
            TabIndex        =   58
            Top             =   3030
            Width           =   1335
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "City"
            Height          =   255
            Left            =   135
            TabIndex        =   57
            Top             =   2070
            Width           =   1215
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Company / Firm Name"
            Height          =   375
            Left            =   135
            TabIndex        =   56
            Top             =   240
            Width           =   2310
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            Height          =   255
            Left            =   135
            TabIndex        =   55
            Top             =   4770
            Width           =   1815
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
            Height          =   255
            Left            =   135
            TabIndex        =   54
            Top             =   4290
            Width           =   1935
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Pos Code"
            Height          =   255
            Left            =   135
            TabIndex        =   53
            Top             =   2490
            Width           =   1815
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile"
            Height          =   375
            Left            =   135
            TabIndex        =   52
            Top             =   3810
            Width           =   1695
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            Height          =   375
            Left            =   135
            TabIndex        =   51
            Top             =   3405
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   375
            Left            =   135
            TabIndex        =   50
            Top             =   1590
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Person"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   135
            TabIndex        =   49
            Top             =   1110
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "  Options  "
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
         Height          =   3210
         Left            =   8325
         TabIndex        =   29
         Top             =   2520
         Width           =   1560
         Begin MSForms.CommandButton cmdclose1 
            Height          =   495
            Left            =   345
            TabIndex        =   13
            ToolTipText     =   "Click to Close (Alt+S)"
            Top             =   1575
            Width           =   975
            Caption         =   "Close"
            PicturePosition =   131072
            Size            =   "1720;873"
            MousePointer    =   99
            Accelerator     =   99
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdExVenDelete 
            Height          =   495
            Left            =   360
            TabIndex        =   12
            ToolTipText     =   "Click to Delete Contents (Alt+D)"
            Top             =   975
            Width           =   975
            Caption         =   "Delete"
            PicturePosition =   131072
            Size            =   "1720;873"
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
         Begin MSForms.CommandButton cmdExVenSave 
            Height          =   495
            Left            =   360
            TabIndex        =   11
            ToolTipText     =   "Click toSave (Alt+S)"
            Top             =   360
            Width           =   975
            Caption         =   "Save"
            PicturePosition =   131072
            Size            =   "1720;873"
            MousePointer    =   99
            Accelerator     =   99
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "  Existing Vendor Details  "
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
         Height          =   2235
         Left            =   60
         TabIndex        =   28
         Top             =   315
         Width           =   9810
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   1845
            Left            =   135
            TabIndex        =   16
            Top             =   360
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   3254
            _Version        =   393216
            BackColor       =   16777215
            BackColorFixed  =   14737632
            BackColorSel    =   16777215
            ForeColorSel    =   14723990
            BackColorBkg    =   16777215
            AllowUserResizing=   1
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
      End
      Begin VB.Frame Frame4 
         Caption         =   "  Details  "
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
         Height          =   3180
         Left            =   45
         TabIndex        =   17
         Top             =   2565
         Width           =   8250
         Begin VB.CheckBox chkAVLM 
            Caption         =   "AVL"
            Height          =   495
            Left            =   6360
            TabIndex        =   66
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox txtVendorCodeM 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   5880
            MaxLength       =   30
            TabIndex        =   64
            Top             =   120
            Width           =   1575
         End
         Begin VB.TextBox TxtRegNoM 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   5880
            MaxLength       =   35
            TabIndex        =   2
            Top             =   623
            Width           =   2175
         End
         Begin VB.TextBox txtExVCountry 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   6390
            MaxLength       =   25
            TabIndex        =   5
            Top             =   1605
            Width           =   1695
         End
         Begin VB.TextBox txtExVCity 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            MaxLength       =   25
            TabIndex        =   4
            Top             =   1605
            Width           =   1440
         End
         Begin VB.TextBox txtExVName 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            MaxLength       =   75
            TabIndex        =   1
            Top             =   645
            Width           =   2340
         End
         Begin VB.TextBox txtExVZip 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3765
            MaxLength       =   15
            TabIndex        =   6
            Top             =   1605
            Width           =   1725
         End
         Begin VB.TextBox txtExVFax 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   6360
            MaxLength       =   20
            TabIndex        =   9
            Top             =   2100
            Width           =   1725
         End
         Begin VB.TextBox txtExVEmai 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            TabIndex        =   10
            Top             =   2565
            Width           =   4005
         End
         Begin VB.TextBox txtExVMob 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3825
            MaxLength       =   20
            TabIndex        =   8
            Top             =   2085
            Width           =   1620
         End
         Begin VB.TextBox txtExVPh 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   7
            Top             =   2085
            Width           =   1455
         End
         Begin VB.TextBox txtExVAddress 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            MaxLength       =   255
            TabIndex        =   3
            Top             =   1125
            Width           =   3885
         End
         Begin VB.TextBox txtExVFirm 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1560
            MaxLength       =   75
            TabIndex        =   0
            Top             =   165
            Width           =   3180
         End
         Begin VB.TextBox txtExVID 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5760
            TabIndex        =   15
            Top             =   2565
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor Code"
            Height          =   195
            Left            =   4920
            TabIndex        =   65
            Top             =   233
            Width           =   930
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Biz. Registration No."
            Height          =   195
            Left            =   4380
            TabIndex        =   31
            Top             =   713
            Width           =   1470
         End
         Begin VB.Label Label21 
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
            Height          =   255
            Left            =   5670
            TabIndex        =   27
            Top             =   1635
            Width           =   735
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "City"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1725
            Width           =   1215
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Company / Firm"
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Person"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   765
            Width           =   1455
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   2685
            Width           =   975
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pos Code"
            Height          =   195
            Left            =   3075
            TabIndex        =   22
            Top             =   1665
            Width           =   675
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   2205
            Width           =   1095
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1245
            Width           =   855
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile"
            Height          =   255
            Left            =   3180
            TabIndex        =   19
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
            Height          =   255
            Left            =   5730
            TabIndex        =   18
            Top             =   2160
            Width           =   495
         End
      End
   End
   Begin VB.Label Vendor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Management"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4155
      TabIndex        =   30
      Top             =   180
      Width           =   2925
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6517
      Stretch         =   -1  'True
      Top             =   6660
      Width           =   2325
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   9037
      Top             =   6810
      Width           =   885
   End
End
Attribute VB_Name = "frmVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rowno As Integer
Dim verify As Boolean
Private Sub cmdclose1_Click()
    Unload Me
End Sub
Private Sub cmdclosevendor_Click()
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Call UPPER_CASE(KeyAscii)
If KeyAscii = "27" Then Call Image2_Click
End Sub
Private Sub Image2_Click()
    Unload Me
End Sub
Private Function validate_fields() As Boolean
    verify = False
    If Len(Trim(txtVendorFirmName.Text)) = 0 Then
        MsgBox "Please enter the Vendor's firm name", vbExclamation
        txtVendorFirmName.SetFocus
        validate_fields = False
    ElseIf Len(Trim(txtVendorName.Text)) = 0 Then
        MsgBox "Please enter the Vendor's name", vbExclamation
        txtVendorName.SetFocus
        validate_fields = False
    ElseIf Len(Trim(txtVendorAdd.Text)) = 0 Then
        MsgBox "Please enter the Vendor's address", vbExclamation
        txtVendorAdd.SetFocus
        validate_fields = False
    ElseIf Len(Trim(txtVendorCity.Text)) = 0 Then
        MsgBox "Please enter the Vendor's city", vbExclamation
        txtVendorCity.SetFocus
        validate_fields = False
    ElseIf Len(Trim(txtVendorCountry.Text)) = 0 Then
        MsgBox "Please enter the Vendor's country", vbExclamation
        txtVendorCountry.SetFocus
        validate_fields = False
    ElseIf Len(Trim(TxtRegNo.Text)) = 0 Then
         MsgBox "Please enter the registration No.", vbExclamation
         TxtRegNo.SetFocus
         validate_fields = False
    Else
        validate_fields = True
    End If
End Function
Private Sub cmdExVenDelete_Click()
'On Error Resume Next
Dim Msg, Style, response
If txtExVID = "" Then
        MsgBox "First Select Any Record", vbOKOnly + vbCritical
Else
    Msg = "Do you really want to delete this record ?"   ' Define message.
    Style = vbYesNo + vbCritical ' Define buttons.
    response = MsgBox(Msg, Style)
    If response = vbYes Then   ' User chose Yes.
cn.Execute " Update Vendor set  v_status='No'  where vendor_id=" & txtExVID.Text
            Call Vendor_Detail
            Call ClearText
            Call Disable_Existing_Vendor
    End If
End If
End Sub
Private Sub cmdExVenSave_Click()
'On Error Resume Next
Dim Msg, Style, response
    verify = ex_validate_fields
    If verify = True Then
        Msg = "Do you really want to update this record ?"   ' Define message.
        Style = vbYesNo + vbCritical ' Define buttons.
        response = MsgBox(Msg, Style)
        If response = vbYes Then   ' User chose Yes.
        Dim RS As New ADODB.Recordset
            RS.Open ("select * from  vendor where vendor_id =" & txtExVID.Text), cn, 3, 2
If RS.EOF Then RS.AddNew
 RS("v_name") = txtExVFirm.Text
 RS("v_contactperson") = txtExVName.Text
 RS("v_address") = txtExVAddress.Text
 RS("v_city") = txtExVCity.Text
 RS("v_country") = txtExVCountry.Text
 RS("v_phone") = txtExVPh.Text
 RS("v_mobile") = txtExVMob.Text
 RS("v_zip") = txtExVZip.Text
 RS("v_status") = "Yes"
 RS("v_fax") = txtExVFax.Text
 RS("v_email") = txtExVEmai.Text
 RS("v_regno") = TxtRegNoM.Text
 RS("VendorCode") = txtVendorCodeM.Text
 RS("AVL") = chkAVLM.Value
 RS.Update
 RS.Close
                Call Vendor_Detail
                Call ClearText
                Call Disable_Existing_Vendor
        End If
    End If
End Sub
Private Function ex_validate_fields() As Boolean
    verify = False
If Not IsNumeric(txtExVID.Text) Then
    MsgBox "Please Choose the Vendor's ID", vbExclamation
        ex_validate_fields = False
ElseIf Trim(txtExVFirm.Text) = "" Then
        MsgBox "Please enter the Vendor's firm name", vbExclamation
        txtExVFirm.SetFocus
        ex_validate_fields = False
    ElseIf Trim(txtExVName.Text) = "" Then
        MsgBox "Please enter the Vendor's name", vbExclamation
        txtExVName.SetFocus
        ex_validate_fields = False
    ElseIf Trim(txtExVAddress.Text) = "" Then
        MsgBox "Please enter the Vendor's address", vbExclamation
        txtExVAddress.SetFocus
        ex_validate_fields = False
    ElseIf Trim(txtExVCity.Text) = "" Then
        MsgBox "Please enter the Vendor's city", vbExclamation
        txtExVCity.SetFocus
        ex_validate_fields = False
    ElseIf Trim(txtExVCountry.Text) = "" Then
        MsgBox "Please enter the Vendor's country", vbExclamation
        txtExVCountry.SetFocus
        ex_validate_fields = False
    ElseIf Len(Trim(TxtRegNoM.Text)) = 0 Then
        MsgBox "Please enter the Registration No", vbExclamation
        TxtRegNoM.SetFocus
        ex_validate_fields = False
    Else
        ex_validate_fields = True
    End If
End Function
Private Sub ClearTextNewVendor()
        txtVendorAdd.Text = ""
        txtVendorCity.Text = ""
        txtVendorCountry.Text = ""
        txtVendorEmail.Text = ""
        txtVendorFax.Text = ""
        txtVendorFirmName.Text = ""
        txtVendorId.Text = ""
        txtVendorMob.Text = ""
        txtVendorName.Text = ""
        txtVendorPh.Text = ""
        txtVendorZip.Text = ""
        TxtRegNo.Text = ""
        txtVendorFirmName.SetFocus
End Sub
Private Sub ClearText()
        txtExVAddress.Text = ""
        txtExVEmai.Text = ""
        txtExVFax.Text = ""
        txtExVFirm.Text = ""
        txtExVMob.Text = ""
        txtExVID.Text = ""
        txtExVName.Text = ""
        txtExVPh.Text = ""
        txtExVZip.Text = ""
        txtExVCity.Text = ""
        txtExVCountry.Text = ""
        TxtRegNoM.Text = ""
        
End Sub

Private Sub cmdSaveVendor_Click()
'On Error Resume Next
    Dim Msg, Style, response
    verify = validate_fields
    If verify = True Then
        Msg = "Do you really want to save this record ....?"   ' Define message.
        Style = vbYesNo + vbCritical ' Define buttons.
        response = MsgBox(Msg, Style)
        If response = vbYes Then   ' User chose Yes.
Dim RS As New ADODB.Recordset
RS.Open ("select * from  vendor"), cn, 3, 2
            RS.AddNew
 RS("v_name") = txtVendorFirmName.Text
 RS("v_contactperson") = txtVendorName.Text
 RS("v_address") = txtVendorAdd.Text
 RS("v_city") = txtVendorCity.Text
 RS("v_country") = txtVendorCountry.Text
 RS("v_phone") = txtVendorPh.Text
 RS("v_mobile") = txtVendorMob.Text
 RS("v_zip") = txtVendorZip.Text
 RS("v_status") = "Yes"
 RS("v_fax") = txtVendorFax.Text
 RS("v_email") = txtVendorEmail.Text
 RS("v_regno") = Trim(TxtRegNo.Text)
 RS("REGDATE") = DTPicker1.Value
 RS("VendorCode") = txtVendorCode.Text
 RS("AVL") = chkAVL.Value
 RS.Update
 RS.Close
            Call ClearTextNewVendor
        Else
            Call VendorID
            txtVendorFirmName.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()


    Call connect
    Call VendorID
    Call Vendor_Detail
    Call Disable_Existing_Vendor
    

    
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
 If SSTab1.Tab = 0 Then
        Call ClearTextNewVendor
        Call VendorID
        txtVendorFirmName.SetFocus
ElseIf SSTab1.Tab = 1 Then
       Call Vendor_Detail
       Call Disable_Existing_Vendor
        Call ClearText
        txtExVFirm.SetFocus
    End If
End Sub
Private Sub Enable_Existing_Vendor()
   
End Sub

Private Sub Disable_Existing_Vendor()
    
End Sub
Private Sub MSFlexGrid1_SelChange()
    'On Error Resume Next
    TxtRegNoM.Text = ""
        MSFlexGrid1.Col = 1
        rowno = MSFlexGrid1.Text
Dim RS As New ADODB.Recordset
        RS.Open ("Select * from vendor where vendor_id = " & rowno), cn, 3, 2
        txtExVID.Text = RS.Fields("vendor_id")
        txtExVAddress.Text = RS.Fields("v_address")
        txtExVFax.Text = RS.Fields("v_fax")
        txtExVFirm.Text = RS.Fields("v_name")
        txtExVMob.Text = RS.Fields("v_mobile")
        txtExVName.Text = RS.Fields("v_contactperson")
        txtExVPh.Text = RS.Fields("v_phone")
        txtExVZip.Text = RS.Fields("v_zip")
        txtExVEmai.Text = "" & RS.Fields("v_email")
        txtExVCity.Text = RS.Fields("v_city")
        txtExVCountry.Text = RS.Fields("v_country")
        TxtRegNoM.Text = RS.Fields("v_regno")
RS.Close
    Call Disable_Existing_Vendor
End Sub
Private Sub VendorID()
'On Error Resume Next
Dim RS As New ADODB.Recordset

   RS.Open ("Select  Max(Vendor_id) from vendor"), cn, 3, 2
    If RS.Fields(0) > 0 Then
        txtVendorId.Text = Val(RS.Fields(0)) + 1
    Else
        txtVendorId.Text = 1
    End If
End Sub
Public Sub Vendor_Detail()
'On Error Resume Next
Dim RS As New ADODB.Recordset
    RS.Open ("Select Vendor_id,v_name,v_contactperson,v_address,v_city,v_country,v_phone,v_mobile,v_zip,v_fax,v_email,v_regno from Vendor where v_status = 'Yes' order by v_name"), cn, 3, 2
Dim i As Integer
    While Not RS.EOF
        i = i + 1
        RS.MoveNext
    Wend
    MSFlexGrid1.Cols = 13
    MSFlexGrid1.Rows = i + 1
    MSFlexGrid1.Row = 0
 
    MSFlexGrid1.Col = 0
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellAlignment = 4
    MSFlexGrid1.Text = "S.No."
    MSFlexGrid1.ColWidth(0) = 550
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Text = "ID"
    MSFlexGrid1.ColWidth(1) = 0
    
    MSFlexGrid1.Col = 2
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellAlignment = 4
    MSFlexGrid1.Text = "Company / Firm Name"
    MSFlexGrid1.ColWidth(2) = 2950
    
    MSFlexGrid1.Col = 3
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellAlignment = 4
    MSFlexGrid1.Text = "Contact Person"
    MSFlexGrid1.ColWidth(3) = 2000
    
    MSFlexGrid1.Col = 4
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellAlignment = 4
    MSFlexGrid1.Text = " Address "
    MSFlexGrid1.ColWidth(4) = 2500
    
    MSFlexGrid1.Col = 5
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellAlignment = 4
    MSFlexGrid1.Text = " City "
    MSFlexGrid1.ColWidth(5) = 1500
    
    MSFlexGrid1.Col = 6
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellAlignment = 4
    MSFlexGrid1.Text = " Country "
    MSFlexGrid1.ColWidth(6) = 1000
    
    MSFlexGrid1.Col = 7
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellAlignment = 4
    MSFlexGrid1.Text = "Phone"
    MSFlexGrid1.ColWidth(7) = 1000
    
    MSFlexGrid1.Col = 8
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellAlignment = 4
    MSFlexGrid1.Text = "Mobile"
    MSFlexGrid1.ColWidth(8) = 1000
    
    MSFlexGrid1.Col = 9
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellAlignment = 4
    MSFlexGrid1.Text = "Zip"
    MSFlexGrid1.ColWidth(9) = 1000
    
    MSFlexGrid1.Col = 10
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellAlignment = 4
    MSFlexGrid1.Text = "Fax"
    MSFlexGrid1.ColWidth(10) = 1000
    
    MSFlexGrid1.Col = 11
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellAlignment = 4
    MSFlexGrid1.Text = "Email"
    MSFlexGrid1.ColWidth(11) = 3000
    
    MSFlexGrid1.Col = 12
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.CellAlignment = 4
    MSFlexGrid1.Text = "Registration No."
    MSFlexGrid1.ColWidth(12) = 3000
    i = 1
    If RS.State = 1 Then
        RS.Close
    End If
      RS.Open ("Select Vendor_id,v_name,v_contactperson,v_address,v_city,v_country,v_phone,v_mobile,v_zip,v_fax,v_email,v_regno from Vendor where v_status = 'Yes' order by v_name"), cn, 3, 2
       
    While Not RS.EOF
        MSFlexGrid1.Row = i
        k = 0
        MSFlexGrid1.Col = 0
        MSFlexGrid1.CellAlignment = 4
        MSFlexGrid1.CellFontBold = True
        MSFlexGrid1.Text = i
        For j = 1 To 12
            MSFlexGrid1.Col = j
            MSFlexGrid1.Text = "" & RS.Fields(k)
            k = k + 1
        Next j
        RS.MoveNext
        i = i + 1
    Wend
End Sub

Private Sub txtExVAddress_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptString(Me.txtExVAddress, KeyAscii)
End Sub

Private Sub txtExVAddress_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtExVCity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptString(Me.txtExVCity, KeyAscii)
End Sub

Private Sub txtExVCity_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtExVCountry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptString(Me.txtExVCountry, KeyAscii)
End Sub

Private Sub txtExVCountry_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


Private Sub txtExVEmai_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub

KeyAscii = TxtAcceptString(Me.txtExVEmai, KeyAscii)
End Sub

Private Sub txtExVEmai_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtExVFax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptString(Me.txtExVFax, KeyAscii)
End Sub

Private Sub txtExVFax_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtExVFirm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptString(Me.txtExVFirm, KeyAscii)
End Sub

Private Sub txtExVFirm_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtExVMob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptNumeric(Me.txtExVMob, KeyAscii)
End Sub

Private Sub txtExVMob_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtExVName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptString(Me.txtExVName, KeyAscii)
End Sub

Private Sub txtExVName_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtExVPh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptNumeric(Me.txtExVPh, KeyAscii)
End Sub

Private Sub txtExVPh_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtExVZip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptNumeric(Me.txtExVZip, KeyAscii)
End Sub

Private Sub txtExVZip_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtRegNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptNumeric(Me.TxtRegNo, KeyAscii)
End Sub

Private Sub txtregno_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub TxtRegNoM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptNumeric(Me.TxtRegNoM, KeyAscii)
End Sub

Private Sub TxtRegNoM_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtVendorAdd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
 KeyAscii = TxtAcceptString(Me.txtVendorAdd, KeyAscii)
End Sub

Private Sub txtVendorAdd_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtVendorCity_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptString(Me.txtVendorCity, KeyAscii)
End Sub

Private Sub txtVendorCity_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtVendorCountry_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptString(Me.txtVendorCountry, KeyAscii)

End Sub

Private Sub txtVendorCountry_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtVendorEmail_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptString(Me.txtVendorEmail, KeyAscii)
End Sub

Private Sub txtVendorEmail_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtVendorFax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub

KeyAscii = TxtAcceptNumeric(Me.txtVendorFax, KeyAscii)


End Sub

Private Sub txtVendorFax_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtVendorFirmName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptString(Me.txtVendorFirmName, KeyAscii)
End Sub

Private Sub txtVendorFirmName_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtVendorMob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptNumeric(Me.txtVendorMob, KeyAscii)



End Sub

Private Sub txtVendorMob_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtVendorName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptString(Me.txtVendorName, KeyAscii)
End Sub

Private Sub txtVendorName_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtVendorPh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptNumeric(Me.txtVendorPh, KeyAscii)

End Sub

Private Sub txtVendorPh_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtVendorZip_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Exit Sub
KeyAscii = TxtAcceptNumeric(Me.txtVendorZip, KeyAscii)

End Sub

Private Sub txtVendorZip_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then SendKeys "{TAB}"
End Sub
