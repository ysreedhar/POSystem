VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "Purchase Order System"
   ClientHeight    =   5385
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8130
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuAdministration 
      Caption         =   "Administration"
      Begin VB.Menu mnuVendorManagement 
         Caption         =   "Vendor Management"
      End
      Begin VB.Menu mnuUserManagement 
         Caption         =   "User Management"
      End
      Begin VB.Menu mnuDeliveryPointManagement 
         Caption         =   "Delivery Point Management"
      End
   End
   Begin VB.Menu mnuTransactions 
      Caption         =   "Transactions"
      Begin VB.Menu mnuPurchaseOrderListing 
         Caption         =   "Purchase Order Listing"
      End
      Begin VB.Menu mnuNewPurchaseOrder 
         Caption         =   "Enter Purchase Order"
      End
      Begin VB.Menu mnuServiceOrderListing 
         Caption         =   "Service Order Listing"
      End
      Begin VB.Menu mnuNewServiceOrder 
         Caption         =   "Enter Service Order"
      End
      Begin VB.Menu mnuReports 
         Caption         =   "Reports"
         Begin VB.Menu mnuPrintOrders 
            Caption         =   "Print Purchase Order"
         End
         Begin VB.Menu mnuPrintServiceOrder 
            Caption         =   "Print Service Order"
         End
      End
   End
   Begin VB.Menu mnuInvoiceTransactions 
      Caption         =   "Invoice Transactions"
      Begin VB.Menu mnuInvoiceLog 
         Caption         =   "Invoice Log"
      End
      Begin VB.Menu mnuTransmittalForm 
         Caption         =   "Transmittal Form"
      End
   End
   Begin VB.Menu mnuExitOptions 
      Caption         =   "Exit"
      Begin VB.Menu mnuExit 
         Caption         =   "LogOff"
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuDeliveryPointManagement_Click()
frmDeliveryPointMgmt.Show
End Sub
Private Sub mnuExit_Click()
End
End Sub
Private Sub mnuInvoiceLog_Click()
FrmInvoiceLog.Show
End Sub
Private Sub mnuNewPurchaseOrder_Click()
frmPOEntry.Show
End Sub
Private Sub mnuNewServiceOrder_Click()
frmSOEntry.Show
End Sub
Private Sub mnuPrintOrders_Click()
frmPrintPO.Show
End Sub

Private Sub mnuPrintServiceOrder_Click()
frmPrintSO.Show
End Sub

Private Sub mnuPurchaseOrderListing_Click()
frmPOListing.Show
End Sub

Private Sub mnuServiceOrderListing_Click()
frmSOListing.Show
End Sub

Private Sub mnuTransmittalForm_Click()
FrmTransmittalForm.Show
End Sub

Private Sub mnuUserManagement_Click()
frmNewUser.Show
End Sub
Private Sub mnuVendorManagement_Click()
frmVendor.Show
End Sub
