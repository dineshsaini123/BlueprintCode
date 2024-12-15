VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MainMenu 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   ClientHeight    =   3795
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6630
   Icon            =   "mainmenu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":0152
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":074A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mainmenu.frx":0BEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Invoice"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cash"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "voucher"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu Menumasterdata 
      Caption         =   "&Master Data"
      Begin VB.Menu menuGenralLedgerMaster 
         Caption         =   "&General Ledger Master"
      End
      Begin VB.Menu menusubleadgermaster 
         Caption         =   "&Sub Ledger Master"
      End
      Begin VB.Menu menuInvoiceandpartmaster 
         Caption         =   "&Invoice End Parts Master"
      End
      Begin VB.Menu menucreditnoteandpartmaster 
         Caption         =   "&Credit Note End Part Master"
      End
      Begin VB.Menu menudiscountcategorymaster 
         Caption         =   "&Discount Category Master"
      End
      Begin VB.Menu menuVCashendpartmaster 
         Caption         =   "Counter Sale End &Part Master"
      End
      Begin VB.Menu sepsetup 
         Caption         =   "-"
      End
      Begin VB.Menu menuBooksmaster 
         Caption         =   "B&ooks Master"
      End
      Begin VB.Menu menuAgentMaster 
         Caption         =   "&Agent Master"
      End
      Begin VB.Menu menuDistrictsmaster 
         Caption         =   "Dis&tricts Master"
      End
      Begin VB.Menu menuBookgroupsmaster 
         Caption         =   "&Book Groups Master"
      End
   End
   Begin VB.Menu menutranscationdata 
      Caption         =   "&Transaction Data"
      Begin VB.Menu menujournalvoucher 
         Caption         =   "&Voucher Entry"
      End
      Begin VB.Menu menusalesinvoice 
         Caption         =   "&Sales Invoice"
      End
      Begin VB.Menu menucreditnoteitems 
         Caption         =   "Credit Note (&Items) "
      End
      Begin VB.Menu menucountersale 
         Caption         =   "Coun&Ter Sale"
      End
      Begin VB.Menu menucreditnote 
         Caption         =   "&Credit Note"
      End
      Begin VB.Menu menudebitnote 
         Caption         =   "&Debit Note"
      End
   End
   Begin VB.Menu menureports 
      Caption         =   "&Reports"
      Begin VB.Menu menucashbook 
         Caption         =   "&Cash/Bank Book"
      End
      Begin VB.Menu menugeneralledgeraccounts 
         Caption         =   "&General Ledger Accounts"
      End
      Begin VB.Menu menusubledgeraccounts 
         Caption         =   "&Sub Ledger Accounts"
      End
      Begin VB.Menu alphaSubLedgerAccountsmnu 
         Caption         =   "&Alphabet Wise S. L. A/C"
      End
      Begin VB.Menu menusubledgertrialbalance 
         Caption         =   "Sub. Ledger &Trial Balance"
      End
      Begin VB.Menu menugeneralledgertrialbalance 
         Caption         =   "General Ledger Trial &Balance"
      End
      Begin VB.Menu menudistrictwisesales 
         Caption         =   "&District Wise Sales"
      End
      Begin VB.Menu mnudistrictwisesalesreturn 
         Caption         =   "District &Wise Sales Return "
      End
      Begin VB.Menu menubankadvice 
         Caption         =   "Ba&nk Advice"
      End
   End
   Begin VB.Menu menutools 
      Caption         =   "T&ools"
      Begin VB.Menu menubackupdata 
         Caption         =   "&Backup Data"
      End
      Begin VB.Menu menuRestoredata 
         Caption         =   "&Restore data"
      End
      Begin VB.Menu menudebitnotechecking 
         Caption         =   "&Debit Note Checking"
      End
      Begin VB.Menu menucreditnotechecking 
         Caption         =   "&Credit Note Checking"
      End
      Begin VB.Menu menufinancialyearchecking 
         Caption         =   "&Financial Year Checking"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu menusetup 
         Caption         =   "&Setup"
      End
   End
   Begin VB.Menu EXIT 
      Caption         =   "E&XIT"
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim I As Integer

Private Sub Command1_Click()

End Sub

Private Sub cOMMANDINVOICE_Click()
    INVOICE.Show '    sales
End Sub

Private Sub alphaSubLedgerAccountsmnu_Click()
 Toolbar1.Visible = False
 SLEDGERPRINT.Show
 SLEDGERPRINT.Label4.Visible = False
 SLEDGERPRINT.Combosubledger.Visible = False
 SLEDGERPRINT.Label5.Visible = True
 SLEDGERPRINT.Alpha.Visible = True
End Sub

Private Sub EXIT_Click()
    End
End Sub

Private Sub MDIForm_Load()
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    With con
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
        .Open
    End With
    rs.Open "SETUP", con, adOpenKeyset, adLockReadOnly, adCmdTable
    rs.Find "uname='" + main.username + "'"
 
End Sub

Private Sub menuAgentMaster_Click()
Toolbar1.Visible = False
    bookmaster.SSTab1.Tab = 3
    bookmaster.Show
    bookmaster.Commandmasteradd.SetFocus
End Sub

Private Sub menubankadvice_Click()
Toolbar1.Visible = False
bankadvice.Show
End Sub

Private Sub menuBookgroupsmaster_Click()
Toolbar1.Visible = False
    bookmaster.SSTab1.Tab = 1
    bookmaster.Show
End Sub

Private Sub menuBooksmaster_Click()
Toolbar1.Visible = False
    bookmaster.SSTab1.Tab = 0
    bookmaster.Show
   
End Sub

Private Sub menucashbankAccountsmaster_Click()
Toolbar1.Visible = False
    master.SSTab1.Tab = 5
    master.Show
End Sub

Private Sub menucashbook_Click()
Toolbar1.Visible = False
CASHBOOK.Show
End Sub

Private Sub menucountersale_Click()
Toolbar1.Visible = False
countersale.Show
End Sub

Private Sub menucreditnote_Click()
Toolbar1.Visible = False
Creditnotefile.Show
End Sub

Private Sub menucreditnoteandpartmaster_Click()
Toolbar1.Visible = False
    master.SSTab1.Tab = 3
    master.Show
End Sub

Private Sub menucreditnoteitems_Click()
Toolbar1.Visible = False
    CRITNOTE.Show '    sales
End Sub

Private Sub menudebitnote_Click()
Toolbar1.Visible = False
Debitnotefile.Show
End Sub

Private Sub menudiscountcategorymaster_Click()
Toolbar1.Visible = False
    master.SSTab1.Tab = 4
    master.Show
End Sub

Private Sub menuDistrictsmaster_Click()
Toolbar1.Visible = False
    bookmaster.SSTab1.Tab = 2
    bookmaster.Show
    bookmaster.Commandmasteradd.SetFocus
End Sub

Private Sub menudistrictwisesales_Click()
Toolbar1.Visible = False
DWsales.Show
End Sub

Private Sub menugeneralledgeraccounts_Click()
Toolbar1.Visible = False
    GLEDGERPRINT.Show
End Sub

Private Sub menuGenralLedgerMaster_Click()
Toolbar1.Visible = False
    master.SSTab1.Tab = 0
    master.Show
End Sub

Private Sub menuInvoiceandpartmaster_Click()
Toolbar1.Visible = False
    master.SSTab1.Tab = 2
    master.Show
End Sub
Private Sub menujournalvoucher_Click()
Toolbar1.Visible = False
    Voucherform.vtype.Text = "J"
    Voucherform.Show
End Sub
Private Sub menupaymentvoucher_Click()
Toolbar1.Visible = False
    Voucherform.vtype.Text = "P"
    Voucherform.Show
     
End Sub

Private Sub menuReceiptvoucher_Click()
Toolbar1.Visible = False
    Voucherform.vtype.Text = "R"
    Voucherform.Show
End Sub

Private Sub menusalesinvoice_Click()
Toolbar1.Visible = False
    INVOICE.Show '    sales
End Sub

Private Sub menusetup_Click()
Toolbar1.Visible = False
    setup.Show
End Sub

Private Sub menusubleadgermaster_Click()
Toolbar1.Visible = False
    master.SSTab1.Tab = 1
    master.Show
End Sub

Private Sub menusubledgeraccounts_Click()
Toolbar1.Visible = False
 SLEDGERPRINT.Show
 SLEDGERPRINT.Label4.Visible = True
 SLEDGERPRINT.Combosubledger.Visible = True
 SLEDGERPRINT.Label5.Visible = False
 SLEDGERPRINT.Alpha.Visible = False
 
End Sub

Private Sub menusubledgertrialbalance_Click()
Toolbar1.Visible = False
    subtrial.Show
End Sub

Private Sub setup_Click()

End Sub

Private Sub menuVCashendpartmaster_Click()
Toolbar1.Visible = False
    master.SSTab1.Tab = 5
    master.Show
End Sub

Private Sub mnudistrictwisesalesreturn_Click()
   Toolbar1.Visible = False
   DWSalesReturn.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Invoice"
    INVOICE.Show '    sales
Case "voucher"
    Voucherform.vtype.Text = "J"
    Voucherform.Show
Case "Cash"
    countersale.Show
 

End Select
Toolbar1.Visible = False
End Sub
