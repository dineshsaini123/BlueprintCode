VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MainMenu 
   AutoShowChildren=   0   'False
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9744
   ClientLeft      =   3468
   ClientTop       =   1212
   ClientWidth     =   18420
   Icon            =   "mmenu.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      ForeColor       =   &H00C0E0FF&
      Height          =   675
      Left            =   0
      ScaleHeight     =   624
      ScaleWidth      =   18372
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   18420
      Begin VB.CommandButton cmdBasil 
         Caption         =   "&Basil"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6420
         Picture         =   "mmenu.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Basil"
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdPaper 
         Caption         =   "&Production"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4860
         Picture         =   "mmenu.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Production"
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cmdStock 
         Caption         =   "&Stock System"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3180
         Picture         =   "mmenu.frx":09D8
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Stock System"
         Top             =   0
         Width           =   1575
      End
      Begin VB.CommandButton cmdConven 
         Caption         =   "&Canvassing"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1620
         Picture         =   "mmenu.frx":0E1A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Canvassing"
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton cndInv 
         Caption         =   "&Invoicing"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   60
         Picture         =   "mmenu.frx":125C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Invoicing"
         Top             =   0
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   288
      Top             =   1548
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   75
      Top             =   3120
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   22
      ImageHeight     =   27
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":169E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":1C24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":223D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":2846
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":2E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":36E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":3D79
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":4471
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":4B79
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":4E93
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":52E5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   2520
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":5737
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":5B89
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":5FDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":C275
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":C6C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":CB19
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":CC73
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":D0C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":F147
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":F2A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":F3FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":F84D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":1069F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmenu.frx":10AF1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Menumasterdata 
      Caption         =   "&Master Data"
      Begin VB.Menu menuGenralLedgerMaster 
         Caption         =   "&General Ledger Master"
      End
      Begin VB.Menu menusubleadgermaster 
         Caption         =   "&Sub Ledger Master"
      End
      Begin VB.Menu mnuDistict 
         Caption         =   "District Master"
      End
      Begin VB.Menu mnuCity 
         Caption         =   "City Master"
      End
      Begin VB.Menu mnuAgentMaster 
         Caption         =   "Agent Master"
      End
      Begin VB.Menu mnuTransport 
         Caption         =   "Transport Master ..."
      End
      Begin VB.Menu mnuBooksMast 
         Caption         =   "Books Master..."
      End
      Begin VB.Menu mnuBooksdet 
         Caption         =   "Book Details ..."
      End
      Begin VB.Menu mnuBookgp 
         Caption         =   "Book Group Master"
      End
      Begin VB.Menu mnuReportBookGp 
         Caption         =   "Report Book Group"
      End
      Begin VB.Menu mnuDiscount 
         Caption         =   "Discount Category .."
      End
      Begin VB.Menu mnuSubGpDiscount 
         Caption         =   "Sub Group Discount Category .."
      End
      Begin VB.Menu mnuInvoiceEnd 
         Caption         =   "Invoice End Part"
      End
      Begin VB.Menu menucreditnoteandpartmaster 
         Caption         =   "Credit Note End Part"
      End
      Begin VB.Menu mnuCounterEndPart 
         Caption         =   "Counter Sale End Part"
      End
      Begin VB.Menu mnuInvoiceEnd_basil 
         Caption         =   "Invoice End Part (basil)"
      End
      Begin VB.Menu mnuInvoiceEnd_basil_ret 
         Caption         =   "Invoice Ret.  End Part (basil)"
      End
      Begin VB.Menu mnuBKIssueEndP 
         Caption         =   "Book Issue End Part"
      End
      Begin VB.Menu mnuBKReturnEndP1 
         Caption         =   "Book Return End Part"
      End
      Begin VB.Menu mnuBookRec_IssueCat 
         Caption         =   "Book Receive/Issue (Category)"
      End
      Begin VB.Menu mnuGodown 
         Caption         =   "Godown Master.."
      End
      Begin VB.Menu mnuBookRet 
         Caption         =   "Book Retailers Details..."
      End
      Begin VB.Menu mnuLine_Godown 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFirm 
         Caption         =   "Firm Master..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPaper_size 
         Caption         =   "Paper Size Master..."
      End
      Begin VB.Menu mnuPaper_gsm 
         Caption         =   "GSM Master..."
      End
      Begin VB.Menu mnuPaper_Maker 
         Caption         =   "Paper Master..."
      End
      Begin VB.Menu mnuwastqty 
         Caption         =   "Wastage Master Qty. Wise .."
      End
      Begin VB.Menu mnuPBookMast 
         Caption         =   "Printing Book Master.."
      End
      Begin VB.Menu mnuSchoolMaster 
         Caption         =   "School Master"
      End
      Begin VB.Menu mnuTeacherDet 
         Caption         =   "Teacher Details"
      End
      Begin VB.Menu mnuopbalance 
         Caption         =   "Opening Balance"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBinder_printer 
         Caption         =   "Binder/Printer Details"
      End
      Begin VB.Menu mnuPrintOn 
         Caption         =   "Print On (Subject Printing)"
      End
      Begin VB.Menu mnuBookSt 
         Caption         =   "Book Status.."
         Visible         =   0   'False
      End
      Begin VB.Menu mnumzn 
         Caption         =   "Rep. add to Manager..."
      End
   End
   Begin VB.Menu menutranscationdata 
      Caption         =   "&Transaction"
      Begin VB.Menu menujournalvoucher 
         Caption         =   "&Voucher Entry"
      End
      Begin VB.Menu mnuPackingINV 
         Caption         =   "Packing Slip For (Invoice) ..."
      End
      Begin VB.Menu mnuPacking 
         Caption         =   "Packing Slip (Specimen)..."
      End
      Begin VB.Menu menusalesinvoice 
         Caption         =   "&Sales Invoice..."
      End
      Begin VB.Menu mnuSaleOrder 
         Caption         =   "Sale Order ..."
      End
      Begin VB.Menu mnuCreditNItem 
         Caption         =   "&Credit Note Item"
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
      Begin VB.Menu mnuIssuedBind 
         Caption         =   "Book Issue To Binder.."
      End
      Begin VB.Menu mnuBoromBinder 
         Caption         =   "Book Receive From Binder.."
      End
      Begin VB.Menu mnuBookRecFBind 
         Caption         =   "Book Issue/Receive"
      End
      Begin VB.Menu mnuStockTranstoGo 
         Caption         =   "Stock Trasn. To Godown"
      End
      Begin VB.Menu mnuInvBasil 
         Caption         =   "Invoice (Basil)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInvoiceBasilRet 
         Caption         =   "Invoice Return (Basil)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBookIssueSp 
         Caption         =   "Book Issue (Specimen)"
      End
      Begin VB.Menu mnuSpRet 
         Caption         =   "Book Return (Specimen)"
      End
      Begin VB.Menu mnuBookIssuetoSh 
         Caption         =   "Book Issue To School..."
      End
      Begin VB.Menu mnuAllotmentQty 
         Caption         =   "Rep. Wise Allotment Specimen Qty."
      End
      Begin VB.Menu mnuDonation 
         Caption         =   "Donation Entry..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPaper_TitlePrint 
         Caption         =   "Order For Title Printing..."
      End
      Begin VB.Menu mnuPaper_subjectprint 
         Caption         =   "Order For Subject Printing..."
      End
      Begin VB.Menu mnuPaper_NegativePrint 
         Caption         =   "Order For Negative Printing..."
      End
      Begin VB.Menu mnuPurchaseOrder 
         Caption         =   "Paper Purchase Order ..."
      End
      Begin VB.Menu mnuOrder 
         Caption         =   "Order Printing..."
      End
      Begin VB.Menu mnuPaerRec 
         Caption         =   "Paper Issue Entry..."
      End
      Begin VB.Menu mnuStockTrans 
         Caption         =   "Stock Transfar (Paper)..."
      End
      Begin VB.Menu mnuPaperSize 
         Caption         =   "Paper Ledger (Size Wise)..."
      End
      Begin VB.Menu mnuPaperPrint 
         Caption         =   "Paper Print Plan..."
      End
      Begin VB.Menu mnuSchoolWise 
         Caption         =   "School Wise Sale Return..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomerAgm 
         Caption         =   "Customer Agreement..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBookIssuetoSchool 
         Caption         =   "Book Issue To School (for Sampling)"
      End
   End
   Begin VB.Menu mnuBluePrint 
      Caption         =   "Transaction(&Blueprint)"
      Visible         =   0   'False
      Begin VB.Menu mnuinv_bluprint 
         Caption         =   "Invoice (Blueprint)..."
      End
      Begin VB.Menu mnuCashDis 
         Caption         =   "Cash Discount List..."
      End
   End
   Begin VB.Menu mnuLedger1 
      Caption         =   "Ledger"
      Begin VB.Menu mnuLedger 
         Caption         =   "Ledger(Blueprint)..."
      End
      Begin VB.Menu mnuLedBluePrint 
         Caption         =   "Ledger (Blue Print)..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSL_invoice 
         Caption         =   "SL Ledger View ..."
      End
      Begin VB.Menu mnuMulti 
         Caption         =   "Multi Mail Statement Option..."
      End
      Begin VB.Menu mnuSchoolLed 
         Caption         =   "School Ledger..."
      End
      Begin VB.Menu mnuLedger_conven 
         Caption         =   "Ledger View ..."
      End
      Begin VB.Menu mnuLedger_Basil 
         Caption         =   "Ledger View(Basil)"
      End
      Begin VB.Menu mnuAuthenticationIssueRec 
         Caption         =   "Authentication (Book Issue/Receive)"
      End
      Begin VB.Menu mnuDonnation 
         Caption         =   "Extra Discount Calculator"
      End
      Begin VB.Menu mnuSch_NetSales 
         Caption         =   "School Wise & Book Wise Net Sales..."
      End
      Begin VB.Menu mnuAdjOp 
         Caption         =   "Adjustment Option..."
      End
      Begin VB.Menu mnutod 
         Caption         =   "T.O.D. List ..."
      End
      Begin VB.Menu mnuApp 
         Caption         =   "Approval Form..."
      End
      Begin VB.Menu mnuBiltyStatus 
         Caption         =   "Bilty Status (Noida)..."
      End
      Begin VB.Menu mnuSaleRet 
         Caption         =   "SaleRetun Status(Noida)..."
      End
      Begin VB.Menu mnuSchoolLedgerSp 
         Caption         =   "School Ledger (For Specimen)..."
      End
      Begin VB.Menu mnu_Ledger1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckedBy 
         Caption         =   "Checked By"
      End
      Begin VB.Menu mnuAuditTrailLog 
         Caption         =   "Audit Trail Log"
      End
   End
   Begin VB.Menu menureports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuDisDaily 
         Caption         =   "Dispatch Report Daily..."
      End
      Begin VB.Menu mnuBookStock 
         Caption         =   "Book Stock Summary (Godown Wise)"
      End
      Begin VB.Menu mnuBinderStock 
         Caption         =   "Binder Stock Summary..."
      End
      Begin VB.Menu binderop 
         Caption         =   "Binder Book Opening..."
      End
      Begin VB.Menu menucashbook 
         Caption         =   "&Cash/Bank Book ..."
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
      Begin VB.Menu menugenledgertrialbalance 
         Caption         =   "Gen. Ledger Trial &Balance"
      End
      Begin VB.Menu menusubledgertrialbalance 
         Caption         =   "Sub. Ledger &Trial Balance"
      End
      Begin VB.Menu menugenledgeropentrialbalance 
         Caption         =   "Gen. Ledger &Opening Trial"
      End
      Begin VB.Menu mnugpsale 
         Caption         =   "Group Wise Sales..."
      End
      Begin VB.Menu menudistrictwisesales 
         Caption         =   "&State Wise/Rep Wise/District Wise/Book Wise Sales"
      End
      Begin VB.Menu mnupartywise 
         Caption         =   "Party Wise && Executive Wise Gross and Sale.."
      End
      Begin VB.Menu mnuPartyOuts 
         Caption         =   "Party Wise Outst. && T.Supply && T.Payment..."
      End
      Begin VB.Menu mnuRepTitlewise 
         Caption         =   "Rep. && Title Wise Net Qty Summary.."
      End
      Begin VB.Menu mnufullsession 
         Caption         =   "Full Session Wise Report"
         Begin VB.Menu mnuReports 
            Caption         =   "Reports"
         End
      End
      Begin VB.Menu mnuDistSales 
         Caption         =   "District &Wise Sales"
         Visible         =   0   'False
      End
      Begin VB.Menu mnudistrictwisesalesreturn 
         Caption         =   "District &Wise Sales Return "
      End
      Begin VB.Menu menuGroupwisesales 
         Caption         =   "&Group Wise Sales/Return"
      End
      Begin VB.Menu menubankadvice1 
         Caption         =   "Ba&nk Advice"
      End
      Begin VB.Menu mnubankadviceReconciliation 
         Caption         =   "Bank Advice Register"
      End
      Begin VB.Menu mnubiltyreturnregister 
         Caption         =   "Bilty Return Register..."
      End
      Begin VB.Menu mnuPartyList 
         Caption         =   "&Party List (Address)..."
      End
      Begin VB.Menu mnuDispachedreg 
         Caption         =   "&Dispatch Register"
      End
      Begin VB.Menu mnuBankReg 
         Caption         =   "Bank Register..."
      End
      Begin VB.Menu mnuCashReg 
         Caption         =   "Dispatch Register(C/M)..."
      End
      Begin VB.Menu mnuPartyProf 
         Caption         =   "Party Profile..."
      End
      Begin VB.Menu mnuBR 
         Caption         =   "Bilty Register (Transport Wise)..."
      End
      Begin VB.Menu conven_mnuBookwiseAgnLedger 
         Caption         =   "Book Wise Representative Ledger..."
      End
      Begin VB.Menu conven_mnuTotalBookQty 
         Caption         =   "Total Book Quantity/Amount..."
      End
      Begin VB.Menu conven_mnuTotalBookAmt 
         Caption         =   "Total Book Amount..."
         Visible         =   0   'False
      End
      Begin VB.Menu conven_mnuCollegeList 
         Caption         =   "College List..."
      End
      Begin VB.Menu conven_mnuAgentwiseIssue 
         Caption         =   "Agent Wise Issue..."
         Visible         =   0   'False
      End
      Begin VB.Menu conven_mnuDispatchReg 
         Caption         =   "Dispatch Register..."
      End
      Begin VB.Menu mnuAmtSaleNet 
         Caption         =   "Amount Wise Sales/Cash List..."
      End
      Begin VB.Menu mnuRepWiseSp 
         Caption         =   "Rep. Wise Specimen Summary..."
      End
      Begin VB.Menu mnuGST 
         Caption         =   "GST Related Report..."
      End
      Begin VB.Menu mnlbl 
         Caption         =   "Balance Sheet.."
      End
      Begin VB.Menu mnuAudit 
         Caption         =   "Stock Audit Report..."
      End
      Begin VB.Menu mnu120 
         Caption         =   "Balance 120 Days"
      End
      Begin VB.Menu mnuCourier 
         Caption         =   "Courier bill of entry..."
      End
      Begin VB.Menu mnuImp 
         Caption         =   "Imprest Register..."
      End
      Begin VB.Menu mnuTCS 
         Caption         =   "TCS Report..."
      End
      Begin VB.Menu mnublsheet 
         Caption         =   "Balance Sheet (Merging Report)"
      End
      Begin VB.Menu mnuLine_stcok 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuChangeMod 
      Caption         =   "&Change Module"
      Begin VB.Menu mnuinv 
         Caption         =   "Invoicing"
      End
      Begin VB.Menu mnicon 
         Caption         =   "Canvassing"
      End
      Begin VB.Menu mnustock 
         Caption         =   "Stock System"
      End
      Begin VB.Menu mnuPaper 
         Caption         =   "Paper"
      End
   End
   Begin VB.Menu menutools 
      Caption         =   "T&ools"
      Begin VB.Menu mnuSchool 
         Caption         =   "Modify School Name..."
      End
      Begin VB.Menu mnuTTrans 
         Caption         =   "Total Transaction Date Wise"
      End
      Begin VB.Menu mnuAuth 
         Caption         =   "Authorisation Option"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuClosing 
         Caption         =   "Closing Transfer"
         Visible         =   0   'False
      End
      Begin VB.Menu menubackupdata 
         Caption         =   "&Backup Data"
         Visible         =   0   'False
      End
      Begin VB.Menu menuRestoredata 
         Caption         =   "&Restore data"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search"
      End
      Begin VB.Menu mnuDq 
         Caption         =   "Grid Setting"
         Visible         =   0   'False
      End
      Begin VB.Menu sep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuuserright 
         Caption         =   "User Rights"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "Update Data On Net...."
         Visible         =   0   'False
      End
      Begin VB.Menu menuchangepassword 
         Caption         =   "Change &Password"
      End
      Begin VB.Menu mnuMail 
         Caption         =   "Mail Status..."
      End
      Begin VB.Menu mnuSms 
         Caption         =   "SMS Status..."
      End
      Begin VB.Menu menusetup 
         Caption         =   "&Setup"
      End
      Begin VB.Menu mnuSetColor 
         Caption         =   "Set Color"
      End
      Begin VB.Menu mnuBalance 
         Caption         =   "Closing Balance Transfar..."
      End
      Begin VB.Menu mnubk 
         Caption         =   "Data Backup option"
      End
   End
   Begin VB.Menu mnuref 
      Caption         =   "Refresh Data"
   End
   Begin VB.Menu mnuReminder 
      Caption         =   "Reminder Message"
   End
   Begin VB.Menu mnuDasboard 
      Caption         =   "Dashboard"
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
Dim I As Integer
Dim c_name As String
'Dim con_LAST As ADODB.Connection
Private Sub alphaSubLedgerAccountsmnu_Click()
 
 s1 = "7"
 SLEDGERPRINT.Show
 SLEDGERPRINT.Label4.Visible = False
 SLEDGERPRINT.Combosubledger.Visible = False
 SLEDGERPRINT.Label5.Visible = True
 SLEDGERPRINT.Alpha.Visible = True
   
End Sub
Private Sub binderop_Click()

bookOp = "y"
frmBinderStock.Show

End Sub
Private Sub cmdBasil_Click()

module_ = "Basil"
permissionUserWise
cmdConven.BackColor = &HC0E0FF
cndInv.BackColor = &HC0E0FF
cmdStock.BackColor = &HC0E0FF
cmdPaper.BackColor = &HC0E0FF
user_with_module "add", , module_
cmdBasil.BackColor = vbWhite

End Sub
Private Sub cmdConven_Click()

module_ = "Canvassing"
permissionUserWise

cmdPaper.BackColor = &HC0E0FF
cmdBasil.BackColor = &HC0E0FF
cmdStock.BackColor = &HC0E0FF
cndInv.BackColor = &HC0E0FF

user_with_module "add", , module_
cmdConven.BackColor = vbWhite


End Sub

Private Sub cmdPaper_Click()
  
    module_ = "Paper"
    permissionUserWise
    
    cmdBasil.BackColor = &HC0E0FF
    cmdConven.BackColor = &HC0E0FF
    cndInv.BackColor = &HC0E0FF
    cmdStock.BackColor = &HC0E0FF
    
    user_with_module "add", , module_
    
    cmdPaper.BackColor = vbWhite

End Sub

Private Sub cmdStock_Click()


module_ = "Stock System"
permissionUserWise

cmdPaper.BackColor = &HC0E0FF
cmdBasil.BackColor = &HC0E0FF
cmdConven.BackColor = &HC0E0FF
cndInv.BackColor = &HC0E0FF
user_with_module "add", , module_
cmdStock.BackColor = vbWhite


End Sub

Private Sub cndInv_Click()
module_ = "Invoicing"
permissionUserWise

cmdPaper.BackColor = &HC0E0FF
cmdBasil.BackColor = &HC0E0FF
cmdStock.BackColor = &HC0E0FF
cmdConven.BackColor = &HC0E0FF

user_with_module "add", , module_

cndInv.BackColor = vbWhite


End Sub
Function user_with_module(Optional str_ As String, Optional butten_ As String, Optional user_ As String) As String

'add => new user Addition
'search => user module search

cmdPaper.BackColor = &HC0E0FF
cmdBasil.BackColor = &HC0E0FF
cmdStock.BackColor = &HC0E0FF
cmdConven.BackColor = &HC0E0FF
cndInv.BackColor = &HC0E0FF

If RS.State = 1 Then RS.close
RS.Open "select _user,_module from userwithModule where (_user='" & UserName & "')", con, adOpenDynamic, adLockOptimistic

If str_ = "add" Then

    If RS.EOF = True Then
    RS.AddNew
    End If
    
    RS.Fields(0).value = UserName
    RS.Fields(1).value = module_
    RS.update

Else
'===========================================================
If RS.EOF = True Then
   cndInv.BackColor = vbWhite
   Exit Function
End If

If RS(1) = "Paper" Then
   
   Call cmdPaper_Click
   cmdPaper.BackColor = vbWhite
   
ElseIf RS(1) = "Stock System" Then

   Call cmdStock_Click
   cmdStock.BackColor = vbWhite

ElseIf RS(1) = "Invoicing" Then

   Call cndInv_Click
   cndInv.BackColor = vbWhite

ElseIf RS(1) = "Canvassing" Then

   Call cmdConven_Click
   cmdConven.BackColor = vbWhite

ElseIf RS(1) = "Basil" Then

   Call cmdBasil_Click
   cmdBasil.BackColor = vbWhite

End If



End If

    
    
End Function

Private Sub conven_mnuAgentwiseIssue_Click()
  frmIssuerpt.Show
End Sub
Private Sub conven_mnuBookwiseAgnLedger_Click()
  frmAgentLadger.Show
End Sub
Private Sub conven_mnuCollegeList_Click()
  frmCollegeList.Show
End Sub
Private Sub conven_mnuDispatchReg_Click()
  frmCash.Show
End Sub
Private Sub conven_mnuTotalBookAmt_Click()

'frmAgentLadger.cr.Reset
'  frmAgentLadger.cr.ReportFileName = rptPath & "/TotalBookLedgerTotal.rpt"
'   frmAgentLadger.cr.Connect = "filedsn=chitradsn;uid= " & sql_user  & ";pwd=" & sql_pass
'    frmAgentLadger.cr.WindowShowPrintBtn = True
'    frmAgentLadger.cr.WindowShowPrintSetupBtn = True
'   frmAgentLadger.cr.WindowShowSearchBtn = True
'  frmAgentLadger.cr.WindowState = crptMaximized
'frmAgentLadger.cr.Action = 1

End Sub
Private Sub conven_mnuTotalBookQty_Click()

frmTotBookQty_Amt.Show 1


    
End Sub
Private Sub EXIT_Click()
'Form1.Show
''If MsgBox("Are you Sure, want to Exit. ", vbInformation + vbYesNo) = vbYes Then
   End
''End If

End Sub
Private Sub frmChallan1_Click()
    Toolbar1.Visible = False
    frmChallan.Show
End Sub
Sub hideMenu()


'Dim o As Object
'For Each o In MainMenu
'If TypeOf o Is Menu Then
'   o.Name.Visible = False
'End If
'Next



End Sub
Private Sub MDIForm_Load()


If RS.State = adStateOpen Then RS.close
RS.Open "select * from setup1 where fyear='" & main.session & "'", con, adOpenKeyset, adLockReadOnly, adCmdText
If Not RS.EOF Then
    
    Dim I As Integer
    Dim strcname As String
    from_date = RS!yarfrom
    to_date = RS!yarto
    
    If Not IsNull(RS!phone3) Then
       donnation_visible = RS!phone3        'for visible donnation menu
    Else
       donnation_visible = "n"
    End If
    
    For I = 0 To RS.RecordCount - 1
    strcname = strcname & RS!setupid & " (" & RS!cname & ")" & ","
        If Not RS.EOF Then
            RS.MoveNext
        End If
    Next
    strcname = Left(strcname, Len(strcname) - 1)
    arycname = Split(strcname, ",")
    
 
End If
RS.close




Dim dd_ As String

dd_ = "31/03/" & Year(to_date) - 1
dd_ = Format(dd_, "MM/dd/yyyy")



con.Execute "update sledger set FromDate='" & dd_ & "'"
con.Execute "update BookOpening set dates='" & from_date & "'"
con.Execute "update BookOpening_ns set dates='" & from_date & "'"

con.Execute "update BookDiff set dates='" & from_date & "'"

con.Execute "update BOOKS set GROUPCODE_sub=null where groupcode_sub =''"

user_with_module "search", , UserName

BackColorFrom Me
menuImage
permissionUserWise

If UCase(UserName) = UCase("Admin") Then
  mnuuserright.Visible = True
  menusetup.Visible = True
  mnuSetColor.Visible = True
  mnuUpdate.Visible = True
  'mnudata.Visible = True
Else
  mnuuserright.Visible = False
  menusetup.Visible = False
  mnuSetColor.Visible = False
  mnuUpdate.Visible = False
  'mnudata.Visible = False
End If

mnuBalance.Visible = False
If UCase(UserName) = UCase("v") Then
  mnuBalance.Visible = True
  mnuuserright.Visible = True
  menusetup.Visible = True
End If

'==================================================================================================
On Error Resume Next

If checkPermission("donnation") = True Then
   con.Execute "exec tmpdata " & UId & ""
End If

If checkPermission("adj") = True Then
   con.Execute "exec tmpdata_saleadj"
End If

'==================================================================================================
'"By RASHI SOFTWARE SERVICES (Contact No : 9997314681,7906524313)"
MainMenu.Caption = "Publication Software System(" & session & ")" & "By RASHI SOFTWARE SERVICES (Mob: 9997314681)"

If (session = "2015-16" Or session = "2016-17") Then
   mnuApp.Visible = False
Else
    
    
    Set con_LAST = New ADODB.Connection
    Last_database = "database=chitraData_" & database_last
    If LCase(server_) = "server" Then
       con_LAST.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew & "; " & Last_database & "; UID=" & sql_user & "; PWD=" & sql_pass
    Else
       con_LAST.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew & "; " & Last_database & ";UID=; PWD=" & sql_pass
    End If
    
    DoEvents
    DoEvents
    
    
    con_LAST.CursorLocation = adUseClient
    If con_LAST.State = 1 Then con_LAST.close
    con_LAST.Open
    DoEvents
    DoEvents
    
    
    a1 = Right(session, 2) - 3
    a2 = Right(session, 2) - 2
   

    Set con_LAST2 = New ADODB.Connection
    Last_database2 = "database=chitraData_" & a1 & "" & a2


    If LCase(server_) = "server" Then
       con_LAST2.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew & "; " & Last_database2 & "; UID=" & sql_user & "; PWD=" & sql_pass
    Else
       con_LAST2.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew & "; " & Last_database2 & ";UID=; PWD=" & sql_pass
    End If
    
    DoEvents
    DoEvents
    
    
    con_LAST2.CursorLocation = adUseClient
    If con_LAST2.State = 1 Then con_LAST2.close
    con_LAST2.Open
    
    DoEvents
    DoEvents



    
    
End If


financialyear



If Right(session, 2) >= 20 Then
If RS.State = 1 Then RS.close
    RS.Open "select Id from reminderTbl where (Show='n' and convert(smalldatetime,invoicedate,101) = convert(smalldatetime,'" & Date & "',103)) order by Id,INVOICEDATE", con
    If RS.RecordCount > 0 Then
       If rs1.State = 1 Then rs1.close
       rs1.Open "select reminder from UsrePermission where UserName='" & UserName & "'", coninfo
       If rs1.EOF = False Then
          If rs1!reminder = "y" Then
             frmReminder.Show
          End If
       End If
End If

End If


mnuDonnation.Visible = True

If session = "2024-25" Then

    mnuDonnation.Visible = True

Else

    If RS.State = 1 Then RS.close
    RS.Open "select top 1 dno from DonnationMain", con
    If RS.RecordCount = 0 Then
       mnuDonnation.Visible = False
    End If

End If



mnuCheckedBy.Visible = False
mnuAuditTrailLog.Visible = False

If AuditTrail = "y" Then

mnuCheckedBy.Visible = True
mnuAuditTrailLog.Visible = True

End If


End Sub
Sub menuImage()
End Sub
Sub permissionUserWise()
    
Dim D As Integer
Dim o As Object

MainMenu.Caption = "Publication Software System : (" & module_ & " - " & session & ")"
'BackColorFrom MainMenu
    
mnuBookIssuetoSchool.Visible = False
mnuPurchaseOrder.Visible = False

mnuBookRet.Visible = False
'=========================
mnuwastqty.Visible = False

mnuCustomerAgm.Visible = False
mnuSchoolWise.Visible = False
mnuSchoolLedgerSp.Visible = False
mnupartywise.Visible = False
mnuRepWiseSp.Visible = False
mnuMulti.Visible = False

menuGenralLedgerMaster.Visible = False
menusubleadgermaster.Visible = False
mnuDistict.Visible = False
mnuCity.Visible = False
mnuAgentMaster.Visible = False
mnuTransport.Visible = False
mnuBooksMast.Visible = False
mnuBookgp.Visible = False
mnuReportBookGp.Visible = False
mnuDiscount.Visible = False
mnuSubGpDiscount.Visible = False
mnuSchoolMaster.Visible = False
mnuTeacherDet.Visible = False
menujournalvoucher.Visible = False
menusalesinvoice.Visible = False
mnuCreditNItem.Visible = False
menucountersale.Visible = False
menucreditnote.Visible = False
menudebitnote.Visible = False
mnuIssuedBind.Visible = False
mnuBoromBinder.Visible = False
mnuBookRecFBind.Visible = False
mnuStockTranstoGo.Visible = False
mnuInvBasil.Visible = False
mnuInvoiceBasilRet.Visible = False
mnuBookIssueSp.Visible = False
mnuSpRet.Visible = False
mnuBookIssuetoSh.Visible = False
mnuDonation.Visible = False
mnuBookStock.Visible = False
mnuBinderStock.Visible = False


menucashbook.Visible = False
mnudistrictwisesalesreturn.Visible = False
menugeneralledgeraccounts.Visible = False
menusubledgeraccounts.Visible = False
alphaSubLedgerAccountsmnu.Visible = False
menugenledgertrialbalance.Visible = False
menusubledgertrialbalance.Visible = False
menugenledgeropentrialbalance.Visible = False
mnugpsale.Visible = False
menudistrictwisesales.Visible = False
menuGroupwisesales.Visible = False
menubankadvice1.Visible = False
mnubankadviceReconciliation.Visible = False
mnubiltyreturnregister.Visible = False
mnuPartyList.Visible = False
mnuDispachedreg.Visible = False
mnuBankReg.Visible = False
mnuCashReg.Visible = False
mnuPartyProf.Visible = False
mnuSL_invoice.Visible = False
mnuBR.Visible = False
conven_mnuBookwiseAgnLedger.Visible = False
conven_mnuTotalBookQty.Visible = False
conven_mnuTotalBookAmt.Visible = False
conven_mnuCollegeList.Visible = False
conven_mnuAgentwiseIssue.Visible = False
conven_mnuDispatchReg.Visible = False
mnuBinder_printer.Visible = False
mnuAmtSaleNet.Visible = False
mnuApp.Visible = False
mnuBiltyStatus.Visible = False
mnuSaleRet.Visible = False

'---End Part--------------------------------
mnuInvoiceEnd.Visible = False
menucreditnoteandpartmaster.Visible = False
mnuCounterEndPart.Visible = False
mnuInvoiceEnd_basil.Visible = False
mnuInvoiceEnd_basil_ret.Visible = False
mnuBKIssueEndP.Visible = False
mnuBKReturnEndP1.Visible = False

mnuBookRec_IssueCat.Visible = False
mnuPaper_size.Visible = False
mnuPaper_gsm.Visible = False
mnuPaper_Maker.Visible = False
mnuPBookMast.Visible = False
mnuGodown.Visible = False

mnuSchoolMaster.Visible = False
mnuTeacherDet.Visible = False
mnuLedger.Visible = False

'''''paper---------------
mnuPaper_NegativePrint.Visible = False
mnuPaper_TitlePrint.Visible = False
mnuPaper_subjectprint.Visible = False
mnuLedger.Visible = False
mnuLedger_conven.Visible = False
mnuOrder.Visible = False
mnuPaerRec.Visible = False
mnuStockTrans.Visible = False
mnuPaperSize.Visible = False
mnuBooksdet.Visible = False
mnuBinder_printer.Visible = False
mnuPrintOn.Visible = False

'========================
'Stock System

mnuBookRec_IssueCat.Visible = False
mnuIssuedBind.Visible = False
mnuBoromBinder.Visible = False
mnuBookRecFBind.Visible = False
mnuStockTranstoGo.Visible = False

'-----------------------
mnuLedger_Basil.Visible = False
mnuPackingINV.Visible = False
mnuPacking.Visible = False
mnuSaleOrder.Visible = False
conven_mnuTotalBookQty.Visible = False
mnuAuthenticationIssueRec.Visible = False
'mnuSchoolLed.Visible = False

mnuDonnation.Visible = False
mnuSch_NetSales.Visible = False
mnuAmtSaleNet.Visible = False

mnuAdjOp.Visible = False
mnutod.Visible = False
mnuPaperPrint.Visible = False
mnlbl.Visible = False

mnuApp.Visible = False
mnu120.Visible = False



Open App.Path + "\dn.neo" For Input As #1
Line Input #1, donnation_
Close #1

'If donnation_ = "1111" Then
'   mnuDonnation.Visible = True
'   Exit Sub
'End If


'============================
If module_ = "Invoicing" Then

    If donnation_visible = "y" Then
       mnuDonnation.Visible = True
    Else
       mnuDonnation.Visible = False
    End If

End If





'============================
'============================


If RS.State = 1 Then RS.close
RS.Open "select taskname,tasktype,permission,Xwidth from UsrePermission where ([module]='" & module_ & "' and username= '" & main.UserName & "') order by tasktype", coninfo, adOpenKeyset, adLockReadOnly
While RS.EOF = False
For Each o In MainMenu
If o.Name = RS!taskType Then
   
   If RS!Permission = "y" Then
      o.Visible = True
   Else
      o.Visible = False
   End If
   
   If RS!Xwidth = "hide" Then
      o.Visible = False
   End If
   
End If
Next
RS.MoveNext
Wend



On Error Resume Next
mnustock.Visible = False
mnuinv.Visible = False
mnuBasil.Visible = False
mnicon.Visible = False
mnuPaper.Visible = False
mnustock.Visible = False
mnuSchoolLed.Visible = False
    
If RS.State = 1 Then RS.close

'''RS.Open "select Module from UsrePermission where username= '" & main.UserName & "' group by Module", con, adOpenKeyset, adLockReadOnly
RS.Open "select [Module] from UsrePermission where username= '" & main.UserName & "' group by [Module]", coninfo, adOpenKeyset, adLockReadOnly
While RS.EOF = False

If RS(0) = "Invoicing" Then
   mnuinv.Visible = True
   mnuGST.Visible = True
End If

If RS(0) = "Basil" Then
   mnuBasil.Visible = True
   mnuGST.Visible = False
End If

If RS(0) = "Canvassing" Then
   mnicon.Visible = True
   mnuGST.Visible = False
End If

If RS(0) = "Paper" Then
   mnuPaper.Visible = True
   mnuGST.Visible = False
End If

If RS(0) = "Stock System" Then
   mnustock.Visible = True
   mnuGST.Visible = False
End If

RS.MoveNext
Wend



If RS.State = 1 Then RS.close
RS.Open "select distinct [Module] from UsrePermission where username= '" & main.UserName & "' and  [Module]='" & module_ & "'", coninfo, adOpenKeyset, adLockReadOnly
If RS.EOF = False Then
    If RS(0) = "Invoicing" Then
       mnuSchoolLed.Visible = True
       mnuGST.Visible = True
       mnlbl.Visible = True
    ElseIf RS(0) = "Paper" Then
       mnuPaperPrint.Visible = True
    End If
End If


'------------------------------------
conven_mnuCollegeList.Visible = False
'------------------------------------

If RS.State = 1 Then RS.close
RS.Open "select top 1 DNo from DonnationMain", con, adOpenKeyset, adLockReadOnly
If RS.EOF = True Then
   mnuDonnation.Visible = False
Else
   mnuDonnation.Visible = True
End If

End Sub
Private Sub menuAgentMaster_Click()
Toolbar1.Visible = False
bookmaster.SSTab1.Tab = 3
bookmaster.Show
bookmaster.Commandmasteradd.SetFocus
End Sub
Sub DSN()


'''Dim FSO As FileSystemObject
'''Dim f As File
'''Dim txt As TextStream
'''Dim matter As String
'''Dim total As String
'''Dim s(1, 2) As String
'''Set FSO = New FileSystemObject
'''Dim ss, database_ As String
'''
'''matter = ""
'''
'''Dim op_system, Dusername, dstrpath As String
'''
'''login.GetCom
'''
'''If RS.State = 1 Then RS.close
'''RS.Open "select * from UserDSN where UserName='" & com_name & "' and usrename_='" & com_user & "'", CCON
'''If RS.EOF = False Then
'''
'''  dstrpath = RS!Path '& "\chitradsn.dsn"
'''  Set txt = FSO.CreateTextFile("" & dstrpath)
'''  'Label1.Caption = dstrpath
'''
'''End If
'''
'''If RS.State = 1 Then RS.close
'''RS.Open "select UID,WSID,[SERVER] from data", CCON
'''If RS.EOF = False Then
'''   ss = "Server=" & RS!server
'''   serverName_ = Mid(ss, 8)
'''   sql_pass = RS!WSID & ""
'''End If
'''
'''
'''
'''If session = "2016-17" Then
'''database_ = "Database=chitraData_1617"
'''ElseIf session = "2015-16" Then
'''database_ = "Database = chitraData"
'''ElseIf session = "2017-18" Then
'''database_ = "Database=chitraData_1718"
'''ElseIf session = "2018-19" Then
'''database_ = "Database=chitraData_1819"
'''
'''End If
'''
'''
'''If LCase(server_) = "client" Then
'''    matter = matter & "[ODBC]" & vbNewLine
'''    matter = matter & "DRIVER=SQL Server" & vbNewLine
'''    matter = matter & "UId=dinesh" & vbNewLine
'''    matter = matter & "Trusted_Connection=Yes" & vbNewLine
'''    matter = matter & "Network=DBMSLPCN" & vbNewLine
'''    matter = matter & database_ & "" & vbNewLine
'''    matter = matter & "WSID=COMPAQ" & vbNewLine
'''    matter = matter & "APP=Microsoft Data Access Components" & vbNewLine
'''    matter = matter & ss & "" & vbNewLine
'''    txt.Write matter
'''    txt.close
'''Else
'''    matter = matter & "[ODBC]" & vbNewLine
'''    matter = matter & "DRIVER=SQL Server" & vbNewLine
'''    matter = matter & "UId=sa" & vbNewLine
'''    matter = matter & database_ & "" & vbNewLine
'''    matter = matter & "WSID=COMPAQ" & vbNewLine
'''    matter = matter & "APP=Microsoft Data Access Components" & vbNewLine
'''    matter = matter & ss & "" & vbNewLine
'''    txt.Write matter
'''    txt.close
'''End If

End Sub
Private Sub MDIForm_Resize()

DSN


End Sub
Private Sub MDIForm_Unload(cancel As Integer)
'If MsgBox("Are you Sure, want to Exit. ", vbInformation + vbYesNo) = vbYes Then
'   End
'Else
'   cancel = 1
'End If

End Sub

Private Sub menubackupdata_Click()
frmbackup.Show
End Sub
Private Sub menubankadvice1_Click()
bankadvice.Show
End Sub
Private Sub menuBatch_Click()
  s2 = 2
  frmRegsiterList.Show
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
 CASHBOOK.Show
End Sub
Private Sub menuchangepassword_Click()
  'Toolbar1.Visible = False
   frmChangePass.Show
End Sub
Private Sub menucountersale_Click()
countersale.Show
End Sub
Private Sub menucreditnote_Click()
mnuMenu_ = "menucreditnote"
Creditnotefile.Show
'frmCrNote.Show

End Sub
Private Sub menucreditnoteandpartmaster_Click()

'Toolbar1.Visible = False
'master.SSTab1.Tab = 3
'master.Show

popupvalue5 = UCase("credititem")
frmEndPart.Show

End Sub
Private Sub menucreditnoteitems_Click()
Toolbar1.Visible = False
crtitem.Show
End Sub
Private Sub menudebitnote_Click()
    mnuMenu_ = "menudebitnote"
    Debitnotefile.Show
End Sub
Private Sub menudiscountcategorymaster_Click()
    
    'Toolbar1.Visible = False
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
    'DWsales.Show
    frmState_Dist_RepWise.Show
End Sub
Private Sub menugeneralledgeraccounts_Click()
s1 = "6"
GLEDGERPRINT.Show
End Sub
Private Sub menugenledgeropentrialbalance_Click()
GOptrial.Show
End Sub
Private Sub menugenledgertrialbalance_Click()
s1 = "8"
Gentrial.Show
End Sub
Private Sub menuGenralLedgerMaster_Click()
'Toolbar1.Visible = False
frmGLedger.Show
End Sub
Private Sub menuGroupwisesales_Click()
groupwisesales.Show
End Sub
Private Sub menuInvoiceandpartmaster_Click()
    Toolbar1.Visible = False
    master.SSTab1.Tab = 2
    master.Show
End Sub
Private Sub menuItem_Click()
  ss1 = 1
  frmItemSale.Show
End Sub
Private Sub menuItem1_Click()
    Toolbar1.Visible = False
    frmIssueItem.Show
End Sub
Private Sub menujournalvoucher_Click()
    Screen.MousePointer = vbHourglass
    Voucherform.Show
    Screen.MousePointer = vbDefault
End Sub
Private Sub menupaymentvoucher_Click()
    Toolbar1.Visible = False
    Voucherform.vtype.text = "P"
    Voucherform.Show
End Sub
Private Sub menuReceiptvoucher_Click()
    Toolbar1.Visible = False
    Voucherform.vtype.text = "R"
    Voucherform.Show
End Sub
Private Sub menuPurchase_Click()
    Toolbar1.Visible = False
    frmPurchase.Show
End Sub
Private Sub menuQualti_Click()
    ss1 = 4
    frmItemSale.Show
End Sub
Private Sub menuSale_Click()
    s2 = 1
    frmRegsiterList.Show
End Sub
Private Sub menuSales_Click()
    ss1 = 2
    frmItemSale.Show
End Sub
Private Sub menuRestoredata_Click()
    frmRestoredata.Show
End Sub
Private Sub menusalesinvoice_Click()
    s1 = 1
    mnuMenu_ = "menusalesinvoice"
    invoice.Show
End Sub
Private Sub menuSattion_Click()
    frmBalancereg.Show
End Sub
Private Sub menusetup_Click()
    setup.Show
End Sub
Private Sub menuStation_Click()
 ss1 = 3
 frmItemSale.Show
End Sub
Private Sub menuStock_Click()
    frmFinishItemStockSumary.Show
End Sub
Private Sub menusubleadgermaster_Click()
    'Toolbar1.Visible = False
    'master.SSTab1.Tab = 1
    'master.Show
    frmSubledger.Show
End Sub
Private Sub menusubledgeraccounts_Click()
    s1 = "7"
    SLEDGERPRINT.Show
    SLEDGERPRINT.Label4.Visible = True
    SLEDGERPRINT.Combosubledger.Visible = True
    SLEDGERPRINT.Label5.Visible = False
    SLEDGERPRINT.Alpha.Visible = False
End Sub
Private Sub menusubledgertrialbalance_Click()
    s1 = "9"
    subtrial.Show
End Sub
Private Sub menuVCashendpartmaster_Click()
    Toolbar1.Visible = False
    master.SSTab1.Tab = 5
    master.Show
End Sub
Private Sub mnicon_Click()
    module_ = "Canvassing"
    permissionUserWise
    conven_mnuTotalBookAmt.Visible = False
End Sub

Private Sub mnlbl_Click()
    frmBalanceSheet.Show
End Sub
Private Sub mnu120_Click()
    frmMonthlyDeb.Show
End Sub
Private Sub mnuAdjOp_Click()
    frmSalesAdjustment.Show
End Sub
Private Sub mnuAgentMaster_Click()
 frmAgnMaster.Show 1
End Sub
Private Sub mnuAnnexureReport_Click()
 frmAnnexure.Show 1
End Sub
Private Sub mnuArea_Click()
 frmArea.Show 1
End Sub
Private Sub mnuAllotmentQty_Click()
frmBookAllotment.Show
End Sub
Private Sub mnuAmtSaleNet_Click()
 frmAmount.Show
End Sub
Private Sub mnuApp_Click()
 frmApproval.Show
End Sub
Private Sub mnuAudit_Click()
    frmAuditRpt.Show
End Sub
Private Sub mnuAuditTrailLog_Click()
    frmAuditTrailLog.Show
End Sub
Private Sub mnuAuth_Click()
    frmBillList.Show
End Sub
Private Sub mnuAuthenticationIssueRec_Click()
 change_Pass = "a"
 frmLoginl.Show
End Sub
Private Sub mnuBalance_Click()
If LCase(UserName) = "v" Then
 Form1.Show
End If
End Sub
Private Sub mnubankadviceReconciliation_Click()
   frmAdviceStatus.Show
End Sub
Private Sub mnuBankR_Click()
   frmBankRecon.Show
End Sub
Private Sub mnuBankReg_Click()
   frmBank.Show
End Sub
Private Sub mnuBasil_Click()
   module_ = "Basil"
   permissionUserWise
End Sub
Private Sub mnubiltyreturnregister_Click()
   BILTYRETURNREGISTER.Show
End Sub

Private Sub mnuBiltyStatus_Click()
frmBiltyStatus.Show
End Sub

Private Sub mnuBinder_printer_Click()
   'HeadTbl = "binder"
   'frmMasters.Show 1
   frmGodown.Show
End Sub

Private Sub mnuBinderStock_Click()
 bookOp = "n"
 frmBinderStock.Show
End Sub

Private Sub mnubk_Click()
   frmDataBK.Show
End Sub

Private Sub mnuBKIssueEndP_Click()
   popupvalue5 = UCase("invoice_sp")
   frmEndPart.Show
End Sub
Private Sub mnuBKReturnEndP1_Click()
   popupvalue5 = UCase("invoice_spret")
   frmEndPart.Show
End Sub

Private Sub mnublsheet_Click()
frmBSheet.Show
End Sub

Private Sub mnuBookIssueSp_Click()
   frmBookIssueSp.Show
End Sub

Private Sub mnuBookIssuetoSchool_Click()
frmBookIssueToSchool.Show
End Sub
Private Sub mnuBookIssuetoSh_Click()
  issueagent.Show
End Sub

Private Sub mnuBookRet_Click()
  frmBookRetailers.Show
End Sub

Private Sub mnuBooksdet_Click()
  frmbook.Show
End Sub
Private Sub mnuBooksMast_Click()
   frmBookMaster.Show
End Sub
Private Sub mnuBR_Click()
   frmTransportWise_Bilty.Show
End Sub
Private Sub mnuCashDis_Click()
  frmCashDis.Show
End Sub
Private Sub mnuCashReg_Click()
  frmCash_inv.Show
End Sub
Private Sub mnuChangeMod_Click()
'frmChangeModule.Show
End Sub

Private Sub mnuCheckedBy_Click()
frmChecked.Show
End Sub

Private Sub mnuCity_Click()
  frmCityMaster.Show
End Sub
Private Sub mnuBookgp_Click()
  frmBookGpMaster.Show
End Sub
Private Sub mnuBookRec_IssueCat_Click()
  frmIssue_ReceiceMaster.Show
End Sub
Private Sub mnuBookRecFBind_Click()
  IssueBook = "Issue"
  Unload frmBookIssue
  frmBookIssue.Show
End Sub

Private Sub mnuBookStock_Click()
frmBookStock.Show

End Sub

Private Sub mnuBoromBinder_Click()
frmBinderRecChallan.Show
End Sub

Private Sub mnuClosing_Click()
FRMOPTRANS.Show
End Sub

Private Sub mnuCounterEndPart_Click()
popupvalue5 = UCase("Cash")
frmEndPart.Show
End Sub

Private Sub mnuCrDrRegister_Click()
frmDebit_CreditNotReg.Show
End Sub

Private Sub mnuCreation_Click()
Toolbar1.Visible = False
InvoiceChallane.Show
End Sub
Private Sub mnuctTaxStructure_Click()
  StateTaxList.Show
End Sub
Private Sub mnuCurMaster_Click()
frmCurValue.Show
End Sub

Private Sub mnuCourier_Click()
frmCouriorList.Show
End Sub

Private Sub mnuCreditNItem_Click()
s1 = 12
Critnote.Show
End Sub
Private Sub mnudata_Click()
Screen.MousePointer = vbHourglass
con.Execute ("exec FatchDataFromSP")
Screen.MousePointer = vbDefault
End Sub
Private Sub mnudealer_Click()
 dealer_ss.Show
End Sub
Private Sub mnuDepart_Click()
 frmDept.Show
End Sub
Private Sub mnuDet_Click()
 frmPayment_Rec_Jen.Show
End Sub

Private Sub mnuCustomerAgm_Click()
frmAgreement.Show
End Sub

Private Sub mnuDasboard_Click()

DoEvents
DoEvents


strProgramName = "\\192.168.0.140\blueprintSales\BlueprintDasboard\bin\Debug\BlueprintDasboard.exe"
'strArgument = "/G"
'
'strProgramName = "C:\SoftwareCode_DotNet_2021\BluePrintDasBoard\BlueprintDasboard\bin\Debug\BlueprintDasboard.exe"
'
'''Call Shell("""" & strProgramName & """ """ & strArgument & """", vbNormalFocus)

Call Shell(strProgramName, vbNormalFocus)


End Sub

Private Sub mnuDiscount_Click()
 frmDiscountCat.Show
End Sub

Private Sub mnuDisDaily_Click()
frmDisPatchDaily.Show
End Sub

Private Sub mnuDispachedreg_Click()
 frmDispatched.Show
End Sub
Private Sub mnuDistict_Click()
 frmDistrict.Show 1
End Sub
Private Sub mnudistrictwisesalesreturn_Click()
 DWSalesReturn.Show
End Sub
Private Sub mnuDSR_Click()
 frmDSR.Show
End Sub

Private Sub mnuDistSales_Click()
DWsales.Show
End Sub

Private Sub mnuDonation_Click()
 frmSchoolPay.Show
End Sub

Private Sub mnuDonnation_Click()
frmDonnation.Show
End Sub
Private Sub mnuDq_Click()
 frmDQry.Show
End Sub
Private Sub mnuexSale_Click()
 s2 = 4
 frmRegsiterList.Show
End Sub
Private Sub mnugp_Click()
 frmProdust.Show
End Sub

Private Sub mnuFirm_Click()
frmFirmMaster.Show
End Sub

Private Sub mnuGodown_Click()
 frmGodown.Show
End Sub
Private Sub mnuGodownList_Click()
 frmCash_SalesReports.Show
End Sub
Private Sub mnugpsale_Click()
frmGPSales.Show
End Sub
Private Sub mnuGrid_Click()
 frmGrid.Show
End Sub

Private Sub mnuIM_Click()
Toolbar1.Visible = False
frmMsgItem.Show
End Sub

Private Sub mnuM_11_Click()
frmTypeProduct.Show
End Sub

Private Sub mnuGST_Click()

frmGSTrpt.Show 1
End Sub

Private Sub mnuImp_Click()
frmImprest.Show
End Sub

Private Sub mnuinv_bluprint_Click()
frmInvoice_blueprint.Show
End Sub

Private Sub mnuINV_Click()
module_ = "Invoicing"
permissionUserWise
End Sub
Private Sub mnuInvBasil_Click()
frmBasilSales.Show
End Sub
Private Sub mnuInvoiceBasilRet_Click()
    frmBasilSales_Ret.Show
End Sub
Private Sub mnuInvoiceEnd_basil_Click()
    popupvalue5 = UCase("cashbasil")
    frmEndPart.Show
End Sub
Private Sub mnuInvoiceEnd_basil_ret_Click()
    popupvalue5 = UCase("cashbasilret")
    frmEndPart.Show
End Sub
Private Sub mnuInvoiceEnd_Click()
    'searchForm = "credititem"
    popupvalue5 = UCase("invoice")
    frmEndPart.Show
End Sub
Private Sub mnuManufacture_Click()
    Toolbar1.Visible = False
    bookmaster.SSTab1.Tab = 5
    bookmaster.Show
End Sub
Private Sub mnuIssuedBind_Click()
frmIssue.Show
End Sub

Private Sub mnuLedBluePrint_Click()
 change_Pass = "a"
 firm = "blueprint"
 frmLoginl.Show

End Sub
Private Sub mnuLedger_Basil_Click()
frmLedgerView_Basil.Show
End Sub
Private Sub mnuLedger_Click()

change_Pass = "a"
If firm = "chitra" Then
   firm = "chitra"
   frmLoginl.Show
Else
   firm = "chitra"
   frmLoginl.Show
End If


End Sub
Private Sub mnuMonthlyChart_Click()
End Sub
Private Sub mnuLedger_conven_Click()
 frmShowLedger.Show
End Sub
Private Sub mnuMail_Click()
 frmMailSt.Show
End Sub
Private Sub mnuMulti_Click()
 frmMultiMail.Show
End Sub
Private Sub mnumzn_Click()
 frmRepWithMzn.Show 1
End Sub

Private Sub mnuopbalance_Click()
 frmopbalance.Show
End Sub
Private Sub mnuOrdermnm_Click()
 frmOrdermgm.Show
End Sub
Private Sub mnuOrder_Click()
 frmbill.Show
End Sub

Private Sub mnuPacking_Click()
packing_ = "invsp"
mnuMenu_ = "mnuPacking"
frmPackingSlip.Show
End Sub

Private Sub mnuPackingINV_Click()
packing_ = "inv"
mnuMenu_ = "mnuPackingINV"
frmPackingSlip.Show
End Sub
Private Sub mnuPaerRec_Click()

'frmpaper.Show
frmpaper.type1.text = "R"
Rec1 = "R"
frmpaper.Label5.Caption = "Paper Issue"

frmpaper.Show
frmpaper.lblto.Visible = False
frmpaper.cboToGodown.Visible = False
frmpaper.cboFromGodown.Visible = True

frmpaper.challan1.Visible = True
frmpaper.txtChallanNo.Visible = True
frmpaper.chdate.Visible = True
frmpaper.txtChallanDate.Visible = True


End Sub
Private Sub mnuPaper_Click()
  module_ = "Paper"
  permissionUserWise
End Sub
Private Sub mnuPaper_gsm_Click()
  frmGSM.Show
End Sub
Private Sub mnuPaper_Maker_Click()
  papermaker.Show
End Sub
Private Sub mnuPaper_NegativePrint_Click()
  frmOrderNeg.Show
End Sub
Private Sub mnuPaper_size_Click()
 frmPapersize.Show
End Sub
Private Sub mnuPaper_subjectprint_Click()
 frmSubjectPrint.Show
End Sub
Private Sub mnuPaper_TitlePrint_Click()
 frmOrderForTitlePrint.Show
End Sub
Private Sub mnupartyclosing_Click()
 s1 = 2
 frmSelectedParty.Show
End Sub
Private Sub mnupst_Click()
 s1 = 1
 frmSelectedParty.Show
End Sub

Private Sub mnuPaperPrint_Click()
 frmPaperPlan.Show
End Sub

Private Sub mnuPaperSize_Click()
frmPaperLedger.Show
End Sub
Private Sub mnuPartyList_Click()
partylist.Show
End Sub

Private Sub mnuPartyOuts_Click()
frmPaymentSupply.Show
End Sub

Private Sub mnuPartyProf_Click()
frmPartyProfile.Show
End Sub

Private Sub mnupartywise_Click()
  frmPartyWiseAgentwise.Show
End Sub

Private Sub mnuPBookMast_Click()
frmbook.Show
End Sub
Private Sub mnuProductsale_Click()
frmProductWiseSale.Show 1
End Sub
Private Sub mnuPsummary_Click()
frmFinishItemStochSumary.Show
End Sub
Private Sub mnupurchaserpt_Click()
s2 = 3
frmRegsiterList.Show
End Sub
Private Sub mnupurparstmt_Click()
ss1 = 5
frmItemSale.Show
End Sub
Private Sub mnuRawItemStock_Click()
frmMsgReports.Show
End Sub
Private Sub mnuRaw_Click()
'frmSalesOrder.Show
frmIssueReadyMade.Show
End Sub

Private Sub mnuPrintOn_Click()
HeadTbl = "SubjectMaster"
frmMasters.Show 1
End Sub

Private Sub mnuPurchaseOrder_Click()
frmPaperPurchaseOrder.Show
End Sub
Private Sub filedelete(filename As String)
On Error Resume Next
Dim filesystemobject As Object
Set filesystemobject = CreateObject("Scripting.filesystemobject")
filesystemobject.DeleteFile filename, True

End Sub

Private Sub mnuref_Click()

filedelete ("\\192.168.0.140\blueprintSales\*.jpg")
filedelete ("\\192.168.0.140\blueprintSales\*.pdf")

con.Execute "delete from treport"
con.Execute "delete from subledgertrail"
If MsgBox("want to refresh.. ?", vbYesNo) = vbYes Then
   con.Execute "delete from tmpDDet"
   con.Execute "delete from tmpSAdjDet"
   addmaster
End If
End Sub

Private Sub mnuReminder_Click()

frmReminder.Show

End Sub

'Private Sub mnuRawRpt_Click()
'frmWeatorWiseSale.Show
'End Sub
Private Sub mnuReportBookGp_Click()
frmBookGpReport.Show 1
End Sub
Private Sub mnureportbookgroup_Click()
Toolbar1.Visible = False
bookmaster.SSTab1.Tab = 4
bookmaster.Show
End Sub
Private Sub mnuRItem_Click()
  frmItemCreation.Show
End Sub
Private Sub mnuRP_Click()
FinishPurchase.Show
End Sub

Private Sub mnuReports_Click()
frmYearReport.Show
End Sub

Private Sub mnuRepTitlewise_Click()
frmRepWiseTitleWise.Show
End Sub

Private Sub mnuRepWiseSp_Click()
frmSpSummary.Show
End Sub

Private Sub mnuSaleOrder_Click()
mnuMenu_ = "mnuSaleOrder"
If Left(session, 4) >= 2018 Then
frmINVOrder.Show
End If
End Sub

Private Sub mnuSaleRet_Click()
frmReturnStatus.Show
End Sub

Private Sub mnuSch_NetSales_Click()
frmNetBookQty.Show
End Sub

Private Sub mnuschool_Click()
frmUpDateSchool.Show
End Sub

Private Sub mnuSchoolbkSummar_Click()

End Sub

Private Sub mnuSchoolLed_Click()
frmSchoolLedger.Show
End Sub

Private Sub mnuSchoolLedgerSp_Click()
frmSchoolLedgerSP.Show
End Sub

Private Sub mnuSchoolMaster_Click()
FrmSchool.Show
End Sub

Private Sub mnuSchoolWise_Click()
frmSaleRet.Show
End Sub

Private Sub mnusearch_Click()
 SubledgerSearch.Show
End Sub
Private Sub mnuSetColor_Click()
frmColor.Show 1
End Sub
Private Sub mnuSL_invoice_Click()
frmSLedgerView.Show
End Sub

Private Sub mnuSms_Click()
 frmSMS.Show
End Sub

Private Sub mnuSpRet_Click()
frmBookIssueSp_Ret.Show
End Sub
Private Sub mnustock_Click()
module_ = "Stock System"
permissionUserWise
End Sub
Private Sub mnuStockTrans_Click()

Rec1 = "D"
frmpaper.type1.text = "D"
frmpaper.Label5.Caption = "PAPER TRANSFER "
frmpaper.Show
frmpaper.lblto.Visible = True
frmpaper.cboFromGodown.Visible = True

frmpaper.challan1.Visible = True
frmpaper.txtChallanNo.Visible = True
frmpaper.chdate.Visible = True
frmpaper.txtChallanDate.Visible = True


End Sub

Private Sub mnuStockTranstoGo_Click()
    
    IssueBook = "StockTransfar"
    Unload frmBookIssue
    frmBookIssue.Show

End Sub



Private Sub mnuSubGpDiscount_Click()
frmSubGP_discount.Show
End Sub

Private Sub mnuTCS_Click()
frmTCS.Show
End Sub

Private Sub mnuTeacherDet_Click()
frmTeacherDetail.Show
End Sub

Private Sub mnutod_Click()
frmTurnOverDis.Show
End Sub

Private Sub mnuTransport_Click()
frmTransport.Show 1
End Sub

Private Sub mnuUnit_Click()
UnitMaster.Show
End Sub

Private Sub mnuTTrans_Click()
frmTotalTrans.Show
End Sub

Private Sub mnuUpdate_Click()
frmUpDataOnNet.Show
End Sub
Private Sub mnuuserright_Click()
'Toolbar1.Visible = False
frmPermission.Show
'userright.Show
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)


Select Case Button.Key

Case "inv"
      module_ = "Invoicing"
Case "basil"
     module_ = "Basil"
Case "paper"
     module_ = "Paper"
Case "st"
     module_ = "Stock System"
Case "con"
     module_ = "Canvassing"
End Select


BackColorFrom Me
permissionUserWise
Me.Caption = "Publication Software System : " & module_

End Sub
Private Sub mnuwastqty_Click()
 frmPaperWastage.Show
End Sub
