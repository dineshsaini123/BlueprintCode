VERSION 5.00
Begin VB.Form userright 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Rights"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Caption         =   "Button Right"
      Height          =   855
      Left            =   90
      TabIndex        =   58
      Top             =   4620
      Width           =   6495
      Begin VB.CheckBox btDelete 
         Caption         =   "Delete"
         Height          =   465
         Left            =   5250
         TabIndex        =   41
         Top             =   210
         Width           =   825
      End
      Begin VB.CheckBox btsave 
         Caption         =   "Save"
         Height          =   465
         Left            =   3210
         TabIndex        =   40
         Top             =   210
         Width           =   825
      End
      Begin VB.CheckBox Btadd 
         Caption         =   "Add"
         Height          =   195
         Left            =   150
         TabIndex        =   38
         Top             =   330
         Width           =   765
      End
      Begin VB.CheckBox btedit 
         Caption         =   "Edit"
         Height          =   195
         Left            =   1590
         TabIndex        =   39
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1020
      Left            =   45
      TabIndex        =   55
      Top             =   30
      Width           =   4200
      Begin VB.ComboBox CmbUname 
         Height          =   315
         Left            =   1530
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   210
         Width           =   2550
      End
      Begin VB.TextBox txtpass 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1530
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   585
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         Height          =   195
         Left            =   315
         TabIndex        =   57
         Top             =   285
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User Password"
         Height          =   195
         Left            =   315
         TabIndex        =   56
         Top             =   645
         Width           =   1065
      End
   End
   Begin VB.CheckBox chkselectall 
      Caption         =   "Select All"
      Height          =   285
      Index           =   0
      Left            =   300
      TabIndex        =   3
      Top             =   1080
      Width           =   1845
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   5340
      TabIndex        =   52
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   5340
      TabIndex        =   51
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   " Master Rights"
      Height          =   2115
      Left            =   60
      TabIndex        =   50
      Top             =   1380
      Width           =   4905
      Begin VB.CheckBox mnubookgroupmaster 
         Caption         =   "Book Group Master"
         Height          =   315
         Left            =   3030
         TabIndex        =   9
         Top             =   150
         Width           =   1845
      End
      Begin VB.CheckBox mnureportbookgroupmaster 
         Caption         =   "Report Book Group"
         Height          =   315
         Left            =   3030
         TabIndex        =   13
         Top             =   1410
         Width           =   1845
      End
      Begin VB.CheckBox mnudistrictmaster 
         Caption         =   "District Master"
         Height          =   195
         Left            =   3030
         TabIndex        =   12
         Top             =   1080
         Width           =   1635
      End
      Begin VB.CheckBox mnuagentmaster 
         Caption         =   "Agent Master"
         Height          =   225
         Left            =   3030
         TabIndex        =   11
         Top             =   780
         Width           =   1755
      End
      Begin VB.CheckBox mnudiscountcategorymaster 
         Caption         =   "Discount Category"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1770
         Width           =   2115
      End
      Begin VB.CheckBox mnuGenralLedgermaster 
         Caption         =   "General Ledger"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1545
      End
      Begin VB.CheckBox mnuSubLedgermaster 
         Caption         =   "Sub Ledger "
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   630
         Width           =   1305
      End
      Begin VB.CheckBox mnuinvendmaster 
         Caption         =   "Invoice End Part "
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   915
         Width           =   1605
      End
      Begin VB.CheckBox mnucashcountersalesmaster 
         Caption         =   "Counter Sale End Part"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1470
         Width           =   2025
      End
      Begin VB.CheckBox mnucreendmaster 
         Caption         =   "Credit Note End Part"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   1875
      End
      Begin VB.CheckBox mnubookmaster 
         Caption         =   "Book Master"
         Height          =   195
         Left            =   3030
         TabIndex        =   10
         Top             =   510
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   585
      Left            =   1245
      ScaleHeight     =   525
      ScaleWidth      =   7665
      TabIndex        =   49
      Top             =   5640
      Width           =   7725
      Begin VB.CommandButton SSCommand3 
         Caption         =   "&Delete"
         Height          =   465
         Left            =   2640
         TabIndex        =   44
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton Help 
         Caption         =   "&Help"
         Height          =   465
         Left            =   210
         TabIndex        =   42
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton REPORTCD 
         Caption         =   "&Print"
         Height          =   465
         Left            =   5085
         TabIndex        =   46
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton cmdabandon 
         Caption         =   "&Abandon"
         Height          =   465
         Left            =   3870
         TabIndex        =   45
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton savecd 
         Caption         =   "&Save"
         Height          =   465
         Left            =   1425
         TabIndex        =   43
         Top             =   30
         Width           =   1155
      End
      Begin VB.CommandButton TESTENTRYCD 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   465
         Left            =   6300
         TabIndex        =   47
         Top             =   30
         Width           =   1155
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Transaction  Rights"
      Height          =   1095
      Left            =   90
      TabIndex        =   48
      Top             =   3510
      Width           =   6555
      Begin VB.CheckBox mnudebitnote 
         Caption         =   "Debit Note"
         Height          =   195
         Left            =   4620
         TabIndex        =   32
         Top             =   660
         Width           =   1215
      End
      Begin VB.CheckBox mnucashcountersales 
         Caption         =   "Cash Counter Sales"
         Height          =   195
         Left            =   2070
         TabIndex        =   30
         Top             =   630
         Width           =   1965
      End
      Begin VB.CheckBox mnucreditnote 
         Caption         =   "Credit Note"
         Height          =   195
         Left            =   4620
         TabIndex        =   31
         Top             =   360
         Width           =   1185
      End
      Begin VB.CheckBox mnucreditnoteitem 
         Caption         =   "Credit Note(Items)"
         Height          =   195
         Left            =   2070
         TabIndex        =   29
         Top             =   330
         Width           =   1725
      End
      Begin VB.CheckBox mnusalesinvoice 
         Caption         =   "Sales Invoice"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   570
         Width           =   2115
      End
      Begin VB.CheckBox mnuvoucherentry 
         Caption         =   "Voucher Entry"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   270
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tools Menu"
      Height          =   1095
      Left            =   6750
      TabIndex        =   33
      Top             =   3510
      Width           =   3945
      Begin VB.CheckBox createuser 
         Caption         =   "Create User"
         Height          =   195
         Left            =   180
         TabIndex        =   35
         Top             =   660
         Width           =   1485
      End
      Begin VB.CheckBox mnutoolsetup 
         Caption         =   "Owner Setup"
         Height          =   195
         Left            =   180
         TabIndex        =   34
         Top             =   330
         Width           =   1515
      End
      Begin VB.CheckBox directentrybutton 
         Height          =   465
         Left            =   1740
         TabIndex        =   37
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox mnutool 
         Height          =   465
         Left            =   1710
         TabIndex        =   36
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Reports Rights"
      Height          =   2115
      Left            =   5010
      TabIndex        =   2
      Top             =   1380
      Width           =   5685
      Begin VB.CheckBox rptbankadvice 
         Caption         =   "Bank Advice"
         Height          =   195
         Left            =   3150
         TabIndex        =   26
         Top             =   1470
         Width           =   2445
      End
      Begin VB.CheckBox rptbankadvicereconcilation 
         Caption         =   "Bank Advice Rconcilation"
         Height          =   195
         Left            =   3150
         TabIndex        =   25
         Top             =   1230
         Width           =   2445
      End
      Begin VB.CheckBox rptbookgroupwisesales 
         Caption         =   "Book Group wise sales"
         Height          =   195
         Left            =   3150
         TabIndex        =   24
         Top             =   990
         Width           =   2445
      End
      Begin VB.CheckBox rptgenralledgerac 
         Caption         =   "General Ledger A/c"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   510
         Width           =   2655
      End
      Begin VB.CheckBox rptcashbankbook 
         Caption         =   "Cash/Bank book"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   270
         Width           =   2385
      End
      Begin VB.CheckBox rptdistictwisesales 
         Caption         =   "District wise sales"
         Height          =   195
         Left            =   3150
         TabIndex        =   22
         Top             =   510
         Width           =   2505
      End
      Begin VB.CheckBox rptsubledgertrialbalance 
         Caption         =   "Sub Ledger Trial balance"
         Height          =   195
         Left            =   3150
         TabIndex        =   21
         Top             =   270
         Width           =   2355
      End
      Begin VB.CheckBox rptdistictwisesalesreturn 
         Caption         =   "District wise sales"
         Height          =   195
         Left            =   3150
         TabIndex        =   23
         Top             =   750
         Width           =   2445
      End
      Begin VB.CheckBox rptsubledgerac 
         Caption         =   "Sub Ledger A/c"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   780
         Width           =   2535
      End
      Begin VB.CheckBox rptalphawisesubledgerac 
         Caption         =   "Alphabat wise sub ledger"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   1050
         Width           =   2775
      End
      Begin VB.CheckBox rptgenledgertrialbalance 
         Caption         =   "Gen Ledger Trial balance"
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   1320
         Width           =   2265
      End
      Begin VB.CheckBox rptgenledgeropentrial 
         Caption         =   "Gen Ledger Opening Trial"
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   1620
         Width           =   2745
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Last login"
      Height          =   195
      Left            =   4410
      TabIndex        =   54
      Top             =   180
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Last Logout"
      Height          =   225
      Left            =   4410
      TabIndex        =   53
      Top             =   510
      Visible         =   0   'False
      Width           =   840
   End
End
Attribute VB_Name = "userright"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkselectall_Click(Index As Integer)
  
  If chkselectall(0).value = 1 Then
    mnuGenralLedgermaster.value = 1
    mnuSubLedgermaster.value = 1
    mnuinvendmaster.value = 1
    mnucreendmaster.value = 1
    mnucashcountersalesmaster.value = 1
    mnudiscountcategorymaster.value = 1
    mnubookmaster.value = 1
    mnubookgroupmaster.value = 1
    mnureportbookgroupmaster.value = 1
    mnudistrictmaster.value = 1
    mnuagentmaster.value = 1
    rptcashbankbook.value = 1
    rptgenralledgerac.value = 1
    rptsubledgerac.value = 1
    rptalphawisesubledgerac.value = 1
    rptgenledgertrialbalance.value = 1
    rptgenledgeropentrial.value = 1
    rptsubledgertrialbalance.value = 1
    rptdistictwisesales.value = 1
    rptdistictwisesalesreturn.value = 1
    rptbookgroupwisesales.value = 1
    rptbankadvicereconcilation.value = 1
    rptbankadvice.value = 1
    mnuvoucherentry.value = 1
    mnusalesinvoice.value = 1
    mnucreditnoteitem.value = 1
    mnucashcountersales.value = 1
    mnucreditnote.value = 1
    mnudebitnote.value = 1
    createuser.value = 1
    mnutoolsetup.value = 1
    Btadd.value = 1
    btedit.value = 1
    btsave.value = 1
    btDelete.value = 1
 Else
        mnuGenralLedgermaster.value = 0
        mnuSubLedgermaster.value = 0
        mnuinvendmaster.value = 0
        mnucreendmaster.value = 0
        mnucashcountersalesmaster.value = 0
        mnudiscountcategorymaster.value = 0
        mnubookmaster.value = 0
        mnubookgroupmaster.value = 0
        mnureportbookgroupmaster.value = 0
        mnudistrictmaster.value = 0
        mnuagentmaster.value = 0
        rptcashbankbook.value = 0
        rptgenralledgerac.value = 0
        rptsubledgerac.value = 0
        rptalphawisesubledgerac.value = 0
        rptgenledgertrialbalance.value = 0
        rptgenledgeropentrial.value = 0
        rptsubledgertrialbalance.value = 0
        rptdistictwisesales.value = 0
        rptdistictwisesalesreturn.value = 0
        rptbookgroupwisesales.value = 0
        rptbankadvicereconcilation.value = 0
        rptbankadvice.value = 0
   mnuvoucherentry.value = 0
   mnusalesinvoice.value = 0
   mnucreditnoteitem.value = 0
   mnucashcountersales.value = 0
   mnucreditnote.value = 0
   mnudebitnote.value = 0
   createuser.value = 0
   mnutoolsetup.value = 0
   Btadd.value = 0
   btedit.value = 0
   btsave.value = 0
   btDelete.value = 0
   
 End If
End Sub

Private Sub CmbUname_Click()
CmbUname_KeyPress 13
End Sub

Private Sub CmbUname_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   'SendKeys "{Tab}"
   If CmbUname.Text = "" Then
     MsgBox "ENTER USER NAME"
     CmbUname.SetFocus
     Exit Sub
  End If
  Dim rsuser As New ADODB.Recordset
   rsuser.Open "Select * from login where  " & stringyear & " and uname  = '" & CmbUname.Text & "'", con, adOpenDynamic, adLockOptimistic
   If rsuser.RecordCount > 0 Then
        txtpass.Text = rsuser!Password
        rsuser!uname = CmbUname.Text
   rsuser!Password = txtpass.Text
   
   
   mnuGenralLedgermaster.value = IIf(rsuser!mnuGenralLedgermaster, 1, 0)
   mnuSubLedgermaster.value = IIf(rsuser!mnuSubLedgermaster, 1, 0)
   mnuinvendmaster.value = IIf(rsuser!mnuinvendmaster, 1, 0)
   mnucreendmaster.value = IIf(rsuser!mnucreendmaster, 1, 0)
   mnucashcountersalesmaster.value = IIf(rsuser!mnucashcountersalesmaster, 1, 0)
   mnudiscountcategorymaster.value = IIf(rsuser!mnudiscountcategorymaster, 1, 0)
   mnubookmaster.value = IIf(rsuser!mnubookmaster, 1, 0)
   mnubookgroupmaster.value = IIf(rsuser!mnubookgroupmaster, 1, 0)
   mnureportbookgroupmaster.value = IIf(rsuser!mnureportbookgroupmaster, 1, 0)
   mnudistrictmaster.value = IIf(rsuser!mnudistrictmaster, 1, 0)
   mnuagentmaster.value = IIf(rsuser!mnuagentmaster, 1, 0)
   
   rptcashbankbook.value = IIf(rsuser!rptcashbankbook, 1, 0)
   rptgenralledgerac.value = IIf(rsuser!rptgenralledgerac, 1, 0)
   rptsubledgerac.value = IIf(rsuser!rptsubledgerac, 1, 0)
   rptalphawisesubledgerac.value = IIf(rsuser!rptalphawisesubledgerac, 1, 0)
   rptgenledgertrialbalance.value = IIf(rsuser!rptgenledgertrialbalance, 1, 0)
   rptgenledgeropentrial.value = IIf(rsuser!rptgenledgeropentrial, 1, 0)
   rptsubledgertrialbalance.value = IIf(rsuser!rptsubledgertrialbalance, 1, 0)
   rptdistictwisesales.value = IIf(rsuser!rptdistictwisesales, 1, 0)
   rptdistictwisesalesreturn.value = IIf(rsuser!rptdistictwisesalesreturn, 1, 0)
   rptbookgroupwisesales.value = IIf(rsuser!rptbookgroupwisesales, 1, 0)
   rptbankadvicereconcilation.value = IIf(rsuser!rptbankadvicereconcilation, 1, 0)
   rptbankadvice.value = IIf(rsuser!rptbankadvice, 1, 0)
   
   mnuvoucherentry.value = IIf(rsuser!mnuvoucherentry, 1, 0)
   mnusalesinvoice.value = IIf(rsuser!mnusalesinvoice, 1, 0)
   mnucreditnoteitem.value = IIf(rsuser!mnucreditnoteitem, 1, 0)
   mnucashcountersales.value = IIf(rsuser!mnucashcountersales, 1, 0)
   mnucreditnote.value = IIf(rsuser!mnucreditnote, 1, 0)
   mnudebitnote.value = IIf(rsuser!mnudebitnote, 1, 0)
   createuser.value = IIf(rsuser!createuser, 1, 0)
   mnutoolsetup.value = IIf(rsuser!mnutoolsetup, 1, 0)
   Btadd.value = IIf(rsuser!badd, 1, 0)
   btedit.value = IIf(rsuser!bedit, 1, 0)
   btsave.value = IIf(rsuser!bsave, 1, 0)
   btDelete.value = IIf(rsuser!bDelete, 1, 0)
  End If
End If

End Sub

Sub cmdabandon_Click()
CmbUname.Text = ""
txtpass.Text = ""
chkselectall(0).value = 0
chkselectall_Click (0)
Me.CmbUname.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{Tab}"
End If

End Sub

Private Sub Form_Load()
     Dim rs2 As New ADODB.Recordset
     rs2.Open "SELECT Uname from login where  " & stringyear, con, adOpenStatic, adLockReadOnly
     If rs2.RecordCount > 0 Then
        rs2.MoveFirst
        While Not rs2.EOF
             CmbUname.AddItem rs2!uname
             rs2.MoveNext
        Wend
     End If
End Sub

Private Sub Form_Unload(cancel As Integer)
'''''MainMenu.Toolbar1.Visible = True
End Sub

Private Sub savecd_Click()
Dim rs1 As New ADODB.Recordset
Dim rsuser As New ADODB.Recordset
  If CmbUname.Text = "" Then
     MsgBox "ENTER USER NAME", , "Massage....."
     CmbUname.SetFocus
     Exit Sub
  End If
  If rsuser.State = 1 Then rsuser.close
  rsuser.Open "Select * from login where  " & stringyear & " and uname = '" & CmbUname.Text & "'", con, adOpenDynamic, adLockOptimistic
  If rsuser.RecordCount <= 0 Then
     rsuser.AddNew
     CmbUname.AddItem CmbUname.Text
  End If
   rsuser!uname = CmbUname.Text
   rsuser!Password = txtpass.Text
   
   rsuser!mnuGenralLedgermaster = mnuGenralLedgermaster.value
   rsuser!mnuSubLedgermaster = mnuSubLedgermaster.value
   rsuser!mnuinvendmaster = mnuinvendmaster.value
   rsuser!mnucreendmaster = mnucreendmaster.value
   rsuser!mnucashcountersalesmaster = mnucashcountersalesmaster.value
   rsuser!mnudiscountcategorymaster = mnudiscountcategorymaster.value
   rsuser!mnubookmaster = mnubookmaster.value
   rsuser!mnubookgroupmaster = mnubookgroupmaster.value
   rsuser!mnureportbookgroupmaster = mnureportbookgroupmaster.value
   rsuser!mnudistrictmaster = mnudistrictmaster.value
   rsuser!mnuagentmaster = mnuagentmaster.value
   
   rsuser!rptcashbankbook = rptcashbankbook.value
   rsuser!rptgenralledgerac = rptgenralledgerac.value
   rsuser!rptsubledgerac = rptsubledgerac.value
   rsuser!rptalphawisesubledgerac = rptalphawisesubledgerac.value
   rsuser!rptgenledgertrialbalance = rptgenledgertrialbalance.value
   rsuser!rptgenledgeropentrial = rptgenledgeropentrial.value
   rsuser!rptsubledgertrialbalance = rptsubledgertrialbalance.value
   rsuser!rptdistictwisesales = rptdistictwisesales.value
   rsuser!rptdistictwisesalesreturn = rptdistictwisesalesreturn.value
   rsuser!rptbookgroupwisesales = rptbookgroupwisesales.value
   rsuser!rptbankadvicereconcilation = rptbankadvicereconcilation.value
   rsuser!rptbankadvice = rptbankadvice.value
   
   rsuser!mnuvoucherentry = mnuvoucherentry.value
   rsuser!mnusalesinvoice = mnusalesinvoice.value
   rsuser!mnucreditnoteitem = mnucreditnoteitem.value
   rsuser!mnucashcountersales = mnucashcountersales.value
   rsuser!mnucreditnote = mnucreditnote.value
   rsuser!mnudebitnote = mnudebitnote.value
   
   rsuser!createuser = createuser.value
   rsuser!mnutoolsetup = mnutoolsetup.value
   rsuser!badd = Btadd.value
   rsuser!bedit = btedit.value
   rsuser!bsave = btsave.value
   rsuser!bDelete = btDelete.value
   
   
   rsuser.update
   MsgBox "Saved User Name/Password & rights ", vbInformation
   cmdabandon_Click
   
   

End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub SSCommand3_Click()
If MsgBox("Are you sure..", vbOKCancel) = vbOK Then
 If CmbUname.Text <> "" Then
    con.Execute "DELETE from login where  " & stringyear & " and uname  = '" & CmbUname.Text & "'"
    cmdabandon_Click
    Dim rs2 As New ADODB.Recordset
    rs2.Open "SELECT UName from login where  " & stringyear & "", con, adOpenStatic, adLockReadOnly
    CmbUname.Clear
    If rs2.RecordCount > 0 Then
      rs2.MoveFirst
      While Not rs2.EOF
             CmbUname.AddItem rs2!uname
             rs2.MoveNext
      Wend
   End If
 Else
   CmbUname.SetFocus
 End If
End If
End Sub

Private Sub TESTENTRYCD_Click()
Unload Me
End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

   'SendKeys "{Tab}"

   Dim rsuser As New ADODB.Recordset
   'rsuser.Open "Select * from userright where username  = '" & CmbUname.Text & "' and password = '" & txtpass.Text & "'", con, adOpenDynamic, adLockOptimistic
   rsuser.Open "Select * from login where  " & stringyear & " and uname  = '" & CmbUname.Text & "'", con, adOpenDynamic, adLockOptimistic
   If rsuser.RecordCount > 0 Then
   mnuGenralLedgermaster.value = IIf(rsuser!mnuGenralLedgermaster, 1, 0)
   mnuSubLedgermaster.value = IIf(rsuser!mnuSubLedgermaster, 1, 0)
   mnuinvendmaster.value = IIf(rsuser!mnuinvendmaster, 1, 0)
   mnucreendmaster.value = IIf(rsuser!mnucreendmaster, 1, 0)
   mnucashcountersalesmaster.value = IIf(rsuser!mnucashcountersalesmaster, 1, 0)
   mnudiscountcategorymaste.value = IIf(rsuser!mnudiscountcategorymaste, 1, 0)
   mnubookmaster.value = IIf(rsuser!mnubookmaster, 1, 0)
   mnubookgroupmaster.value = IIf(rsuser!mnubookgroupmaster, 1, 0)
   mnureportbookgroupmaster.value = IIf(rsuser!mnureportbookgroupmaster, 1, 0)
   mnudistrictmaster.value = IIf(rsuser!mnudistrictmaster, 1, 0)
   mnuagentmaster.value = IIf(rsuser!mnuagentmaster, 1, 0)
   
   rptcashbankbook.value = IIf(rsuser!rptcashbankbook, 1, 0)
   rptgenralledgerac.value = IIf(rsuser!rptgenralledgerac, 1, 0)
   rptsubledgerac.value = IIf(rsuser!rptsubledgerac, 1, 0)
   rptalphawisesubledgerac.value = IIf(rsuser!rptalphawisesubledgerac, 1, 0)
   rptgenledgertrialbalance.value = IIf(rsuser!rptgenledgertrialbalance, 1, 0)
   rptgenledgeropentrial.value = IIf(rsuser!rptgenledgeropentrial, 1, 0)
   rptsubledgertrialbalance.value = IIf(rsuser!rptsubledgertrialbalance, 1, 0)
   rptdistictwisesales.value = IIf(rsuser!rptdistictwisesales, 1, 0)
   rptdistictwisesalesreturn.value = IIf(rsuser!rptdistictwisesalesreturn, 1, 0)
   rptbookgroupwisesales.value = IIf(rsuser!rptbookgroupwisesales, 1, 0)
   rptsubtest.value = IIf(rsuser!rptbankadvicereconcilation, 1, 0)
   rptbankadvice.value = IIf(rsuser!rptbankadvice, 1, 0)
   mnuvoucherentry.value = IIf(rsuser!mnuvoucherentry, 1, 0)
   mnusalesinvoice.value = IIf(rsuser!mnusalesinvoice, 1, 0)
   mnucreditnoteitem.value = IIf(rsuser!mnucreditnoteitem, 1, 0)
   mnucashcountersales.value = IIf(rsuser!mnucashcountersales, 1, 0)
   mnucreditnote.value = IIf(rsuser!mnucreditnote, 1, 0)
   mnudebitnote.value = IIf(rsuser!mnudebitnote, 1, 0)
   rsuser!createuser = IIf(createuser.value, 1, 0)
   rsuser!mnutoolsetup = IIf(mnutoolsetup.value, 1, 0)
   End If
End If
End Sub
