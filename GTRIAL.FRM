VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Gentrial 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5805
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Gtrial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5790
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton export 
      Height          =   345
      Left            =   1440
      Picture         =   "Gtrial.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4710
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Gtrial.frx":045D
      Left            =   2370
      List            =   "Gtrial.frx":0470
      TabIndex        =   12
      Text            =   "100 %"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton print1 
      Height          =   345
      Left            =   1920
      Picture         =   "Gtrial.frx":0496
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4710
      Visible         =   0   'False
      Width           =   345
   End
   Begin RichTextLib.RichTextBox r1 
      Height          =   4095
      Left            =   840
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   7223
      _Version        =   393217
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      MaxLength       =   99999999
      RightMargin     =   20000
      TextRTF         =   $"Gtrial.frx":0608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   405
      Left            =   4830
      TabIndex        =   5
      Top             =   2640
      Width           =   1545
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      Height          =   405
      Left            =   3060
      TabIndex        =   4
      Top             =   2640
      Width           =   1545
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   345
      Left            =   3060
      TabIndex        =   2
      Top             =   1950
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      _Version        =   393216
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   3030
      TabIndex        =   0
      Top             =   690
      Width           =   3885
   End
   Begin MSMask.MaskEdBox date2 
      Height          =   345
      Left            =   5670
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      _Version        =   393216
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox alpha 
      Height          =   345
      Left            =   3060
      TabIndex        =   1
      Top             =   1290
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   1
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Alphabat"
      Height          =   195
      Left            =   990
      TabIndex        =   9
      Top             =   1350
      Width           =   630
   End
   Begin VB.Label Label3 
      Caption         =   " - To - "
      Height          =   315
      Left            =   4770
      TabIndex        =   8
      Top             =   1950
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "From The Date"
      Height          =   195
      Left            =   930
      TabIndex        =   7
      Top             =   1950
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Gen. Ledger Desc."
      Height          =   195
      Left            =   960
      TabIndex        =   6
      Top             =   720
      Width           =   1350
   End
End
Attribute VB_Name = "Gentrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As Connection
Dim rs As Recordset

Private Sub alpha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub COMBOGENLEDGER_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub COMBOGENLEDGER_LostFocus()
    If Trim(COMBOGENLEDGER.Text) <> "" Then
        rs.Open "select * from gledger", con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not rs.BOF Then
            rs.Find "gledger='" + Trim(COMBOGENLEDGER.Text) + "'"
            If rs.EOF Then
                COMBOGENLEDGER.SetFocus
            End If
        Else
            COMBOGENLEDGER.SetFocus
        End If
        rs.Close
    End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Commandreturn_Click()
MainMenu.Toolbar1.Visible = True
    Unload Me
End Sub
Private Sub Commandshow_Click()
'If Trim(COMBOGENLEDGER.Text) <> "" Then
    If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
        MsgBox "invalid date"
        Exit Sub
    End If
    Dim Trptrs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim rs3 As New ADODB.Recordset
    Dim rs3cash As New ADODB.Recordset
    
    Dim rs4 As New ADODB.Recordset
    Dim rs4c As New ADODB.Recordset
    
    Dim rs5 As New ADODB.Recordset
    Dim rs5C As New ADODB.Recordset
    Dim rs6 As New ADODB.Recordset
    Dim rs6C As New ADODB.Recordset
    Dim rs7 As New ADODB.Recordset
    Dim rs7C As New ADODB.Recordset
    Dim rs8 As New ADODB.Recordset
    
    
    Dim Ors1 As New ADODB.Recordset
    Dim Ors2 As New ADODB.Recordset
    Dim Ors3 As New ADODB.Recordset
    Dim Ors3cash As New ADODB.Recordset
    Dim Ors4 As New ADODB.Recordset
    Dim Ors41 As New ADODB.Recordset
    Dim Ors5 As New ADODB.Recordset
    Dim Ors51 As New ADODB.Recordset
    Dim Ors6 As New ADODB.Recordset
    Dim Ors61 As New ADODB.Recordset
    Dim Ors7 As New ADODB.Recordset
    Dim Ors71 As New ADODB.Recordset
    Dim Ors8 As New ADODB.Recordset
    
 
    Dim RSopBalDR As New ADODB.Recordset
    Dim RSopBalCR As New ADODB.Recordset
    Dim RsDRCR As New ADODB.Recordset
     
    
    
    Dim RsVDr As New ADODB.Recordset
    Dim RsVCr As New ADODB.Recordset
    
    
    
    Dim CON1 As ADODB.Connection
    Dim BALANCE As Double
    Dim OPBALANCE As Double
    Dim SDamount As Double
    Dim SCamount As Double
    Set CON1 = New ADODB.Connection
    Set rs1 = New ADODB.Recordset
    Dim viewsubledger As Boolean
    viewsubledger = False
    BALANCE = 0
    OPBALANCE = 0
    
    CON1.CursorLocation = adUseClient
    CON1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + Trim(VB.App.Path) + "\" + Trim(main.directory) + "\tchitra.mdb"
    CON1.Open
    If rs.State = 1 Then
        rs.Close
    End If
    
    con.Execute "Delete from TemprptTrialBalance"
    con.BeginTrans
        
            
     If Trim(Alpha.Text) <> "" Then
                con.Execute "INSERT INTO TemprptTrialBalance (Gledger, OpeningBalance ) SELECT Gledger.gledger,  Gledger.YEAROPENING from Gledger where  Gledger.gledger like '" + Trim(Alpha.Text) + "%'"
  
                
                'RSopBal.Open "Select GenLedger, (SELECT  sum(amount)as DAmount  FROM VOUCHERS  where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate<cdate('" + Trim(date1.Text) + "') and DebitorCredit='D'  GROUP BY GenLedger, SubLedger ) as SDr, (SELECT  sum(amount)as CAmount  FROM VOUCHERS  where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate<cdate('" + Trim(date1.Text) + "') and DebitorCredit='C'   and  Genledger like '" + Trim(alpha.Text) + "%'  GROUP BY GenLedger ) as SCr from vouchers  where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate<cdate('" + Trim(date1.Text) + "') and Genledger like '" + Trim(alpha.Text) + "%'    GROUP BY GenLedger  ", con, adOpenKeyset, adLockReadOnly, adCmdText
                
                RSopBalDR.Open "SELECT GenLedger,  sum(amount)as DAmount  FROM VOUCHERS  where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate<cdate('" + Trim(date1.Text) + "') and DebitorCredit='D'  GROUP BY GenLedger  ", con, adOpenKeyset, adLockReadOnly, adCmdText
                RSopBalCR.Open "SELECT GenLedger,  sum(amount) as CAmount  FROM VOUCHERS  where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate<cdate('" + Trim(date1.Text) + "') and DebitorCredit='C' GROUP BY GenLedger   ", con, adOpenKeyset, adLockReadOnly, adCmdText
        
                
                'RsV.Open "select GenLedger,(SELECT  sum(amount) as DAmount  FROM VOUCHERS  where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate >= cdate('" + Trim(date1.Text) + "') and VoucherDate <= cdate('" + Trim(date2.Text) + "') and DebitorCredit='D' and Genledger like '" + Trim(alpha.Text) + "%' GROUP BY GenLedger ) as SDr , ( SELECT  sum(amount) as CAmount  FROM VOUCHERS  where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate >= cdate('" + Trim(date1.Text) + "') and VoucherDate <= cdate('" + Trim(date2.Text) + "') and DebitorCredit='C'and Genledger like '" + Trim(alpha.Text) + "%'  GROUP BY GenLedger  ) as SCr   from  vouchers  where  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate >= cdate('" + Trim(date1.Text) + "') and VoucherDate <= cdate('" + Trim(date2.Text) + "')  and Genledger like '" + Trim(alpha.Text) + "%' GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
                
                RsVDr.Open "SELECT GenLedger, sum(amount) as DAmount  FROM VOUCHERS  where  VoucherDate >= cdate('" + Trim(date1.Text) + "') and VoucherDate <= cdate('" + Trim(date2.Text) + "') and DebitorCredit='D'   and genledger like '" + Trim(Alpha.Text) + "%' GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
                RsVCr.Open "SELECT GenLedger, sum(amount) as CAmount  FROM VOUCHERS  where  VoucherDate >= cdate('" + Trim(date1.Text) + "') and VoucherDate <= cdate('" + Trim(date2.Text) + "') and DebitorCredit='C' and genledger like '" + Trim(Alpha.Text) + "%' GROUP BY GenLedger  ", con, adOpenKeyset, adLockReadOnly, adCmdText
       
        
                Ors1.Open "select GenLedger,  sum(netamount) as SAmount from invoicea where INVOICEDATE < cdate('" + Trim(date1.Text) + "')  and  genledger like '" + Trim(Alpha.Text) + "%' GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
                Ors2.Open "select GenLedger,  sum(netamount) as  SAmount from CREDITa where INVOICEDATE < cdate('" + Trim(date1.Text) + "') and genledger like '" + Trim(Alpha.Text) + "%'  GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
                Ors3.Open "select GenLedger, sum(NETamount) as SAmount  from casha where  INVOICEDATE < cdate('" + Trim(date1.Text) + "')  and genledger like '" + Trim(Alpha.Text) + "%'   GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
                Ors3cash.Open "select GenLedger, sum(baa) as SAmount  from casha where INVOICEDATE < cdate('" + Trim(date1.Text) + "')  and genledger like '" + Trim(Alpha.Text) + "%'   GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
   
                Ors4.Open "select pgld, sum(na) as SAmount  from Cnf1a where  cnd<cdate('" + Trim(date1.Text) + "')   and  dc = 'D'  and pgld like '" + Trim(Alpha.Text) + "%'  GROUP BY pgld   ", con, adOpenKeyset, adLockReadOnly, adCmdText
                Ors41.Open "select pgld, sum(na) as SAmount  from Cnf1a where cnd<cdate('" + Trim(date1.Text) + "') and  dc = 'C' and pgld like '" + Trim(Alpha.Text) + "%'  GROUP BY pgld", con, adOpenKeyset, adLockReadOnly, adCmdText
                
                
                Ors5.Open "select pgld , sum(na) as SAmount  from dnfa where  dnd < cdate('" + Trim(date1.Text) + "')  and   dc = 'D' and pgld like '" + Trim(Alpha.Text) + "%' GROUP BY pgld", con, adOpenKeyset, adLockReadOnly, adCmdText
                Ors51.Open "select pgld , sum(na) as SAmount  from dnfa where dnd < cdate('" + Trim(date1.Text) + "')  and   dc = 'C' and pgld like '" + Trim(Alpha.Text) + "%' GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
                
                Ors6.Open "Select gld,sum(a) as SAmount from  CNF1B where  cnd < cdate('" + Trim(date1.Text) + "') and    dc = 'D' and gld like '" + Trim(Alpha.Text) + "%'  GROUP BY gld", con, adOpenKeyset, adLockReadOnly, adCmdText
                Ors61.Open "select gld, sum(a) as SAmount from Cnf1B where cnd < cdate('" + Trim(date1.Text) + "') and    dc = 'C' and gld like '" + Trim(Alpha.Text) + "%'  GROUP BY gld", con, adOpenKeyset, adLockReadOnly, adCmdText
                
               Ors7.Open "select gld, sum(a) as SAmount from dnfB where   dnd < cdate('" + Trim(date1.Text) + "')  and dc='D'  and gld like '" + Trim(Alpha.Text) + "%' GROUP BY gld", con, adOpenKeyset, adLockReadOnly, adCmdText
               Ors71.Open "select gld, sum(a) as SAmount from dnfB where  dnd < cdate('" + Trim(date1.Text) + "') and dc='C' and gld like '" + Trim(Alpha.Text) + "%'  GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
          
               ' opening balance end
        
               'start dr or cr amount for trialbalance
               'invoice
               rs1.Open "select GenLedger, sum(netamount) as SAmount from invoicea where  INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') and genledger like '" + Trim(Alpha.Text) + "%' GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               'credit Item
               rs2.Open "select GenLedger,  sum(netamount) as  SAmount from CREDITa where  INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') and genledger like '" + Trim(Alpha.Text) + "%' GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
              ' cash memo
               rs3.Open "select GenLedger, sum(NETamount) as SAmount  from casha where  INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') and genledger like '" + Trim(Alpha.Text) + "%'   GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
               rs3cash.Open "select GenLedger,   sum(Baa) as SAmount  from casha where INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') and genledger like '" + Trim(Alpha.Text) + "%'    GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               'cnf1a
               rs4.Open "select pgld, sum(na) as SAmount  from Cnf1a where  cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "') and dc ='D'  and pgld like '" + Trim(Alpha.Text) + "%' GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
               rs4c.Open "select pgld, sum(na) as SAmount  from Cnf1a where  cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "') and dc ='C' and pgld like '" + Trim(Alpha.Text) + "%'  GROUP BY pgld", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               'dnfa
               rs5.Open "select pgld, sum(na) as SAmount from dnfa where  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "') and dc ='D'  and pgld like '" + Trim(Alpha.Text) + "%' GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
               rs5C.Open "select pgld, sum(na) as SAmount from dnfa where dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "') and  dc ='C' and pgld like '" + Trim(Alpha.Text) + "%'  GROUP BY pgld", con, adOpenKeyset, adLockReadOnly, adCmdText
               'cnf1b
               rs6.Open "select gld, sum(a) as SAmount from Cnf1B where   cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "') and dc ='D'  and gld like '" + Trim(Alpha.Text) + "%' GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
               rs6C.Open "select gld, sum(a) as SAmount from Cnf1B where  cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "') and dc ='C' and gld like '" + Trim(Alpha.Text) + "%'  GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
               'dnfb
               rs7.Open "select gld, sum(a) as SAmount from dnfB where  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "') and dc ='D' and gld like '" + Trim(Alpha.Text) + "%' GROUP BY gld", con, adOpenKeyset, adLockReadOnly, adCmdText
               rs7C.Open "select gld, sum(a) as SAmount from dnfB where dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "') and dc ='C'  and gld like '" + Trim(Alpha.Text) + "%'  GROUP BY gld", con, adOpenKeyset, adLockReadOnly, adCmdText
                
                              
                              

         Else
                con.Execute "INSERT INTO TemprptTrialBalance (Gledger, OpeningBalance ) SELECT Gledger.gledger, Gledger.YEAROPENING from Gledger where Gledger.gledger = '" + Trim(COMBOGENLEDGER.Text) + "'"
                RSopBalDR.Open "SELECT GenLedger, sum(amount)as DAmount  FROM VOUCHERS  where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate<cdate('" + Trim(date1.Text) + "') and DebitorCredit='D'  GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
                RSopBalCR.Open "SELECT GenLedger,  sum(amount) as CAmount  FROM VOUCHERS  where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate<cdate('" + Trim(date1.Text) + "') and DebitorCredit='C' GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
        
                RsVDr.Open "SELECT GenLedger,   sum(amount) as DAmount  FROM VOUCHERS  where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate >= cdate('" + Trim(date1.Text) + "') and VoucherDate <= cdate('" + Trim(date2.Text) + "') and DebitorCredit='D' GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
                RsVCr.Open "SELECT GenLedger,  sum(amount) as CAmount  FROM VOUCHERS  where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate >= cdate('" + Trim(date1.Text) + "') and VoucherDate <= cdate('" + Trim(date2.Text) + "') and DebitorCredit='C' GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
       
        
                Ors1.Open "select GenLedger,  sum(netamount) as SAmount from invoicea where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE < cdate('" + Trim(date1.Text) + "') GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
                Ors2.Open "select GenLedger,   sum(netamount) as  SAmount from CREDITa where genledger='" + Trim(COMBOGENLEDGER.Text) + "'  and INVOICEDATE < cdate('" + Trim(date1.Text) + "') GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
                Ors3.Open "select GenLedger,   sum(NETamount) as SAmount  from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE < cdate('" + Trim(date1.Text) + "') GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
                Ors3cash.Open "select GenLedger, sum(baa) as SAmount  from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE < cdate('" + Trim(date1.Text) + "') GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
  
                Ors4.Open "select pgld, sum(na) as SAmount  from Cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "'  and cnd<cdate('" + Trim(date1.Text) + "')   and  dc = 'D' GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
                Ors41.Open "select pgld, sum(na) as SAmount  from Cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "'  and cnd<cdate('" + Trim(date1.Text) + "') and  dc = 'C'  GROUP BY pgld  ", con, adOpenKeyset, adLockReadOnly, adCmdText
                
                Ors5.Open "select pgld, sum(na) as SAmount  from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "')  and   dc = 'D'  GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
                Ors51.Open "select pgld, sum(na) as SAmount  from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "')  and   dc = 'C'  GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
                
                Ors6.Open "select gld, sum(a) as SAmount from Cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and  cnd < cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')  and   dc = 'D'  GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
                Ors61.Open "select gld, sum(a) as SAmount from Cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and  cnd < cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')  and   dc = 'C'  GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
                
                Ors7.Open "select gld, sum(a) as SAmount from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and  dnd < cdate('" + Trim(date1.Text) + "')  and dc='D'  GROUP BY gld  ", con, adOpenKeyset, adLockReadOnly, adCmdText
                Ors71.Open "select gld, sum(a) as SAmount from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and  dnd < cdate('" + Trim(date1.Text) + "') and dc='D'  GROUP BY gld  ", con, adOpenKeyset, adLockReadOnly, adCmdText
          
               ' opening balance end
        
               'start dr or cr amount for trialbalance
               rs1.Open "select GenLedger,  sum(netamount) as SAmount from invoicea where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
               rs2.Open "select GenLedger,   sum(netamount) as  SAmount from CREDITa where genledger='" + Trim(COMBOGENLEDGER.Text) + "'  and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               rs3.Open "select GenLedger,  sum(NETamount) as SAmount  from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
               rs3cash.Open "select GenLedger,   sum(baa) as SAmount  from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               rs4.Open "select pgld, sum(na) as SAmount  from Cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "'  and cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "') and dc ='D' GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
               rs4c.Open "select pgld, sum(na) as SAmount  from Cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "'  and cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "') and dc ='C' GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               
               rs5.Open "select pgld, sum(na) as SAmount from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "') and dc ='D'  GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
               rs5C.Open "select pgld, sum(na) as SAmount from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "') and  dc ='C'  GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               rs6.Open "select gld, sum(a) as SAmount from Cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and  cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "') and dc ='D'  GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
               rs6C.Open "select gld, sum(a) as SAmount from Cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and  cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "') and dc ='C' GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
               
               rs7.Open "select gld, sum(a) as SAmount from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "') and dc ='D' GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
               rs7C.Open "select gld, sum(a) as SAmount from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "') and dc ='C' GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
   
        End If
  
  rs.Open "select * from TemprptTrialBalance ", con, adOpenKeyset, adLockReadOnly, adCmdText
  If Not rs.BOF Then
            
            Do While Not rs.EOF
               OPBALANCE = rs!openingbalance
               SDamount = 0
               SCamount = 0
            
              
              If RSopBalDR.RecordCount > 0 Then
                RSopBalDR.Find "genledger='" + Trim(rs!gledger) + "'"
                  If Not RSopBalDR.EOF Then
                      OPBALANCE = OPBALANCE + RSopBalDR!DAMOUNT
                      viewsubledger = True
                  End If
              End If

              If RSopBalCR.RecordCount > 0 Then
                  RSopBalCR.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not RSopBalCR.EOF Then
                      OPBALANCE = OPBALANCE - RSopBalCR!Camount
                  End If
              End If
            
            If Ors1.RecordCount > 0 Then
                  Ors1.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not Ors1.EOF Then
                      OPBALANCE = OPBALANCE + Ors1!Samount
                  End If
            End If
              
            If Ors2.RecordCount > 0 Then
                  Ors2.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not Ors2.EOF Then
                      OPBALANCE = OPBALANCE + Ors2!Samount
                    
                  End If
              End If
                
            If Ors3.RecordCount > 0 Then
                  Ors3.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not Ors3.EOF Then
                      OPBALANCE = OPBALANCE + Ors3!Samount
                     
                  End If
            End If
            
            If Ors3cash.RecordCount > 0 Then
                  Ors3cash.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not Ors3cash.EOF Then
                      OPBALANCE = OPBALANCE - Ors3cash!Samount
                     
                  End If
            End If
            
            If Ors4.RecordCount > 0 Then
                  Ors4.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not Ors4.EOF Then
                      OPBALANCE = OPBALANCE + Ors4!Samount
                  End If
            End If
            
            If Ors41.RecordCount > 0 Then
                  Ors41.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not Ors41.EOF Then
                      OPBALANCE = OPBALANCE - Ors41!Samount
                  End If
            End If
            
         
            
            
            If Ors5.RecordCount > 0 Then
                  Ors5.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not Ors5.EOF Then
                      OPBALANCE = OPBALANCE + Ors5!Samount
                  End If
            End If
            If Ors51.RecordCount > 0 Then
                  Ors51.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not Ors51.EOF Then
                      OPBALANCE = OPBALANCE - Ors51!Samount
                  End If
            End If





            If Ors6.RecordCount > 0 Then
                  Ors6.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not Ors6.EOF Then
                      OPBALANCE = OPBALANCE + Ors6!Samount
                  End If
            End If

            If Ors61.RecordCount > 0 Then
                  Ors61.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not Ors61.EOF Then
                      OPBALANCE = OPBALANCE + Ors61!Samount
                  End If
            End If
            
            
            
            
            If Ors7.RecordCount > 0 Then
                  Ors7.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not Ors7.EOF Then
                      OPBALANCE = OPBALANCE + Ors7!Samount
                  End If
            End If

            If Ors71.RecordCount > 0 Then
                  Ors71.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not Ors71.EOF Then
                      OPBALANCE = OPBALANCE + Ors71!Samount
                  End If
            End If

            
            
            con.Execute "UPDATE TemprptTrialBalance SET OpeningBalance  = " & OPBALANCE & " where gledger ='" + Trim(rs!gledger) + "'"
            If RsVDr.RecordCount > 0 Then
                  RsVDr.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not RsVDr.EOF Then
                        SDamount = SDamount + RsVDr!DAMOUNT
                 End If
            End If

            If RsVCr.RecordCount > 0 Then
                  RsVCr.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not RsVCr.EOF Then
                        SCamount = SCamount + RsVCr!Camount
                 End If
            End If

            If rs1.RecordCount > 0 Then
                  rs1.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not rs1.EOF Then
                      SDamount = SDamount + rs1!Samount
                  End If
            End If
            
            If rs2.RecordCount > 0 Then
                  rs2.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not rs2.EOF Then
                      SCamount = SCamount + rs2!Samount
                  End If
            End If
            
            If rs3.RecordCount > 0 Then
                  rs3.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not rs3.EOF Then
                      SDamount = SDamount + rs3!Samount
                  End If
            End If
            
            If rs3cash.RecordCount > 0 Then
                  rs3cash.Find "Genledger='" + Trim(rs!gledger) + "'"
                  If Not rs3cash.EOF Then
                      SCamount = SCamount + rs3cash!Samount
                  End If
            End If
 
            
            If rs4.RecordCount > 0 Then
                  rs4.Find "pgld ='" + Trim(rs!gledger) + "'"
                  If Not rs4.EOF Then
                      SDamount = SDamount + rs4!Samount
                  End If
            End If
            
            If rs4c.RecordCount > 0 Then
                  rs4c.Find "pgld= '" + rs!gledger + "'"
                  If Not rs4c.EOF Then
                      SCamount = SCamount + rs4c!Samount
                  End If
            End If
            
            
            If rs5.RecordCount > 0 Then
                  rs5.Find "pgld='" + Trim(rs!gledger) + "'"
                  If Not rs5.EOF Then
                      SDamount = SDamount + rs5!Samount
                  End If
            End If

            If rs5C.RecordCount > 0 Then
                  rs5C.Find "pgld='" + Trim(rs!gledger) + "'"
                  If Not rs5C.EOF Then
                      SCamount = SCamount + rs5C!Samount
                  End If
            End If
            
            If rs6.RecordCount > 0 Then
                  rs6.Find "gld ='" + Trim(rs!gledger) + "'"
                  If Not rs6.EOF Then
                      SDamount = SDamount + rs6!Samount
                  End If
            End If
            
            If rs6C.RecordCount > 0 Then
                  rs6C.Find "gld='" + Trim(rs!gledger) + "'"
                  If Not rs6C.EOF Then
                      SCamount = SCamount + rs6C!Samount
                  End If
            End If
  
            If rs7.RecordCount > 0 Then
                  rs7.Find "gld='" + Trim(rs!gledger) + "'"
                  If Not rs7.EOF Then
                      SDamount = SDamount + rs7!Samount
                  End If
            End If
            If rs7C.RecordCount > 0 Then
                  rs7C.Find "gld='" + Trim(rs!gledger) + "'"
                  If Not rs7C.EOF Then
                      SCamount = SCamount + rs7C!Samount
                  End If
            End If

  
            con.Execute "UPDATE TemprptTrialBalance SET TemprptTrialBalance.DAmount = TemprptTrialBalance.DAmount+  " & SDamount & " ,TemprptTrialBalance.CAmount = TemprptTrialBalance.CAmount+  " & SCamount & "  where Gledger ='" + Trim(rs!gledger) + "'"
            'con.Execute "UPDATE TemprptTrialBalance SET TemprptTrialBalance.CAmount = TemprptTrialBalance.CAmount+  " & SCamount & " where Gledger='" + Trim(rs!gledger) + "'"
         
         
             
            If RSopBalDR.RecordCount > 0 Then RSopBalDR.MoveFirst
            If RSopBalCR.RecordCount > 0 Then RSopBalCR.MoveFirst
           
            
            If RsVDr.RecordCount > 0 Then RsVDr.MoveFirst
            If RsVCr.RecordCount > 0 Then RsVCr.MoveFirst
            
            If rs1.RecordCount > 0 Then rs1.MoveFirst
            If rs2.RecordCount > 0 Then rs2.MoveFirst
            If rs3.RecordCount > 0 Then rs3.MoveFirst
            If rs3cash.RecordCount > 0 Then rs3cash.MoveFirst
            
            If rs4.RecordCount > 0 Then rs4.MoveFirst
            If rs4c.RecordCount > 0 Then rs4c.MoveFirst
            
            If rs5.RecordCount > 0 Then rs5.MoveFirst
            If rs5C.RecordCount > 0 Then rs5C.MoveFirst
            
            If rs6.RecordCount > 0 Then rs6.MoveFirst
            If rs6C.RecordCount > 0 Then rs6C.MoveFirst
            
            If rs7.RecordCount > 0 Then rs7.MoveFirst
            If rs7C.RecordCount > 0 Then rs7C.MoveFirst
            
            If Ors1.RecordCount > 0 Then rs1.MoveFirst
            If Ors2.RecordCount > 0 Then rs2.MoveFirst
            If Ors3.RecordCount > 0 Then rs3.MoveFirst
            If Ors3cash.RecordCount > 0 Then rs3.MoveFirst
            
            If Ors4.RecordCount > 0 Then rs4.MoveFirst
            If Ors41.RecordCount > 0 Then rs41.MoveFirst
            If Ors5.RecordCount > 0 Then rs5.MoveFirst
            If Ors51.RecordCount > 0 Then rs51.MoveFirst
            
            If Ors6.RecordCount > 0 Then rs6.MoveFirst
            If Ors61.RecordCount > 0 Then rs61.MoveFirst
            
            If Ors7.RecordCount > 0 Then rs7.MoveFirst
            If Ors71.RecordCount > 0 Then rs71.MoveFirst
            
            If Not rs.EOF Then
                rs.MoveNext
            End If
      
      Loop
    End If
      con.CommitTrans
      Genrate
      r1.LoadFile (VB.App.Path + "\print.TXT")
      r1.Visible = True
'End If

If r1.Visible = True Then
        Me.print1.Visible = True
        Me.export.Visible = True
        Me.Combo1.Visible = True
        Me.Command1.Visible = True
End If
End Sub

Private Sub date1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub date1_LostFocus()
    If Trim(date1.Text) <> "" Then
        If Not checkdate(Trim(date1.Text), date1) Then
            date1.SetFocus
        End If
    End If
End Sub

Private Sub date2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub date2_LostFocus()
    If Trim(date2.Text) <> "" Then
        If Not checkdate(Trim(date2.Text), date2) Then
            date2.SetFocus
        End If
    End If
End Sub
Private Sub Form_Load()
Me.r1.Top = 10
Me.r1.Left = 10
Do While Trim(UCase(VB.screen.ActiveForm.Name)) <> Trim(UCase("MainMenu"))
    If Trim(UCase(VB.screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
        Unload VB.screen.ActiveForm
    End If
Loop
Me.Top = 0
Me.Left = 0

Set con = New ADODB.Connection
con.CursorLocation = adUseClient
Set rs = New ADODB.Recordset

    With con
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
        .Open
    End With
    'rs.Open "select * from gledger where slf=true", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs.Open "select * from gledger", con, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            COMBOGENLEDGER.AddItem Trim(rs!gledger)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    

    rs.Close
    rs.Open "setup1", con, adOpenKeyset, adLockReadOnly, adCmdTable
    date1.Text = rs!yarfrom
    date2.Text = rs!yarto
    rs.Close
End Sub

Private Sub print_Click()
Rsinvoicea.Open "select GenLedger,  SubLedger , sum(amount) as INVAmount from invoicea where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') order by invoicedate", con, adOpenKeyset, adLockReadOnly, adCmdText
   
       RsCREDITa.Open "select * from CREDITa where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "')", con, adOpenKeyset, adLockReadOnly, adCmdText
       RsCnf1a.Open "select * from Cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')", con, adOpenKeyset, adLockReadOnly, adCmdText
       Rsdnfa.Open "select * from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "')", con, adOpenKeyset, adLockReadOnly, adCmdText
       RsCnf1B.Open "select * from Cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')", con, adOpenKeyset, adLockReadOnly, adCmdText
       RsdnfB.Open "select * from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "')", con, adOpenKeyset, adLockReadOnly, adCmdText
       RScasha.Open "select * from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') order by invoicedate", con, adOpenKeyset, adLockReadOnly, adCmdText
  
End Sub


Sub Genrate()
    
    Command1.Top = r1.Top + r1.Height + 30
    Combo1.Top = r1.Top + r1.Height + 30
    Set CON1 = New ADODB.Connection
    Set rs = New ADODB.Recordset
    With CON1
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\tchitra.mdb"
        .Open
    End With
   main.reportname = "Sub. Ledger Trial"
    Dim called1, called2 As Boolean



    Dim maxline As Integer
    Dim T1, T2, T3, t4, t5, t6, t7, t8 As Integer
    Dim paperwidth As Integer
    Dim xtemp As String
    Dim Trs As ADODB.Recordset
    Set Trs = New ADODB.Recordset
       paperwidth = 150
        T1 = 10
        T2 = 25
        T3 = 40
        t4 = 55
        t5 = 70
        t6 = 85
        t7 = 100
        t8 = 115
        maxline = 50
        called1 = False
        called2 = False
        Dim line As Integer
        Dim rs1 As ADODB.Recordset
        Dim kkk As ADODB.Recordset
        Dim Pno As Integer
        Dim GopenBal As Double
        Dim GopenDr As Double
        Dim GopenCr As Double
        Dim GopenCl As Double
        GopenBal = 0
        GopenDr = 0
        GopenCr = 0
        GopenCl = 0
        
        Set kkk = New ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        main.reportdata
        main.repors.Find "reportname='" + Trim(main.reportname) + "'"
        maxline = main.repors!totalline
        If main.repors!comp = True Then
            paperwidth = Int(main.repors!totalcolumn * 1.75)
        Else
            paperwidth = main.repors!totalcolumn
        End If
        Open "" + VB.App.Path + "\print.TXT" For Output As #1
        line = 0
        Pno = 1
header:
        Dim I As Integer
        For I = 1 To main.repors!topmargin
            Print #1, ""
            line = line + 1
        Next
        If kkk.State = 1 Then
            kkk.Close
        End If
        kkk.Open "select * from setup1", con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kkk.BOF Then
          
    
            Print #1, Chr(27) + Chr(15) + Chr(14); Tab((((paperwidth - (Len(Trim(kkk!CNAME)) * 2)) / 2) - 15) + LEFTM); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!CNAME)); Chr(27) + Chr(14), Tab(110); "Page No:"; Pno
              
            'Print #1, Tab(((paperwidth - (Len(Trim(kkk!add1)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
            'Print #1, Tab(((paperwidth - Len(Trim(kkk!phone1))) / 2) + LEFTM); Trim(kkk!phone1)
            line = line + 1
        End If
        If Trs.State = 1 Then
            Trs.Close
        End If
        Trs.Open "treport", CON1, adOpenKeyset, adLockReadOnly, adCmdTable
        Print #1, Tab(((paperwidth - (Len(Trim(Trs!header)))) / 2) + LEFTM); Chr(27) + Chr(15) + Chr(14); "Sub Ledger Trial Balance"
        line = line + 1
        Print #1, Tab(LEFTM + ((paperwidth - Len(Trim(COMBOGENLEDGER.Text))) / 2)); Trim(COMBOGENLEDGER.Text)
        Print #1, ""
        line = line + 1
       
        xstr = date1.Text & " To " & date2.Text
        Print #1, Chr(27) + Chr(15) + Chr(14); Tab(LEFTM + ((paperwidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Trim("Period : " + date1.Text & " To " & date2.Text);
        line = line + 1
        Print #1, ""
        line = line + 1
    
        Print #1, Tab(LEFTM); repli("-", paperwidth)
        line = line + 1
        Print #1, Tab(8); "Sub. Ledger Description"; Tab(46); "Opening Balance"; Tab(67); "Amount (Dr.)"; Tab(89); "Amount (Cr.)"; Tab(110); "Closing Balance"
       
        Print #1, Tab(LEFTM); repli("-", paperwidth); Chr(27) + Chr(14)
        line = line + 1
        Print #1, ""
        line = line + 1
    
        Trs.Close
        If called1 Then
            GoTo printagain1
        End If
        If rs.State = 1 Then
            rs.Close
        End If
        
        If rs.State = 1 Then rs.Close
        rs.Open "select * from TemprptTrialBalance where openingbalance<>0   or damount<>0 or cAmount<>0 order by subledger ", con, adOpenKeyset, adLockReadOnly, adCmdText
        Dim CB As Double
           While Not rs.EOF
              CB = 0
              CB = CB + rs(2) + rs(3) - Abs(rs(4))
              Print #1, Tab(1); rs!gledger; Tab(46); IIf(rs!openingbalance <> 0, rsets(Trim(Format(rs!openingbalance, "0.00")), 12), ""); Tab(65); IIf(rs(3) <> 0, rsets(Trim(Format(rs(3), "0.00")), 12), ""); Tab(85); IIf(rs(4) <> 0, rsets(Trim(Format(Str(rs(4)), "0.00")), 12), ""); Tab(110); IIf(CB <> 0, IIf(CB > 0, rsets(Trim(Format(Str(CB), "0.00")), 12) & " Dr. ", rsets(Trim(Format(Str(CB), "0.00")), 12) & "   Cr."), "")
              line = line + 1
               
              GopenBal = GopenBal + rs(2)
              GopenDr = GopenDr + rs(3)
              GopenCr = GopenCr + rs(4)
              GopenCl = GopenCl + rs(2) + rs(3) - Abs(rs(4))
        
               
             
               If line > maxline Then
                        called1 = True
                        GoTo printfooter
printnext:
                        line = 0
                        Pno = Pno + 1
                        Print #1, Chr(12)
                        GoTo header
printagain1:
                        line = 0
                        called1 = False
                End If
                
                If Not rs.EOF Then
                    rs.MoveNext
                End If
              
        
            Wend
            
            
            
printfooter:

             If called1 Then
                GoTo printnext
            End If
            Print #1, Tab(LEFTM); repli("-", paperwidth)
            Print #1, Tab(LEFTM); "* * * NET BALANCE * * * "; Tab(46); IIf(GopenBal <> 0, rsets(Format(Trim(GopenBal), "0.00"), 12), ""); Tab(65); IIf(GopenDr <> 0, rsets(Format(Trim(GopenDr), "0.00"), 12), ""); Tab(85); IIf(GopenCr <> 0, rsets(Format(Trim(GopenCr), "0.00"), 12), ""); Tab(110); IIf(GopenCl <> 0, IIf(GopenCl > 0, rsets(Format(Trim(GopenCl), "0.00"), 12) & " Dr.", rsets(Format(Trim(GopenCl), "0.00"), 12) & "   Cr."), "")
            Print #1, Tab(LEFTM); repli("-", paperwidth)
                  
             
            If line < maxline Then
               Do While line < maxline - 5
                    Print #1, " "
                    line = line + 1
                Loop
            End If
            
            Dim tempdata As ADODB.Recordset
            Set tempdata = New ADODB.Recordset
            tempdata.Open "setup1", con, adOpenKeyset, adLockReadOnly, adCmdTable
                   
           Print #1, ""
           Print #1, ""
           Print #1, ""
           If Trs.State = 1 Then
              Trs.Close
           End If
       
        Close #1
End Sub
