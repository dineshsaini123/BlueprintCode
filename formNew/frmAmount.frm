VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAmount 
   ClientHeight    =   4152
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4716
   Icon            =   "frmAmount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4152
   ScaleWidth      =   4716
   Begin VB.CommandButton Command2 
      Caption         =   "&Cash In Bank"
      Height          =   435
      Left            =   1620
      TabIndex        =   9
      Top             =   2808
      Width           =   1788
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cash Amt. Print"
      Height          =   435
      Left            =   1620
      TabIndex        =   8
      Top             =   2352
      Width           =   1788
   End
   Begin Crystal.CrystalReport cr 
      Left            =   180
      Top             =   2712
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtAmt 
      Height          =   315
      Left            =   1620
      TabIndex        =   6
      Top             =   1332
      Width           =   1395
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   468
      Left            =   1620
      TabIndex        =   5
      Top             =   3336
      Width           =   1788
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Bill Amt. Print "
      Height          =   435
      Left            =   1620
      TabIndex        =   4
      Top             =   1872
      Width           =   1788
   End
   Begin MSComCtl2.DTPicker fromDate 
      Height          =   312
      Left            =   1620
      TabIndex        =   0
      Top             =   312
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   550
      _Version        =   393216
      Format          =   544407553
      CurrentDate     =   42539
   End
   Begin MSComCtl2.DTPicker ToDate 
      Height          =   312
      Left            =   1620
      TabIndex        =   1
      Top             =   852
      Width           =   1452
      _ExtentX        =   2561
      _ExtentY        =   550
      _Version        =   393216
      Format          =   544407553
      CurrentDate     =   42539
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill/Cash Amount"
      Height          =   252
      Left            =   300
      TabIndex        =   7
      Top             =   1332
      Width           =   1332
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date :"
      Height          =   252
      Left            =   300
      TabIndex        =   3
      Top             =   852
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date :"
      Height          =   252
      Left            =   300
      TabIndex        =   2
      Top             =   312
      Width           =   1092
   End
End
Attribute VB_Name = "frmAmount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Unload Me
End Sub
Private Sub cmdPrint_Click()

DSNNew

If LCase(com_user) = LCase("satyam.CHITRA") Then

CR.Reset
CR.ReportFileName = rptPath & "\inv_Cash.rpt"
CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
CR.ReplaceSelectionFormula "{Sale_CashList.netamount}>=" & Val(txtAmt) & " and ({Sale_CashList.invoicedate}>=datevalue('" & Format(fromdate.value, "MM/dd/yyyy") & "') and {Sale_CashList.invoicedate}<=datevalue('" & Format(todate.value, "dd/MM/yyyy") & "'))"
CR.Formulas(0) = "bal1=" & Val(txtAmt) & ""
CR.Formulas(1) = "fdate='" & fromdate.value & "'"
CR.Formulas(2) = "tdate='" & todate.value & "'"
CR.WindowShowPrintSetupBtn = True
CR.WindowShowPrintBtn = True
CR.WindowState = crptMaximized
CR.Action = 1


Else


CR.Reset
CR.ReportFileName = rptPath & "\inv_Cash.rpt"
CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
CR.ReplaceSelectionFormula "{Sale_CashList.netamount}>=" & Val(txtAmt) & " and ({Sale_CashList.invoicedate}>=datevalue('" & Format(fromdate.value, "MM/dd/yyyy") & "') and {Sale_CashList.invoicedate}<=datevalue('" & Format(todate.value, "MM/dd/yyyy") & "'))"
CR.Formulas(0) = "bal1=" & Val(txtAmt) & ""
CR.Formulas(1) = "fdate='" & fromdate.value & "'"
CR.Formulas(2) = "tdate='" & todate.value & "'"
CR.WindowShowPrintSetupBtn = True
CR.WindowShowPrintBtn = True
CR.WindowState = crptMaximized
CR.Action = 1

End If

End Sub
Private Sub Command1_Click()

Screen.MousePointer = vbHourglass

Dim sum1 As Double
Dim kk2 As Integer


con.Execute "update SLEDGER set cashamt1=0 where gledger='SUNDRY DEBTORS'"

If RS.State = 1 Then RS.close
RS.Open "select  PartyName from ReceiveIssueParty group by PartyName", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False

sum1 = 0

If rs1.State = 1 Then rs1.close
rs1.Open "select  Particullar,cr from ReceiveIssueParty where PartyName='" & RS!partyname & "'", con, adOpenDynamic, adLockOptimistic
While rs1.EOF = False

kk2 = InStr(rs1!Particullar, "cash")
If kk2 >= 1 Then
   sum1 = sum1 + rs1!CR
End If
rs1.MoveNext

Wend

If sum1 > 0 Then
   con.Execute "update SLEDGER set cashamt1=" & sum1 & " where SubLEDGER='" & RS!partyname & "'"
End If

RS.MoveNext
Wend

Screen.MousePointer = vbDefault
DSNNew

If MsgBox("Want to View ?", vbQuestion + vbYesNo) = vbYes Then

CR.Reset
CR.ReportFileName = rptPath & "\CashAmount.rpt"
CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
CR.ReplaceSelectionFormula "{Sale_CashList.cashAmt1}>=" & Val(txtAmt) & ""
CR.Formulas(0) = "bal1=" & Val(txtAmt) & ""
CR.Formulas(1) = "fdate='" & fromdate.value & "'"
CR.Formulas(2) = "tdate='" & todate.value & "'"
CR.WindowShowPrintSetupBtn = True
CR.WindowShowPrintBtn = True
CR.WindowState = crptMaximized
CR.Action = 1

End If



End Sub

Private Sub Command2_Click()

Screen.MousePointer = vbHourglass

Dim sum1 As Double
Dim kk2 As Integer


con.Execute "delete from treport_net_bk"

If RS.State = 1 Then RS.close
RS.Open "SELECT [VoucherType],[VoucherDate],[VoucherNumber],[GenLedger],[SubLedger]," & _
" [Amount],[DESCRIPTION],[DebitorCredit] FROM VOUCHERS " & _
" where (GenLedger like '%Bank%'  and [DESCRIPTION] like '%cash%' and DebitorCredit='D')", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False

con.Execute "insert into treport_net_bk(vtype,vdate,vno,genledger,ad,narration,text)" & _
" values('" & RS!VoucherType & "','" & Format(RS!voucherDATE, "MM/dd/yyyy") & "','" & RS!VOUCHERNUMBER & "','" & RS!Genledger & "','" & RS!amount & "','" & RS!DESCRIPTION & "','" & RS!DebitorCredit & "')"

RS.MoveNext
Wend

Screen.MousePointer = vbDefault
DSNNew

If MsgBox("Want to View ?", vbQuestion + vbYesNo) = vbYes Then

CR.Reset
CR.ReportFileName = rptPath & "\Cash_in_bank.rpt"
CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
CR.ReplaceSelectionFormula "{Sale_CashList.ad}>=" & Val(txtAmt) & ""
CR.Formulas(0) = "bal1=" & Val(txtAmt) & ""
CR.Formulas(1) = "fdate='" & fromdate.value & "'"
CR.Formulas(2) = "tdate='" & todate.value & "'"
CR.WindowShowPrintSetupBtn = True
CR.WindowShowPrintBtn = True
CR.WindowState = crptMaximized
CR.Action = 1

End If



End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Unload Me
End If
End Sub
Private Sub Form_Load()



    If RS.State = 1 Then RS.close
    RS.Open "setup1", con, adOpenDynamic, adLockReadOnly, adCmdTable
    fromdate.value = RS!yarfrom
    todate.value = RS!yarto
  
  Me.top = 1200
  Me.Left = 1200
  Me.Width = 5000
  Me.Height = 4400
  
  BackColorFrom Me
End Sub

