VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form CASHBOOK 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4212
   ClientLeft      =   252
   ClientTop       =   1416
   ClientWidth     =   8196
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "CASHBOOK.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4212
   ScaleWidth      =   8196
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3_cash 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Date Wise Cash Summary"
      Height          =   564
      Left            =   1440
      Picture         =   "CASHBOOK.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3132
      Visible         =   0   'False
      Width           =   1860
   End
   Begin Crystal.CrystalReport cr 
      Left            =   225
      Top             =   2520
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Party Wise Monthly Cash Receipt"
      Enabled         =   0   'False
      Height          =   630
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2475
      Width           =   1860
   End
   Begin VB.CommandButton cmdCashSale 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cash Sale Summary"
      Enabled         =   0   'False
      Height          =   630
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2475
      Width           =   1770
   End
   Begin VB.CommandButton cmdCashReceipt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cash Receipt Summary"
      Enabled         =   0   'False
      Height          =   630
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2475
      Width           =   1725
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   3780
      Visible         =   0   'False
      Width           =   8130
      _ExtentX        =   14330
      _ExtentY        =   550
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Check Cash Book"
      Height          =   675
      Left            =   1440
      Picture         =   "CASHBOOK.frx":0BF0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1725
      Width           =   1860
   End
   Begin VB.CommandButton Commandreturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   675
      Left            =   5100
      Picture         =   "CASHBOOK.frx":17D4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1725
      Width           =   1770
   End
   Begin VB.CommandButton Commandshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Show"
      Height          =   675
      Left            =   3375
      Picture         =   "CASHBOOK.frx":23B8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1725
      Width           =   1725
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   2535
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   375
      Width           =   3885
   End
   Begin MSMask.MaskEdBox date1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   2535
      TabIndex        =   1
      Top             =   885
      Width           =   1155
      _ExtentX        =   2053
      _ExtentY        =   550
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox date2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   5085
      TabIndex        =   2
      Top             =   885
      Width           =   1290
      _ExtentX        =   2265
      _ExtentY        =   550
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " - To - "
      Height          =   315
      Left            =   4185
      TabIndex        =   7
      Top             =   915
      Width           =   585
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From The Date"
      Height          =   315
      Left            =   345
      TabIndex        =   6
      Top             =   915
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gen. Ledger Desc."
      Height          =   285
      Left            =   375
      TabIndex        =   4
      Top             =   405
      Width           =   2055
   End
End
Attribute VB_Name = "CASHBOOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As Connection
Public CheckCash As Boolean
Dim RS As Recordset
Private Sub cmdCashReceipt_Click()
    
    
    con.Execute "delete from treport where (genledger <> 'CASH-IN-HAND' and genledger <> 'SUNDRY DEBTORS')"
    con.Execute "delete from treport where (dorc = 'D')"
    con.Execute "delete from treport where subledger is null"
    
    If MsgBox("want to view ?", vbQuestion + vbYesNo) = vbYes Then
    cr.Reset
    cr.ReportFileName = rptPath & "/CashSummary.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.ReplaceSelectionFormula "{treport.genledger} = 'SUNDRY DEBTORS' and {treport.vtype} = 'R'"
    cr.Formulas(0) = "pname='CASH SALE'"
    cr.WindowShowExportBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowPrintBtn = True
    cr.WindowState = crptMaximized
    cr.WindowShowRefreshBtn = True
    cr.Action = 1
    
    End If
    
End Sub

Private Sub cmdCashSale_Click()
    con.Execute "delete from treport where (genledger <> 'CASH-IN-HAND' and genledger <> 'SUNDRY DEBTORS')"
    con.Execute "delete from treport where (dorc = 'D')"
    con.Execute "delete from treport where subledger is null"
    
    If MsgBox("want to view ?", vbQuestion + vbYesNo) = vbYes Then
    cr.Reset
    cr.ReportFileName = rptPath & "/CashSummary.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.ReplaceSelectionFormula "{treport.genledger} = 'CASH-IN-HAND' and {treport.vtype} = 'S'"
    cr.Formulas(0) = "pname='CASH SALE'"
    cr.WindowShowExportBtn = True
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowPrintBtn = True
    cr.WindowState = crptMaximized
    cr.WindowShowRefreshBtn = True
    cr.Action = 1
    
    End If

End Sub

Private Sub COMBOGENLEDGER_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
      sendkeys "{Down}"
      sendkeys "{tab}"
End If



End Sub

Private Sub Combosubledger_GotFocus()
    If Trim(COMBOGENLEDGER.text) = "" Then
        COMBOGENLEDGER.SetFocus
    End If
End Sub

Private Sub Combosubledger_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
sendkeys "{tab}"

End If
End Sub

Private Sub Combosubledger_LostFocus()
If Trim(Combosubledger.text) <> "" Then
    If Trim(COMBOGENLEDGER.text) <> "" Then
        If RS.State = 1 Then
            RS.Close
        End If
        RS.Open "select * from cbmf where gld='" + Trim(COMBOGENLEDGER.text) + "' and sld='" + Trim(Combosubledger.text) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
        If RS.BOF Then
            Combosubledger.SetFocus
        End If
        RS.Close
    Else
        Combosubledger.text = ""
    End If
    
End If
End Sub

Private Sub Command1_Click()

CheckCash = True
ledger_ = 0
s1 = 501
Commandshow_Click


End Sub

Private Sub Command2_Click()
    
    DSNNew
    
    con.Execute "delete from treport where (genledger <> 'CASH-IN-HAND' and genledger <> 'SUNDRY DEBTORS')"
    con.Execute "delete from treport where (dorc = 'D')"
    con.Execute "delete from treport where subledger is null"
    
    If MsgBox("want to view ?", vbQuestion + vbYesNo) = vbYes Then
       cr.Reset
       cr.ReportFileName = rptPath & "/CashReceiptMonthly.rpt"
       cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
       cr.WindowShowExportBtn = True
       cr.WindowShowPrintSetupBtn = True
       cr.WindowShowPrintBtn = True
       cr.WindowShowRefreshBtn = True
       cr.WindowState = crptMaximized
       cr.Action = 1
       
    End If
    
    
End Sub



Private Sub Command3_cash_Click()


'Commandshow.Enabled = False
's1 = 501
'
'If Trim(COMBOGENLEDGER.text) <> "" Then
'    If DateDiff("d", Trim(date1.text), Trim(date2.text)) <= 0 Then
'        MsgBox "invalid date"
'        Exit Sub
'    End If
'    Dim vflag As Boolean
'    Dim rs1 As ADODB.Recordset
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim Balance As Double
'
'    Balance = 0
'
'' opening balance start
'   If RS.State = 1 Then RS.Close
'    RS.Open "select * from gledger where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
'    Balance = Round(RS!YEAROPENING, 2)
'    RS.Close
'  ' current balance start
'
'    RS.Open "select sum(amount) from vouchers where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VoucherDate,103)<convert(smalldatetime,'" + Trim(date1.text) + "',103) and DebitorCredit='D' and vouchertype='J'"
'    If RS(0) <> 0 Then
'        Balance = Balance + RS(0)
'    End If
'    RS.Close
'    RS.Open "select sum(amount) from vouchers where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VoucherDate,103)<convert(smalldatetime,'" + Trim(date1.text) + "',103) and DebitorCredit='C' and vouchertype='J'"
'    If RS(0) <> 0 Then
'        Balance = Balance - RS(0)
'    End If
'    RS.Close
'
'    RS.Open "select sum(amount) from vouchers where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VoucherDate,103)<convert(smalldatetime,'" + Trim(date1.text) + "',103) and DebitorCredit='D' and vouchertype<>'J'   and cashcheck= 1 "
'    If RS(0) <> 0 Then
'        Balance = Balance + RS(0)
'    End If
'    RS.Close
'    RS.Open "select sum(amount) from vouchers where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VoucherDate,103)<convert(smalldatetime,'" + Trim(date1.text) + "',103) and DebitorCredit='C' and vouchertype<>'J'  and  cashcheck= 1 "
'    If RS(0) <> 0 Then
'        Balance = Balance - RS(0)
'    End If
'    RS.Close
'
'    RS.Open "select sum(baa) as aa  from CASHa where " & stringyear & " and baa <> 0 and convert(smalldatetime,INVOICEDATE,103)<convert(smalldatetime,'" + Trim(date1.text) + "',103)"
'    If RS(0) <> 0 Then
'        Balance = Balance + RS(0)
'    End If
'    RS.Close
'
'    RS.Open "select sum(AMOUNT) from cashC where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,INVOICEDATE,103)<convert(smalldatetime,'" + Trim(date1.text) + "',103)"
'    If RS(0) <> 0 Then
'        Balance = Balance + RS(0)
'    End If
'    RS.Close
'    'current balance end'
'
'
'    Set rs1 = New ADODB.Recordset
'    con.Execute ("delete from treport where " & stringyear)
'    rs1.Open "select * from treport where " & stringyear & "", con, adOpenStatic, adLockOptimistic
'    rs1.AddNew
'    rs1!text = "** Opening Balance as on " + Trim(date1.text)
'    rs1!ad = Balance
'    rs1!Period = Trim(date1.text) + "  To  " + Trim(date2.text)
'    rs1!header = "CASH BOOK"
'    rs1!dorc = "D"
'
'    rs1!setupid = setupid
'    rs1!fyear = session
'    rs1.update
'
'    Dim date_condition, date_conditionINV As String
'
'    date_condition = "convert(smalldatetime,voucherDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,voucherDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)"
'
'
'    con.Execute " INSERT INTO treport ( vdate, genledger, subledger, vtype, vno, ad, dorc, narration, cbno ) " & _
'    " SELECT VoucherDate, VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber," & _
'    " VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.DESCRIPTION, VOUCHERS.CBND from VOUCHERS " & _
'    "WHERE " & stringyear & " and VOUCHERS.GenLedger = '" + Trim(COMBOGENLEDGER.text) + "' and " & date_condition & " and VOUCHERS.DebitorCredit='D'  and  VOUCHERS.VoucherType ='J'ORDER BY VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
'
'    con.Execute " INSERT INTO treport ( vdate, genledger, subledger, vtype, vno, aC, dorc, narration, cbno,setupid,fyear ) SELECT VOUCHERS.VoucherDate," & _
'    " VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.Amount, VOUCHERS.DebitorCredit," & _
'    " VOUCHERS.DESCRIPTION, VOUCHERS.CBND," & setupid & "," & session & " from VOUCHERS WHERE " & stringyear & " and VOUCHERS.GenLedger = '" + Trim(COMBOGENLEDGER.text) + "' and  " & date_condition & "" & _
'    "  and  VOUCHERS.DebitorCredit='C'and   VOUCHERS.VoucherType ='J'ORDER BY VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
'
'    con.Execute " INSERT INTO treport ( vdate, genledger, subledger, vtype, vno, ad, dorc, narration, cbno,setupid,fyear ) SELECT VoucherDate, VOUCHERS.GenLedger," & _
'    " VOUCHERS.SubLedger, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.DESCRIPTION," & _
'    " VOUCHERS.CBND," & setupid & "," & session & " from VOUCHERS WHERE " & stringyear & " and VOUCHERS.GenLedger <> '" + Trim(COMBOGENLEDGER.text) + "' and  " & date_condition & " and " & _
'    " VOUCHERS.DebitorCredit='D'  and  VOUCHERS.VoucherType <>'J' and vouchers.cashcheck = 1 ORDER BY VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
'
'    con.Execute " INSERT INTO treport ( vdate, genledger, subledger, vtype, vno, aC, dorc, narration, cbno,setupid,fyear ) SELECT VoucherDate, VOUCHERS.GenLedger," & _
'    " VOUCHERS.SubLedger, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.DESCRIPTION, VOUCHERS.CBND" & _
'    "," & setupid & "," & session & " from VOUCHERS WHERE " & stringyear & " and VOUCHERS.GenLedger <> '" + Trim(COMBOGENLEDGER.text) + "' and  " & date_condition & " and  VOUCHERS.DebitorCredit='C'and " & _
'    "  VOUCHERS.VoucherType <>'J' and vouchers.cashcheck= 1 ORDER BY VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
'
'    Dim rs6 As New ADODB.Recordset
'
'
'   '**************** FOR  COUNTER SALE cashc
'   date_conditionINV = "convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)"
'
'   rs2.Open "select * from CASHC where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "'  and " & date_conditionINV & "", con, adOpenDynamic, adLockReadOnly, adCmdText
'   If Not rs2.EOF Then
'        Do While Not rs2.EOF
'            rs1.AddNew
'           If rs2!amount > 0 Then
'            rs1!Genledger = rs2!Genledger & ""
'            rs1!subledger = rs2!subledger & ""
'            rs1!vdate = rs2!invoiceDate & ""
'            rs1!vtype = Trim("S")
'            rs1!vno = Trim(rs2!invoiceNo)
'            rs1!aC = rs2!amount
'            rs1!dorc = "C"
'             rs1!narration = "Sales C/M"
'            Balance = Balance - rs2!amount
'            rs1!Balance = Balance
'            rs1.update
'           End If
'           If Not rs2.EOF Then
'               rs2.MoveNext
'           End If
'        Loop
'    End If
'
' 'for CashA
'
'  con.Execute "INSERT INTO treport ( vdate, vno, vtype, genledger, subledger, ac, dorc, narration ) SELECT CASHA.INVOICEDATE, CASHA.INVOICENO," & _
'  " 'S' AS Expr3, 'CASH-IN-HAND' AS Expr4, CASHA.cashpartyname, CASHA.BAA, 'C' AS Expr1, 'Sales C/M' AS Expr2 from CASHA  " & _
'  "Where " & stringyear & " and CASHA.BAA <> 0 and " & date_conditionINV & " ORDER BY CASHA.INVOICEDATE, CASHA.INVOICENO"
'
'  con.Execute "update treport set setupid='" & setupid & "',fyear='" & session & "'"
'
'''  If RS.State = 1 Then RS.close
'''  If ledger_ = 0 Then
'''     viewcash.GenrepNoFooterNew
'''  Else
'''     viewcash.GenrepNoFooterNew_Ledger
'''  End If
'''
'''  PrintOption.Show
'
'sss = Format(date1, "yyyy/mm/dd")
'con.Execute "update treport set vdate='" & sss & "' where genledger is null"
'
'Dim op, cashRec, cashSale, bank, expence As Double
'
'Dim vdate_
'
'op = 0
'cashRec = 0
'
'If RS.State = 1 Then RS.Close
'RS.Open "select ad from treport where genledger is null", con
'If RS.EOF = False Then
'   op = RS(0)
'End If
'
'Dim kk1 As Integer
'kk1 = 1
'
'Dim rss As New ADODB.Recordset
'
''If rsS.State = 1 Then rsS.close
''rsS.Open "select vdate from treport group by vdate", con
''While rsS.EOF = False
''
''
''rsS.MoveNext
''Wend
'
'
'k2 = 1
'If RS.State = 1 Then RS.Close
'RS.Open "select vdate from treport group by vdate order by vdate", con
'While RS.EOF = False
'con.Execute "update treport set userid=" & k2 & " where convert(smalldatetime,VDATE,103)=convert(smalldatetime,'" + Trim(RS!vdate) + "',103)"
'k2 = k2 + 1
'RS.MoveNext
'Wend
'
'
'If RS.State = 1 Then RS.Close
'RS.Open "select vdate,userid from treport group by vdate,userid order by vdate", con
'While RS.EOF = False
'
'cashRec = 0
'cashSale = 0
'bank = 0
'expence = 0
'
'
'
'
'If kk1 > 1 Then
'
'If rs1.State = 1 Then rs1.Close
''rs1.Open "select OpeningBalance from treport where (convert(smalldatetime,VDATE,103)=convert(smalldatetime,'" + Trim(RS!vdate) + "',103))", con
'rs1.Open "select OpeningBalance from treport where userid=" & kk1 - 1 & "", con
'If Not IsNull(rs1(0)) Then
'   op = rs1(0)
'End If
'
'
'End If
'
'
'
'If rs1.State = 1 Then rs1.Close
'rs1.Open "select sum(ac) from treport where (vtype='S' and  userid=" & (RS!userid - 1) & ")", con
'If Not IsNull(rs1(0)) Then
'   cashSale = rs1(0)
'End If
'
'If rs1.State = 1 Then rs1.Close
'rs1.Open "select sum(ac) from treport where (vtype='R'  and   userid=" & (RS!userid - 1) & ")", con
'If Not IsNull(rs1(0)) Then
'   cashRec = rs1(0)
'End If
'
'If rs1.State = 1 Then rs1.Close
'rs1.Open "select sum(ad) from treport where ((vtype='J' or vtype='P') and CHARINDEX('Bank', treport.genledger) > 0 and  userid=" & (RS!userid - 1) & ")", con
'If Not IsNull(rs1(0)) Then
'   bank = rs1(0)
'End If
'
'If rs1.State = 1 Then rs1.Close
'rs1.Open "select (sum(ad)+sum(ac)) from treport where ((vtype='J' or vtype='P') and  CHARINDEX('Bank', treport.genledger) =0 and  userid=" & (RS!userid - 1) & ")", con
'If Not IsNull(rs1(0)) Then
'   expence = rs1(0)
'End If
'
'
'
'cashSale = ((cashSale + cashRec + op) - (bank + expence))
''
''If rs1.State = 1 Then rs1.close
''rs1.Open "select vdate from treport where (userid=" & RS!userid + 1 & ")", con
''If rs1.EOF = False Then
''   vdate_ = rs1!vdate
''End If
'
'
''con.Execute "update treport set OpeningBalance='" & cashSale & "' where convert(smalldatetime,VDATE,103)=convert(smalldatetime,'" + Trim(RS!vdate) + "',103)"
'If kk1 = 1 Then
'  con.Execute "update treport set OpeningBalance='" & op & "' where userid=" & (RS!userid) & ""
'Else
'  con.Execute "update treport set OpeningBalance='" & cashSale & "' where userid=" & (RS!userid) & ""
'
'End If
'
'kk1 = kk1 + 1
'
'RS.MoveNext
'Wend
'
'
' If MsgBox("want to view ?", vbQuestion + vbYesNo) = vbYes Then
'    cr.Reset
'
'    cr.ReportFileName = rptPath & "/CASHBOOK_datewise.rpt"
'    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
'
'    'cr.ReplaceSelectionFormula "{treport.genledger} = 'CASH-IN-HAND' and {treport.vtype} = 'S'"
'    'cr.Formulas(0) = "pname='CASH SALE'"
'    cr.Formulas(0) = "fromdate='" & date1_ & "'"
'    cr.Formulas(1) = "todate='" & date2_ & "'"
'
'    cr.WindowShowExportBtn = True
'    cr.WindowShowPrintSetupBtn = True
'    cr.WindowShowPrintBtn = True
'    cr.WindowState = crptMaximized
'    cr.Action = 1
'
'    End If
'
'
'
'
'
''' end code
'
'
'Else
'    MsgBox "gen. ledger or sub. ledger not selected"
'End If
'
'
'
'Commandshow.Enabled = True
'
'cmdCashSale.Enabled = True
'cmdCashReceipt.Enabled = True
'Command2.Enabled = True



End Sub

Private Sub CommandReturn_Click()
    Unload Me
    
End Sub

Private Sub Commandshow_Click()

Commandshow.Enabled = False
s1 = 501

If Trim(COMBOGENLEDGER.text) <> "" Then
    If DateDiff("d", Trim(date1.text), Trim(date2.text)) <= 0 Then
        MsgBox "invalid date"
        Exit Sub
    End If
    Dim vflag As Boolean
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim Balance As Double
    
    Balance = 0
    
' opening balance start
   If RS.State = 1 Then RS.Close
    RS.Open "select * from gledger where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "'", con, adOpenDynamic, adLockReadOnly, adCmdText
    Balance = Round(RS!YEAROPENING, 2)
    RS.Close
  ' current balance start

    RS.Open "select sum(amount) from vouchers where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VoucherDate,103)<convert(smalldatetime,'" + Trim(date1.text) + "',103) and DebitorCredit='D' and vouchertype='J'"
    If RS(0) <> 0 Then
        Balance = Balance + RS(0)
    End If
    RS.Close
    RS.Open "select sum(amount) from vouchers where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VoucherDate,103)<convert(smalldatetime,'" + Trim(date1.text) + "',103) and DebitorCredit='C' and vouchertype='J'"
    If RS(0) <> 0 Then
        Balance = Balance - RS(0)
    End If
    RS.Close

    RS.Open "select sum(amount) from vouchers where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VoucherDate,103)<convert(smalldatetime,'" + Trim(date1.text) + "',103) and DebitorCredit='D' and vouchertype<>'J'   and cashcheck= 1 "
    If RS(0) <> 0 Then
        Balance = Balance + RS(0)
    End If
    RS.Close
    RS.Open "select sum(amount) from vouchers where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VoucherDate,103)<convert(smalldatetime,'" + Trim(date1.text) + "',103) and DebitorCredit='C' and vouchertype<>'J'  and  cashcheck= 1 "
    If RS(0) <> 0 Then
        Balance = Balance - RS(0)
    End If
    RS.Close
   
    RS.Open "select sum(baa) as aa  from CASHa where " & stringyear & " and baa <> 0 and convert(smalldatetime,INVOICEDATE,103)<convert(smalldatetime,'" + Trim(date1.text) + "',103)"
    If RS(0) <> 0 Then
        Balance = Balance + RS(0)
    End If
    RS.Close
   
    RS.Open "select sum(AMOUNT) from cashC where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,INVOICEDATE,103)<convert(smalldatetime,'" + Trim(date1.text) + "',103)"
    If RS(0) <> 0 Then
        Balance = Balance + RS(0)
    End If
    RS.Close
    'current balance end'
    
    
    Set rs1 = New ADODB.Recordset
    con.Execute ("delete from treport where " & stringyear)
    rs1.Open "select * from treport where " & stringyear & "", con, adOpenStatic, adLockOptimistic
    rs1.AddNew
    rs1!text = "** Opening Balance as on " + Trim(date1.text)
    rs1!ad = Balance
    rs1!Period = Trim(date1.text) + "  To  " + Trim(date2.text)
    rs1!header = "CASH BOOK"
    rs1!dorc = "D"
    
    rs1!setupid = setupid
    rs1!fyear = session
    rs1.update
    
    Dim date_condition, date_conditionINV As String
    
    date_condition = "convert(smalldatetime,voucherDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,voucherDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)"
    
   
    con.Execute " INSERT INTO treport ( vdate, genledger, subledger, vtype, vno, ad, dorc, narration, cbno ) " & _
    " SELECT VoucherDate, VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber," & _
    " VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.DESCRIPTION, VOUCHERS.CBND from VOUCHERS " & _
    "WHERE " & stringyear & " and VOUCHERS.GenLedger = '" + Trim(COMBOGENLEDGER.text) + "' and " & date_condition & " and VOUCHERS.DebitorCredit='D'  and  VOUCHERS.VoucherType ='J'ORDER BY VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
    
    con.Execute " INSERT INTO treport ( vdate, genledger, subledger, vtype, vno, aC, dorc, narration, cbno,setupid,fyear ) SELECT VOUCHERS.VoucherDate," & _
    " VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.Amount, VOUCHERS.DebitorCredit," & _
    " VOUCHERS.DESCRIPTION, VOUCHERS.CBND," & setupid & "," & session & " from VOUCHERS WHERE " & stringyear & " and VOUCHERS.GenLedger = '" + Trim(COMBOGENLEDGER.text) + "' and  " & date_condition & "" & _
    "  and  VOUCHERS.DebitorCredit='C'and   VOUCHERS.VoucherType ='J'ORDER BY VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
    
    con.Execute " INSERT INTO treport ( vdate, genledger, subledger, vtype, vno, ad, dorc, narration, cbno,setupid,fyear ) SELECT VoucherDate, VOUCHERS.GenLedger," & _
    " VOUCHERS.SubLedger, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.DESCRIPTION," & _
    " VOUCHERS.CBND," & setupid & "," & session & " from VOUCHERS WHERE " & stringyear & " and VOUCHERS.GenLedger <> '" + Trim(COMBOGENLEDGER.text) + "' and  " & date_condition & " and " & _
    " VOUCHERS.DebitorCredit='D'  and  VOUCHERS.VoucherType <>'J' and vouchers.cashcheck = 1 ORDER BY VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
    
    con.Execute " INSERT INTO treport ( vdate, genledger, subledger, vtype, vno, aC, dorc, narration, cbno,setupid,fyear ) SELECT VoucherDate, VOUCHERS.GenLedger," & _
    " VOUCHERS.SubLedger, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.DESCRIPTION, VOUCHERS.CBND" & _
    "," & setupid & "," & session & " from VOUCHERS WHERE " & stringyear & " and VOUCHERS.GenLedger <> '" + Trim(COMBOGENLEDGER.text) + "' and  " & date_condition & " and  VOUCHERS.DebitorCredit='C'and " & _
    "  VOUCHERS.VoucherType <>'J' and vouchers.cashcheck= 1 ORDER BY VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
    
    Dim rs6 As New ADODB.Recordset
    
   
   '**************** FOR  COUNTER SALE cashc
   date_conditionINV = "convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)"
   
   rs2.Open "select * from CASHC where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "'  and " & date_conditionINV & "", con, adOpenDynamic, adLockReadOnly, adCmdText
   If Not rs2.EOF Then
        Do While Not rs2.EOF
            rs1.AddNew
           If rs2!amount > 0 Then
            rs1!Genledger = rs2!Genledger & ""
            rs1!subledger = rs2!subledger & ""
            rs1!vdate = rs2!invoiceDate & ""
            rs1!vtype = Trim("S")
            rs1!vno = Trim(rs2!invoiceNo)
            rs1!aC = rs2!amount
            rs1!dorc = "C"
             rs1!narration = "Sales C/M"
            Balance = Balance - rs2!amount
            rs1!Balance = Balance
            rs1.update
           End If
           If Not rs2.EOF Then
               rs2.MoveNext
           End If
        Loop
    End If

 'for CashA

  con.Execute "INSERT INTO treport ( vdate, vno, vtype, genledger, subledger, ac, dorc, narration ) SELECT CASHA.INVOICEDATE, CASHA.INVOICENO," & _
  " 'S' AS Expr3, 'CASH-IN-HAND' AS Expr4, CASHA.cashpartyname, CASHA.BAA, 'C' AS Expr1, 'Sales C/M' AS Expr2 from CASHA  " & _
  "Where " & stringyear & " and CASHA.BAA <> 0 and " & date_conditionINV & " ORDER BY CASHA.INVOICEDATE, CASHA.INVOICENO"
  
  con.Execute "update treport set setupid='" & setupid & "',fyear='" & session & "'"
  
  If RS.State = 1 Then RS.Close
  If ledger_ = 0 Then
     viewcash.GenrepNoFooterNew
  Else
     viewcash.GenrepNoFooterNew_Ledger
  End If
  
  PrintOption.Show
Else
    MsgBox "gen. ledger or sub. ledger not selected"
End If



Commandshow.Enabled = True

cmdCashSale.Enabled = True
cmdCashReceipt.Enabled = True
Command2.Enabled = True


End Sub

Private Sub date1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
sendkeys "{tab}"

End If
End Sub

Private Sub date1_LostFocus()
    If Trim(date1.text) <> "" Then
        If Not checkdate(Trim(date1.text), date1) Then
            date1.SetFocus
        End If
    End If
End Sub

Private Sub date2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    sendkeys "{tab}"
End If
End Sub

Private Sub date2_LostFocus()
    
    If Trim(date2.text) <> "" Then
        If Not checkdate(Trim(date2.text), date2) Then
            date2.SetFocus
        End If
    End If
End Sub
Private Sub Form_Load()
On Error GoTo aa1
Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> (Trim(UCase("MainMenu") Or Trim(UCase("frmbilllist"))))
    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
        Unload VB.Screen.ActiveForm
    End If
Loop

aa1:



Me.top = 0
Me.Left = 0
CheckCash = False
Set RS = New ADODB.Recordset
    RS.Open "select * from gledger where " & stringyear & " and cashbankbook=1", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        Do While Not RS.EOF
            COMBOGENLEDGER.AddItem Trim(RS!gledger)
            If Not RS.EOF Then
                   RS.MoveNext
            End If
        Loop
    End If
    
    If RS.State = 1 Then RS.Close
    RS.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly
    If RS.EOF = False Then
        CNSetup
        date1.text = RS!yarfrom
        date2.text = RS!yarto
    End If
    
    RS.Close
    
    
  Dim rs6 As New ADODB.Recordset
   If rs6.State = 1 Then rs6.Close
    rs6.Open "SELECT DISTINCT VOUCHERNUMBER,VoucherDate,Vouchertype  FROM VOUCHERS where  " & stringyear & " and genledger = '" + Trim("CASH-IN-HAND") + "' and  Vouchertype<>'J'", con, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rs6.RecordCount > 0 Then
       While Not rs6.EOF
            con.Execute "Update vouchers set cashcheck= 1 where  " & stringyear & " and VOUCHERNUMBER =" & rs6!VOUCHERNUMBER & " and   VoucherType =  '" & rs6!VoucherType & "' and  convert(smalldatetime,voucherDATE,103) =  convert(smalldatetime,'" & rs6!voucherDATE & "',103)"
            rs6.MoveNext
        Wend
  End If
  
   If rs6.State = 1 Then rs6.Close
    rs6.Open "SELECT DISTINCT VOUCHERNUMBER,VoucherDate,Vouchertype  FROM VOUCHERS where  " & stringyear & " and genledger = '" + Trim("CASH-IN-HAND") + "' and  Vouchertype='J'", con, adOpenForwardOnly, adLockReadOnly, adCmdText
    If rs6.RecordCount > 0 Then
       While Not rs6.EOF

          con.Execute "Update vouchers set cashcheck= 1 where  " & stringyear & " and VOUCHERNUMBER =" & rs6!VOUCHERNUMBER & " and   VoucherType =  '" & rs6!VoucherType & "' and   convert(smalldatetime,voucherDATE,103) =  convert(smalldatetime,'" & rs6!voucherDATE & "',103)"
          rs6.MoveNext
          
        Wend
  End If
    
BackColorFrom Me
  
End Sub
Sub TRYY()
Set rs1 = New ADODB.Recordset
    CON1.Execute ("delete * from treport where " & stringyear)
    rs1.Open "select * frm treport where " & stringyear, CON1, adOpenDynamic, adLockOptimistic, adCmdTable
    rs1.AddNew
    rs1!text = "** Opening Balance as on " + Trim(date1.text)
    rs1!ad = Balance
    rs1!Period = Trim(date1.text) + "  To  " + Trim(date2.text)
    rs1!header = "CASH BOOK"
  ' rs1!SUBLEDGER = Trim(Combosubledger.Text)
    rs1!dorc = "D"
    rs1.update
    RS.Open "select * from vouchers where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and VoucherDate>=cdate('" + Trim(date1.text) + "') and VoucherDate<=cdate('" + Trim(date2.text) + "')", con, adOpenDynamic, adLockReadOnly, adCmdText
    rs2.Open "select * from CASHC where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "'  and INVOICEDATE>=cdate('" + Trim(date1.text) + "') and INVOICEDATE<=cdate('" + Trim(date2.text) + "')", con, adOpenDynamic, adLockReadOnly, adCmdText
    
    If Not RS.BOF Then
        Do While Not RS.EOF 'And Not rs2.EOF
            'If rs!VoucherDate <= rs2!INVOICEDATE Then
                rs1.AddNew
                rs1!Genledger = RS!Genledger
                rs1!subledger = RS!subledger
                rs1!vdate = RS!voucherDATE
                rs1!vtype = RS!VoucherType
                rs1!vno = Trim(Str(RS!VOUCHERNUMBER))
                If Trim(RS!Description) <> "" Then
                    rs1!narration = Trim(RS!Description)
                Else
                    rs1!narration = " "
                End If
                 If Trim(UCase(RS!DebitorCredit)) = Trim(UCase("D")) Then
                    rs1!ad = RS!amount
                    rs1!dorc = "D"
                    Balance = Balance + RS!amount
                    rs1!Balance = Balance
                Else
                    rs1!aC = RS!amount
                    rs1!dorc = "C"
                    Balance = Balance - RS!amount
                    rs1!Balance = Balance
                End If
                If Trim(RS!CBND) <> "" Then
                    rs1!cbno = Trim(RS!CBND)
                Else
                    rs1!cbno = " "
                End If
                
                
                If Not RS.EOF Then
                    RS.MoveNext
                End If
            
            rs1.update
        Loop
    End If
    
    
End Sub


Sub ppp()

  rs1!text = "** Opening Balance as on " + Trim(date1.text)
    rs1!ad = Balance
    rs1!Period = Trim(date1.text) + "  To  " + Trim(date2.text)
    rs1!header = "CASH BOOK"
    rs1!dorc = "D"
    rs1.update
    
    RS.Open "SELECT * from VOUCHERS WHERE VOUCHERNUMBER IN (SELECT DISTINCT VOUCHERNUMBER, voucherdate  FROM VOUCHERS where  " & stringyear & " and genledger = '" + Trim(COMBOGENLEDGER.text) + "'  and   VoucherType ='R') and   genledger <> '" + Trim(COMBOGENLEDGER.text) + "' AND voucherDATE>=cdate('" + Trim(date1.text) + "') and voucherDATE<=cdate('" + Trim(date2.text) + "') and  VoucherType ='R'", con, adOpenDynamic, adLockReadOnly, adCmdText
    rs2.Open "select * from CASHC where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "'  and INVOICEDATE>=cdate('" + Trim(date1.text) + "') and INVOICEDATE<=cdate('" + Trim(date2.text) + "')", con, adOpenDynamic, adLockReadOnly, adCmdText
    rs3.Open "select * from casha where " & stringyear & " and INVOICEDATE>=cdate('" + Trim(date1.text) + "') and INVOICEDATE<=cdate('" + Trim(date2.text) + "')", con, adOpenDynamic, adLockReadOnly, adCmdText
    vflag = True
    
voucheroption:
    If Not RS.BOF Then
        Do While Not RS.EOF
            'If rs!VoucherDate <= rs2!INVOICEDATE Then
                rs1.AddNew
                rs1!Genledger = RS!Genledger & ""
                rs1!subledger = RS!subledger & ""
                rs1!vdate = RS!voucherDATE & ""
                rs1!vtype = RS!VoucherType & ""
                rs1!vno = Trim(Str(RS!VOUCHERNUMBER))
                If Trim(RS!Description) <> "" Then
                    rs1!narration = Trim(RS!Description)
                Else
                    rs1!narration = " "
                End If
                 If Trim(UCase(RS!DebitorCredit)) = Trim(UCase("D")) Then
                    rs1!ad = RS!amount
                    rs1!dorc = "D"
                    Balance = Balance + RS!amount
                    rs1!Balance = Balance
                Else
                    rs1!aC = RS!amount
                    rs1!dorc = "C"
                    Balance = Balance - RS!amount
                    rs1!Balance = Balance
                End If
                If Trim(RS!CBND) <> "" Then
                    rs1!cbno = Trim(RS!CBND)
                Else
                    rs1!cbno = " "
                End If
                
                If Not RS.EOF Then
                    RS.MoveNext
                End If
                rs1.update
        Loop
    End If
    If RS.State = 1 Then RS.Close
    If vflag = True Then
        RS.Open "SELECT * from VOUCHERS WHERE VOUCHERNUMBER IN (SELECT DISTINCT VOUCHERNUMBER  FROM VOUCHERS where  " & stringyear & " and genledger = '" + Trim(COMBOGENLEDGER.text) + "'   and   VoucherType ='P') and   genledger <> '" + Trim(COMBOGENLEDGER.text) + "'    AND voucherDATE>=cdate('" + Trim(date1.text) + "') and voucherDATE<=cdate('" + Trim(date2.text) + "') and  VoucherType ='P'", con, adOpenDynamic, adLockReadOnly, adCmdText
        If RS.RecordCount > 0 Then
            vflag = False
            GoTo voucheroption
        End If
    End If
    
    Dim rs5 As New ADODB.Recordset
    
    rs5.Open "SELECT DISTINCT VOUCHERNUMBER,VoucherDate,Vouchertype  FROM VOUCHERS where  " & stringyear & " and genledger = '" + Trim(COMBOGENLEDGER.text) + "' and   voucherDATE>=cdate('" + Trim(date1.text) + "') and voucherDATE<=cdate('" + Trim(date2.text) + "') and  VoucherType ='J'", con, adOpenDynamic, adLockReadOnly, adCmdText
    If rs5.RecordCount > 0 Then
       Set rs4 = New ADODB.Recordset
       rs4.Open "SELECT * from VOUCHERS WHERE " & stringyear & " and VOUCHERNUMBER =" & rs5!VOUCHERNUMBER & " and  genledger <> '" + Trim(COMBOGENLEDGER.text) + "' and   VoucherType =  '" & rs5!VoucherType & "' and   voucherDATE =  cdate('" & rs5!voucherDATE & "')", con, adOpenDynamic, adLockReadOnly, adCmdText
       If Not rs4.BOF Then
          Do While Not rs4.EOF
                'If rs!VoucherDate <= rs2!INVOICEDATE Then
                    rs1.AddNew
                    rs1!Genledger = rs4!Genledger & ""
                    rs1!subledger = rs4!subledger & ""
                    rs1!vdate = rs4!voucherDATE & ""
                    rs1!vtype = rs4!VoucherType & ""
                    rs1!vno = Trim(Str(rs4!VOUCHERNUMBER))
                    If Trim(rs4!Description) <> "" Then
                        rs1!narration = Trim(rs4!Description)
                    Else
                        rs1!narration = " "
                    End If
                     If Trim(UCase(rs4!DebitorCredit)) = Trim(UCase("D")) Then
                        rs1!ad = rs4!amount
                        rs1!dorc = "D"
                        Balance = Balance + rs4!amount
                        rs1!Balance = Balance
                    Else
                        rs1!aC = rs4!amount
                        rs1!dorc = "C"
                        Balance = Balance - rs4!amount
                        rs1!Balance = Balance
                    End If
                    If Trim(rs4!CBND) <> "" Then
                        rs1!cbno = Trim(rs4!CBND)
                    Else
                        rs1!cbno = " "
                    End If
                
                    If Not rs4.EOF Then
                        rs4.MoveNext
                    End If
                    rs1.update
            Loop
        End If
    End If
    
    
    '**************** FOR  COUNTER SALE cashc
    If Not rs2.EOF Then
        Do While Not rs2.EOF
            rs1.AddNew
           If rs2!amount > 0 Then
            rs1!Genledger = rs2!Genledger & ""
            rs1!subledger = rs2!subledger & ""
            rs1!vdate = rs2!invoiceDate & ""
            rs1!vtype = Trim("S")
            rs1!vno = Trim(rs2!invoiceNo)
            rs1!aC = rs2!amount
            rs1!dorc = "C"
             rs1!narration = "Sales C/M"
            Balance = Balance - rs2!amount
            rs1!Balance = Balance
            rs1.update
           End If
           If Not rs2.EOF Then
               rs2.MoveNext
           End If
  
           
        Loop
    End If
        '**************** FOR  COUNTER SALE casha
' only in the case of Cash in hand
If UCase(COMBOGENLEDGER) = "CASH-IN-HAND" Then
    If Not rs3.EOF Then
       Do While Not rs3.EOF
            If rs3!baa > 0 Then
            rs1.AddNew
            
            rs1!Genledger = "CASH-IN-HAND"
            rs1!subledger = rs3!CASHPARTYNAME & ""
            rs1!vdate = rs3!invoiceDate & ""
            rs1!vtype = Trim("S")
            rs1!vno = Trim(rs3!invoiceNo)
            rs1!aC = rs3!baa
            rs1!narration = "Sales C/M"
            rs1!dorc = "C"
            Balance = Balance + rs3!baa
            rs1!Balance = Balance
            rs1.update
          End If
          If Not rs3.EOF Then
               rs3.MoveNext
          End If
            
      Loop
   End If
End If

    
    
    
    If rs1.State = 1 Then
        rs1.Close
    End If
    
    If RS.State = 1 Then
        RS.Close
    End If
   
    viewcash.GenrepNoFooter
    PrintOption.Show
   
   
   'viewcash.Top = 0
   'viewcash.Left = 0
   'viewcash.Show

End Sub
