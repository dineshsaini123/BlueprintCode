VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmSelectedParty 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4335
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSelectedParty.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbodist 
      Height          =   315
      Left            =   2940
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4665
   End
   Begin VB.ComboBox cboAgent 
      Height          =   315
      Left            =   2940
      TabIndex        =   1
      Top             =   555
      Width           =   2580
   End
   Begin VB.ListBox Combosubledger 
      Appearance      =   0  'Flat
      Height          =   1830
      Left            =   2955
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   930
      Width           =   4665
   End
   Begin VB.TextBox Alpha 
      Height          =   315
      Left            =   6690
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2820
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton Commandreturn 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Return"
      Height          =   405
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3465
      Width           =   1545
   End
   Begin VB.CommandButton Commandshow 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Show"
      Height          =   405
      Left            =   2955
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3465
      Width           =   1545
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   2940
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   105
      Visible         =   0   'False
      Width           =   4665
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
      Left            =   2970
      TabIndex        =   3
      Top             =   2835
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
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
      Left            =   4785
      TabIndex        =   4
      Top             =   2835
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Name"
      Height          =   285
      Left            =   840
      TabIndex        =   14
      Top             =   555
      Width           =   1980
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From The Date"
      Height          =   315
      Left            =   840
      TabIndex        =   13
      Top             =   2865
      Width           =   1995
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " - To - "
      Height          =   315
      Left            =   4185
      TabIndex        =   12
      Top             =   2850
      Width           =   585
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Alphabat"
      Height          =   285
      Left            =   6000
      TabIndex        =   11
      Top             =   2850
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub. Ledger Desc."
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Top             =   930
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Station"
      Height          =   285
      Left            =   840
      TabIndex        =   9
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmSelectedParty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As Recordset
    Dim K As Integer
    Dim arr() As String
    Dim cust As String
Function rsets(ST As String, length As Integer) As String
      Dim kk As String
            kk = Trim(ST)
            If Len(kk) < length Then
                Do While Not Len(kk) = length
                    kk = " " + kk
                Loop
            End If
            If Len(kk) > length Then
                Do While Not Len(kk) = length
                    kk = Mid$(kk, 0, Len(kk) - 1)
                Loop
            End If
        rsets = kk
End Function
Private Sub alpha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub
Private Sub cboagent_Click()
  If RS.State = 1 Then
        RS.close
    End If
    RS.Open "select distinct(subledger) from invoicea where  " & stringyear & " and AgentName='" + Trim(cboagent.Text) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
    Combosubledger.Clear
    If Not RS.BOF Then
        Do While Not RS.EOF
            Combosubledger.AddItem Trim(RS!SUBLEDGER)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
End Sub

Private Sub cbodist_Click()
  If RS.State = 1 Then
        RS.close
    End If
    RS.Open "select distinct(subledger) from invoicea where  " & stringyear & " and district='" + Trim(cboDist.Text) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
    Combosubledger.Clear
    If Not RS.BOF Then
        Do While Not RS.EOF
            Combosubledger.AddItem Trim(RS!SUBLEDGER)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
End Sub

Private Sub COMBOGENLEDGER_Change()
'    If RS.State = 1 Then
'        RS.Close
'    End If
'    RS.Open "select * from sledger where  " &  stringyear & " and DISTCODE='" + Trim(COMBOGENLEDGER.Text) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
'    Combosubledger.Clear
'    If Not RS.BOF Then
'        Do While Not RS.EOF
'            Combosubledger.AddItem Trim(RS!subledger)
'            If Not RS.EOF Then
'                RS.MoveNext
'            End If
'        Loop
'    End If
'    RS.Close
    
End Sub

Private Sub COMBOGENLEDGER_Click()
'If RS.State = 1 Then
'        RS.Close
'    End If
'    RS.Open "select * from sledger where  " &  stringyear & " and DISTCODE='" + Trim(COMBOGENLEDGER.Text) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
'    Combosubledger.Clear
'    If Not RS.BOF Then
'        Do While Not RS.EOF
'            Combosubledger.AddItem Trim(RS!subledger)
'            If Not RS.EOF Then
'                RS.MoveNext
'            End If
'        Loop
'    End If
'    RS.Close
End Sub

Private Sub COMBOGENLEDGER_KeyPress(KeyAscii As Integer)
If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
   SendKeys "{tab}"
   Exit Sub
End If
If KeyAscii = 13 Then
   SendKeys "{DOWN}"
   SendKeys "{tab}"
End If

End Sub

Private Sub COMBOGENLEDGER_LostFocus()
'    If Trim(COMBOGENLEDGER.Text) <> "" Then
'        RS.Open "select * from gledger where  " &  stringyear & " and slf=1", CON, adOpenStatic, adLockReadOnly, adCmdText
'        If Not RS.BOF Then
'            RS.Find "gledger='" + Trim(COMBOGENLEDGER.Text) + "'"
'            If RS.EOF Then
'                COMBOGENLEDGER.SetFocus
'            End If
'        Else
'            COMBOGENLEDGER.SetFocus
'        End If
'        RS.Close
'    End If
End Sub

Private Sub Combosubledger_GotFocus()
    If Trim(COMBOGENLEDGER.Text) = "" Then
        COMBOGENLEDGER.SetFocus
    End If
   '' Me.KeyPreview = False
End Sub
Private Sub Combosubledger_KeyPress(KeyAscii As Integer)
'If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
'   SendKeys "{tab}"
'   Exit Sub
'End If

'If KeyAscii = 13 Then
'   SendKeys "{Down}"
'   SendKeys "{tab}"
'End If
End Sub

Private Sub Combosubledger_LostFocus()
If Trim(Combosubledger.Text) <> "" Then
    If Trim(COMBOGENLEDGER.Text) <> "" Then
        If RS.State = 1 Then
            RS.close
        End If
        RS.Open "select * from sledger where  " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
        If RS.BOF Then
            Combosubledger.SetFocus
        End If
        RS.close
    Else
        Combosubledger.Text = ""
       '' Me.KeyPreview = True
    End If
End If

End Sub
 Sub ALPHAB()
    If RS.State = 1 Then RS.close
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    'con.Execute ("DELETE from treport where  " &  stringyear & "")
    If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
        MsgBox "invalid date"
        Exit Sub
    End If
    Dim rs1 As New ADODB.Recordset
    Dim Balance As Double
    Dim OPBALANCE As Double
    Dim SDamount As Double
    Dim SCamount As Double
    Dim RsT As New ADODB.Recordset
    Dim viewsubledger As Boolean
    viewsubledger = False
    Balance = 0
    OPBALANCE = 0
    Dim tempdate As String
    tempdate = date1.Text
    If s1 <> 1 Then
    date1.Text = date2.Text
    End If
    
    OPENINGSUBLEDGERS
    
    If s1 <> 1 Then
    date1.Text = tempdate
    End If
    DoEvents
    If s1 = 1 Then
    If Trim(alpha.Text) <> "" And alpha.Visible = True Then
      ' vouchers creditors
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(alpha.Text) + "%'  AND   VOUCHERS.DebitorCredit='C' and convert(smalldatetime,voucherdate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
      ' vouchers debtors
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where   " & stringyear & " and  genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(alpha.Text) + "%' AND  VOUCHERS.DebitorCredit='D' and  convert(smalldatetime,voucherdate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)      ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
      ' invoice
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear ,setupid)  SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & " FROM INVOICEA  where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(alpha.Text) + "%'  and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
      ' cash credit
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(alpha.Text) + "%'   and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)  "
      ' cash debit
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear ,setupid)   SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where  " & stringyear & " and CASHA.BAA<>0 and  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(alpha.Text) + "%'   and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
       ' credit a
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear ,setupid) SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & " FROM CREDITA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(alpha.Text) + "%' and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)   "
       ' dnfadr
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, '', " & UId & ",'" & main.session & "'," & main.setupid & "   From DNFA  where    " & stringyear & " and Pgld ='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld like '" + Trim(alpha.Text) + "%'  and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
       'cnf1cr
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '', " & UId & ",'" & main.session & "'," & main.setupid & "  From CNF1A where  " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld like '" + Trim(alpha.Text) + "%'  and convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
 
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From DNFB  where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld like '" + Trim(alpha.Text) + "%'   and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '', " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1B where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld like '" + Trim(alpha.Text) + "%'   and convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
   End If
   
   
'    If Trim(Alpha.Text) = "" And Alpha.Visible = True Then
'
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & " From VOUCHERS Where    " &  stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "'  AND   VOUCHERS.DebitorCredit='C' and  convert(smalldatetime,voucherdate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & " From VOUCHERS Where    " &  stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
'
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid) SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3  , " & UId & " FROM INVOICEA  where    " &  stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "'  and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
'
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid) SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Export Sales' , CASHA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & " FROM CASHA   where    " &  stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "'    and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)   "
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Export Sales' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & "  FROM CASHA  where   " &  stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "'   and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND CASHA.BAA <>0  "
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3, " & UId & "  FROM CREDITA  where    " &  stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "'    and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
'
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, ''  , " & UId & " From DNFA  where    " &  stringyear & " and Pgld ='" + Trim(COMBOGENLEDGER.Text) + "'and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
'
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '' , " & UId & " From CNF1A where  " &  stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.Text) + "'  and convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
'
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, '', " & UId & "  From DNFB  where  " &  stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "'   and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & " From CNF1B where  " &  stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103) ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
'End If

If Trim(alpha.Text) = "" And cust <> "" Then
''' Code Multiple Party
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & " From VOUCHERS Where    " &  stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "'  AND   VOUCHERS.DebitorCredit='C' and  convert(smalldatetime,voucherdate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )   SELECT GenLedger, SubLedger, VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & " From VOUCHERS Where    " &  stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & " FROM INVOICEA  where    " &  stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "'  and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Export Sales' , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & "  FROM CASHA  where    " &  stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "'   and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)   "
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Export Sales' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & "  FROM CASHA  where    " &  stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "'   and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND CASHA.BAA <>0  "
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )    SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & " FROM CREDITA  where    " &  stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "'    and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )   SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, '' , " & UId & "   From DNFA  where    " &  stringyear & " and Pgld ='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld = '" & Combosubledger.Text & "' and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid )   SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '' , " & UId & " From CNF1A where  " &  stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld = '" & Combosubledger.Text & "'    and convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid )   SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, ''     , " & UId & " From DNFB  where  " &  stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld = '" & Combosubledger.Text & "'   and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
'                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid )   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & " From CNF1B where  " &  stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld = '" & Combosubledger.Text & "' and  convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103) ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"

For I = 0 To K - 1
cust = arr(I)
If arr(I) <> "" Then
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid  )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where   " & stringyear & " and  genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'  AND   VOUCHERS.DebitorCredit='C' and  convert(smalldatetime,voucherdate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid  )   SELECT GenLedger, SubLedger, VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid  )  SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM INVOICEA  where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'  and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'   and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)   "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid  )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'   and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND CASHA.BAA <>0  "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid )    SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM CREDITA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'    and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid )   SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & "   From DNFA  where    " & stringyear & " and Pgld ='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld = '" & cust & "' and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid,fyear,setupid  )   SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1A where  " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld = '" & cust & "'    and convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid,fyear,setupid  )   SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, ''     , " & UId & ",'" & main.session & "'," & main.setupid & " From DNFB  where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld = '" & cust & "'   and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid ,fyear,setupid )   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1B where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld = '" & cust & "' and  convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103) ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
End If
Next
cust = "a"
End If

If Trim(alpha.Text) = "" And cust = "" Then

con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' AND   VOUCHERS.DebitorCredit='C' and  convert(smalldatetime,voucherdate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)  and convert(smalldatetime,voucherdate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )   SELECT GenLedger, SubLedger, VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )  SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM INVOICEA  where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid  )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)   "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "'  and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND CASHA.BAA <>0  "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid  )    SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM CREDITA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )   SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & "   From DNFA where   " & stringyear & " and   Pgld ='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid,fyear,setupid  )   SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1A where  " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.Text) + "'  and convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid,fyear,setupid  )   SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, ''     , " & UId & ",'" & main.session & "'," & main.setupid & " From DNFB  where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "'  and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid,fyear,setupid  )   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1B where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103) <convert(smalldatetime,'" + Trim(date2.Text) + "',103) ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
End If
End If
con.Execute "insert into Treport ( Genledger,Subledger,openingbalance,userid,fyear,setupid) SELECT '" + Trim(COMBOGENLEDGER.Text) + "' as genled, SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId,'" & main.session & "'," & main.setupid & " from subledgertrail where  " & stringyear & " GROUP BY SUBLEDGER;"

Dim sum, cr, dr As Double
sum = 0
dr = 0
cr = 0



main.reportname = "Sub. Ledger"

If SLEDGERPRINT.alpha.Visible = True Then
   viewledger.SelectedParty
Else
   viewledger.SelectedParty
End If


If s1 = 1 Then
PrintOption.Show
Else
MainMenu.cr1.Connect = constr
MainMenu.cr1.WindowState = crptMaximized
MainMenu.cr1.SelectionFormula = "{winrpt.fyear}='" & main.session & "' and {winrpt.setupid}=" & main.setupid & " and {winrpt.uid}=" & main.UId
MainMenu.cr1.ReportFileName = rptPath & "\SLClosingAccount.rpt"
MainMenu.cr1.Formulas(0) = "fromdate='" & frmSelectedParty.date1.Text & "'"
MainMenu.cr1.Formulas(1) = "todate='" & frmSelectedParty.date2.Text & "'"
MainMenu.cr1.WindowShowPrintBtn = True
MainMenu.cr1.WindowShowPrintSetupBtn = True
MainMenu.cr1.WindowState = crptMaximized
MainMenu.cr1.Action = 1
End If
End Sub

Private Sub CommandReturn_Click()
    '''MainMenu.Toolbar1.Visible = False
    Unload Me
End Sub
Private Sub Commandshow_Click()

con.Execute "DELETE from treport where  " & stringyear & ""
con.Execute "DELETE from subledgertrail where  " & stringyear & ""
cust = ""
For I = 0 To Me.Combosubledger.ListCount - 1
If Me.Combosubledger.Selected(I) = True Then
cust = Me.Combosubledger.List(I)
GoTo abc:
End If
Next
abc:

 Commandshow.Enabled = False
'********sub for alpha wise and Partywise according to new fast mathed
 DoEvents
 
 con.Execute "DELETE from subledgertrail where  " & stringyear & ""
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 ALPHAB
 Commandshow.Enabled = True
 
 
End Sub

Private Sub date1_LostFocus()
    If Trim(date1.Text) <> "" Then
        If Not checkdate(Trim(date1.Text), date1) Then
            date1.SetFocus
            End If
    End If
End Sub

Private Sub date2_LostFocus()
    If Trim(date2.Text) <> "" Then
        If Not checkdate(Trim(date2.Text), date2) Then
            date2.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{TAB}"
End If

End Sub

Private Sub Form_Load()


If s1 = 1 Then
date1.Visible = True
date2.Visible = True
Label2.Visible = True
Label3.Visible = True
Else
date1.Visible = False
date2.Visible = False
Label2.Visible = False
Label3.Visible = False
End If
  


On Error GoTo ac1

Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> (Trim(UCase("MainMenu") Or Trim(UCase("frmbilllist"))))
    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
        Unload VB.Screen.ActiveForm
    End If
Loop

ac1:



COMBOGENLEDGER.Text = "SUNDRY DEBTORS"

Me.Top = 10
Me.Left = 10
'Set CON = New ADODB.Connection
Set RS = New ADODB.Recordset
'    With CON
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\data.mdb"
'        .Open
'    End With
    RS.Open "select distinct(district) from invoicea where  " & stringyear & "", con, adOpenStatic, adLockReadOnly
    
    If Not RS.BOF Then
        Do While Not RS.EOF
            cboDist.AddItem Trim(RS(0))
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    RS.Open "Select * from setup1 where " & stringyear & "", con, adOpenStatic, adLockReadOnly
    CNSetup
    date1.Text = RS!yarfrom
    date2.Text = RS!yarto
    RS.close
    
    
    If RS.State = 1 Then RS.close
    RS.Open "select distinct(AgentName) from invoicea where  " & stringyear & " order by AGENTNAME", con
    While RS.EOF = False
    cboagent.AddItem RS(0)
    RS.MoveNext
    Wend
    
    
End Sub
Sub xx()
End Sub
Sub OPENINGSUBLEDGERS()
Dim k1 As Integer

If Trim(alpha.Text) <> "" Then
        'CON.Execute "Insert into subledgertrail  SELECT SUBLEDGER , YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId FROM SLEDGER where   " &  stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' AND subledger like '" + Trim(alpha.Text) + "%'", p, adCmdText
        'subledger opening start
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER ,YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER  where   " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "'  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'"
    ' from invoice a
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.fyear = INVOICEA.fyear) AND (SLEDGER.setupid = INVOICEA.setupid))" _
        & " where  sledger.fyear='" & main.session & "' and invoicea.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER "
   ' from casha
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT,  " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER)  AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid)) " _
        & " where  sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'  and convert(smalldatetime,invoicedate,103)< convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; "
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,0 AS YEAROPENING,0 AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT ," & UId & " as UserId,'" & main.session & "'," & main.setupid & "  " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))" _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'   and convert(smalldatetime,invoicedate,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103)  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; "
        
    ' from credita
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING,0 AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.fyear = credita.fyear) AND (SLEDGER.setupid = credita.setupid)) " _
        & " where sledger.fyear='" & main.session & "' and credita.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; "
        
        
   ' from vouchers
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT ,  " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger)  AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid)" _
        & " WHERE sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'D' and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103)  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;"
        
        
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'C' and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; "
         ''''ok
  'from cnf1a
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD)  AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid)" _
        & " WHERE sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='D' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
                
                con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING,0 AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD)  AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid)" _
        & " WHERE  sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='C' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
     
        
  ' from dnfa
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)  AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid)" _
        & " WHERE sledger.fyear='" & main.session & "' and dnfa.setupid=" & main.setupid & " and DNFA.DC = 'D' and  pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING,0 AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)  AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid)" _
        & " WHERE  sledger.fyear='" & main.session & "' and dnfa.setupid=" & main.setupid & " and DNFA.DC='C' and   pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,('" + Trim(date1.Text) + "'),103) AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
   ' from cnf1b
        con.Execute "Insert into subledgertrail     SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid) " _
        & " WHERE  sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC='D' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD)  AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid)" _
        & " WHERE  sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC= 'C' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.SUBLEDGER like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
   ' dnfb
        con.Execute "Insert into subledgertrail     SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD)  AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid)" _
        & " WHERE  sledger.fyear='" & main.session & "' and dnfb.setupid=" & main.setupid & " and DNFB.DC='D' and   gld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103)  AND SLEDGER.SUBLEDGER like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD) AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid)" _
        & " WHERE  sledger.fyear='" & main.session & "' and dnfb.setupid=" & main.setupid & " and DNFB.DC= 'C' and  gld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.SUBLEDGER like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
        
        'CON.Execute "DELETE from TemprptTrialBalance where  " &  stringyear & ""
        'CON.Execute "insert into TemprptTrialBalance ( Subledger,openingbalance,userid ) SELECT  SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId from subledgertrail where  " &  stringyear & " GROUP BY SUBLEDGER;"
    
  End If
          
          

'Sub Ledger Account Code

If Trim(alpha.Text) = "" And cust = "" Then
 K = 0
   
    k1 = 0
    'cust = ""
    ReDim Preserve arr(25)
    arr(K) = ""
    'ReDim arr(10)
    For I = 0 To Me.Combosubledger.ListCount - 1
       arr(k1) = Me.Combosubledger.List(I)
       K = K + 1
       k1 = k1 + 1
    Next
    
    For I = 0 To K - 1
    cust = arr(I)
    If arr(I) <> "" Then






 con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,SLEDGER.YEAROPENING,0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT," & UId & " as UserId,'" & main.session & "'," & main.setupid & "  FROM SLEDGER where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER = '" & cust & "'", p, adCmdText
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,0 as YEAROPENING,sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT," & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER and SLEDGER.fyear = INVOICEA.fyear and SLEDGER.setupid = INVOICEA.setupid) )" _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER=  '" & cust & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER "
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,0 AS YEAROPENING,Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT ," & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER and SLEDGER.fyear = cashA.fyear and SLEDGER.setupid = casha.setupid)) " _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER=  '" & cust & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,0 AS YEAROPENING,0 AS OPAMOUNTDEBIT,Sum(CASHA.BAA) AS OPAMOUNTCREDIT, " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER and SLEDGER.fyear = cashA.fyear and SLEDGER.setupid = casha.setupid))" _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER ='" & cust & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,0 AS YEAROPENING,0 AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT) AS OPAMOUNTCREDIT," & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER and SLEDGER.fyear = credita.fyear and SLEDGER.setupid = credita.setupid)) " _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER='" & cust & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
        
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT," & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid) " _
        & " WHERE  sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER= '" & cust & "' and convert(smalldatetime,voucherdate,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = VOUCHERS.fyear) AND (SLEDGER.setupid = VOUCHERS.setupid) " _
        & " WHERE  sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER= '" & cust & "' and convert(smalldatetime,voucherdate,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
         ''''ok
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD)  AND (SLEDGER.fyear = CNF1A.fyear) AND (SLEDGER.setupid = CNF1A.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and CNF1A.DC='D' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and  CNF1A.PSLD = '" & cust & "'  and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING,0 AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD)  AND (SLEDGER.fyear = CNF1A.fyear) AND (SLEDGER.setupid = CNF1A.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and CNF1A.DC='C' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and CNF1A.PSLD = '" & cust & "'  and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = dnfa.PSLD)  AND (SLEDGER.fyear = dNFA.fyear) AND (SLEDGER.setupid = dNFA.setupid) " _
        & " WHERE  sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DNFA.DC = 'D' and  pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and   DNFA.PSLD = '" & cust & "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING,0 AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = dnfa.PSLD)  AND (SLEDGER.fyear = dNFA.fyear) AND (SLEDGER.setupid = dNFA.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DNFA.DC='C' and   pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and DNFA.PSLD = '" & cust & "'   and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1b.SLD)  AND (SLEDGER.fyear = CNF1b.fyear) AND (SLEDGER.setupid = CNF1b.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and CNF1B.DC='D' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and CNF1B.SLD = '" & cust & "'    and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1b.SLD)  AND (SLEDGER.fyear = CNF1b.fyear) AND (SLEDGER.setupid = CNF1b.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and CNF1B.DC= 'C' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and CNF1B.SLD = '" & cust & "'      and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = dnfb.SLD)  AND (SLEDGER.fyear = dNFb.fyear) AND (SLEDGER.setupid = dNFb.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DNFB.DC='D' and   gld = '" + Trim(COMBOGENLEDGER.Text) + "' and DNFB.SLD = '" & cust & "'     And convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = dnfb.SLD)  AND (SLEDGER.fyear = dNFb.fyear) AND (SLEDGER.setupid = dNFb.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DNFB.DC= 'C' and  gld = '" + Trim(COMBOGENLEDGER.Text) + "' and DNFB.SLD = '" & cust & "'    and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
        
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , YEAROPENING,  0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
'        & " FROM SLEDGER where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "'", p, adCmdText
'
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
'        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER and SLEDGER.fyear = INVOICEA.fyear and SLEDGER.setupid = INVOICEA.setupid))  " _
'        & " where   invoicea.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "'and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'        & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
'
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
'        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER and SLEDGER.fyear = cashA.fyear and SLEDGER.setupid = casha.setupid)) " _
'        & " where casha.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and  SLEDGER.SUBLEDGER <>'CASH PARTY'    and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
'
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,0 AS YEAROPENING ,0 AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
'        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER and SLEDGER.fyear = cashA.fyear and SLEDGER.setupid = casha.setupid))" _
'        & " where   casha.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
'
'
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING,0 AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
'        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER and SLEDGER.fyear = credita.fyear and SLEDGER.setupid = credita.setupid)) " _
'        & " where credita.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
'
'        '==============================================================
'
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
'        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid) " _
'        & " WHERE  sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
'
'
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
'        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger)  AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid)" _
'        & " WHERE  sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
'         ''''ok
'
'        '==============================================================
'
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
'        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid) " _
'        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and CNF1A.DC='D' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
'
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING,0 AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
'        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid) " _
'        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and CNF1A.DC='C' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
'
'
'
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
'        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid) " _
'        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DNFA.DC = 'D' and  pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
'
'
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING,0 AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
'        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid)" _
'        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DNFA.DC='C' and   pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
'
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
'        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid) " _
'        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and CNF1B.DC='D' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
'
'
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
'        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD)  AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid)" _
'        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and CNF1B.DC= 'C' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
'
'
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
'        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD)  AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid)" _
'        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DNFB.DC='D' and   gld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
'
'        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
'        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD) AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid)" _
'        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DNFB.DC= 'C' and  gld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
       End If


cust = ""



        Next I
  End If
  
If Trim(alpha.Text) = "" And cust <> "" Then
    K = 0
    k1 = 0
    'cust = ""
    ReDim Preserve arr(25)
    arr(K) = ""
    'ReDim arr(10)
    For I = 0 To Me.Combosubledger.ListCount - 1
    If Me.Combosubledger.Selected(I) = True Then
       arr(k1) = Me.Combosubledger.List(I)
       K = K + 1
       k1 = k1 + 1
    End If
    Next
    
    For I = 0 To K - 1
    cust = arr(I)
    If arr(I) <> "" Then
    
    If Me.Combosubledger.Text <> "" Then
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,SLEDGER.YEAROPENING,0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT," & UId & " as UserId,'" & main.session & "'," & main.setupid & "  FROM SLEDGER where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER = '" & cust & "'", p, adCmdText
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,0 as YEAROPENING,sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT," & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER and SLEDGER.fyear = INVOICEA.fyear and SLEDGER.setupid = INVOICEA.setupid) )" _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER=  '" & cust & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER "
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,0 AS YEAROPENING,Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT ," & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER and SLEDGER.fyear = cashA.fyear and SLEDGER.setupid = casha.setupid)) " _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER=  '" & cust & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,0 AS YEAROPENING,0 AS OPAMOUNTDEBIT,Sum(CASHA.BAA) AS OPAMOUNTCREDIT, " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER and SLEDGER.fyear = cashA.fyear and SLEDGER.setupid = casha.setupid))" _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER ='" & cust & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,0 AS YEAROPENING,0 AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT) AS OPAMOUNTCREDIT," & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER and SLEDGER.fyear = credita.fyear and SLEDGER.setupid = credita.setupid)) " _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER='" & cust & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
        
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT," & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid) " _
        & " WHERE  sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER= '" & cust & "' and convert(smalldatetime,voucherdate,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = VOUCHERS.fyear) AND (SLEDGER.setupid = VOUCHERS.setupid) " _
        & " WHERE  sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER= '" & cust & "' and convert(smalldatetime,voucherdate,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
         ''''ok
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD)  AND (SLEDGER.fyear = CNF1A.fyear) AND (SLEDGER.setupid = CNF1A.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and CNF1A.DC='D' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and  CNF1A.PSLD = '" & cust & "'  and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING,0 AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD)  AND (SLEDGER.fyear = CNF1A.fyear) AND (SLEDGER.setupid = CNF1A.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and CNF1A.DC='C' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and CNF1A.PSLD = '" & cust & "'  and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = dnfa.PSLD)  AND (SLEDGER.fyear = dNFA.fyear) AND (SLEDGER.setupid = dNFA.setupid) " _
        & " WHERE  sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DNFA.DC = 'D' and  pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and   DNFA.PSLD = '" & cust & "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING,0 AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = dnfa.PSLD)  AND (SLEDGER.fyear = dNFA.fyear) AND (SLEDGER.setupid = dNFA.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DNFA.DC='C' and   pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and DNFA.PSLD = '" & cust & "'   and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1b.SLD)  AND (SLEDGER.fyear = CNF1b.fyear) AND (SLEDGER.setupid = CNF1b.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and CNF1B.DC='D' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and CNF1B.SLD = '" & cust & "'    and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1b.SLD)  AND (SLEDGER.fyear = CNF1b.fyear) AND (SLEDGER.setupid = CNF1b.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and CNF1B.DC= 'C' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and CNF1B.SLD = '" & cust & "'      and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = dnfb.SLD)  AND (SLEDGER.fyear = dNFb.fyear) AND (SLEDGER.setupid = dNFb.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DNFB.DC='D' and   gld = '" + Trim(COMBOGENLEDGER.Text) + "' and DNFB.SLD = '" & cust & "'     And convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = dnfb.SLD)  AND (SLEDGER.fyear = dNFb.fyear) AND (SLEDGER.setupid = dNFb.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and DNFB.DC= 'C' and  gld = '" + Trim(COMBOGENLEDGER.Text) + "' and DNFB.SLD = '" & cust & "'    and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
  End If
  End If
  Next I
  
  cust = "a"
  End If
  
  
  
End Sub






