VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form SLEDGERPRINT 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5115
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   11190
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "SLPRINT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      BackColor       =   &H0078CFE9&
      Caption         =   "All"
      Height          =   195
      Left            =   7065
      TabIndex        =   14
      Top             =   1215
      Width           =   1305
   End
   Begin VB.ListBox Combosubledger 
      Appearance      =   0  'Flat
      Height          =   4080
      Left            =   1755
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   720
      Width           =   4665
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0078CFE9&
      Caption         =   "Dos Report"
      Height          =   360
      Left            =   9075
      TabIndex        =   12
      Top             =   1695
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0078CFE9&
      Caption         =   "Crystal Report"
      Height          =   315
      Left            =   7095
      TabIndex        =   11
      Top             =   1740
      Value           =   -1  'True
      Width           =   1770
   End
   Begin VB.TextBox Alpha 
      Height          =   315
      Left            =   1755
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CommandButton Commandreturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   675
      Left            =   8835
      Picture         =   "SLPRINT.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2220
      Width           =   1545
   End
   Begin VB.CommandButton Commandshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Show"
      Height          =   675
      Left            =   7080
      Picture         =   "SLPRINT.frx":0BF0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2220
      Width           =   1545
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   1755
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   330
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
      Left            =   8190
      TabIndex        =   2
      Top             =   675
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
      Left            =   9945
      TabIndex        =   3
      Top             =   675
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From The Date"
      Height          =   315
      Left            =   7035
      TabIndex        =   10
      Top             =   705
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " - To - "
      Height          =   315
      Left            =   9405
      TabIndex        =   9
      Top             =   675
      Width           =   585
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Alphabat"
      Height          =   285
      Left            =   195
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub. Ledger Desc."
      Height          =   285
      Left            =   195
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gen. Ledger Desc."
      Height          =   285
      Left            =   195
      TabIndex        =   6
      Top             =   330
      Width           =   2055
   End
End
Attribute VB_Name = "SLEDGERPRINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As Connection
 Dim K As Integer
    Dim arr() As String
        Dim cust As String
Dim RS As Recordset

Private Sub Check1_Click()

For K = 0 To Combosubledger.ListCount - 1
   Combosubledger.Selected(K) = False
Next

If Check1.value = 1 Then
For K = 0 To Combosubledger.ListCount - 1
    Combosubledger.Selected(K) = True
Next
   
End If

End Sub

Private Sub COMBOGENLEDGER_Change()
    If RS.State = 1 Then
        RS.close
    End If
    RS.Open "select * from sledger where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
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

Private Sub COMBOGENLEDGER_Click()
If RS.State = 1 Then
        RS.close
    End If
    RS.Open "select * from sledger where  " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
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

Private Sub COMBOGENLEDGER_LostFocus()
    If Trim(COMBOGENLEDGER.Text) <> "" Then
        If RS.State = 1 Then RS.close
        RS.Open "select * from gledger where  " & stringyear & " and slf=1", CON, adOpenStatic, adLockReadOnly, adCmdText
        If Not RS.BOF Then
            RS.Find "gledger='" + Trim(COMBOGENLEDGER.Text) + "'"
            If RS.EOF Then
                COMBOGENLEDGER.SetFocus
            End If
        Else
            COMBOGENLEDGER.SetFocus
        End If
        RS.close
    End If
End Sub

Private Sub Combosubledger_GotFocus()
    If Trim(COMBOGENLEDGER.Text) = "" Then
        COMBOGENLEDGER.SetFocus
    End If
End Sub

Private Sub Combosubledger_LostFocus()
On Error Resume Next
If Trim(Combosubledger.Text) <> "" Then
    If Trim(COMBOGENLEDGER.Text) <> "" Then
        If RS.State = 1 Then
            RS.close
        End If
        RS.Open "select * from sledger where  " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
        If RS.BOF Then
            Combosubledger.SetFocus
        End If
        RS.close
    Else
        Combosubledger.Text = ""
    End If
End If
End Sub

 Sub ALPHAB()
    If RS.State = 1 Then RS.close
    DoEvents
    CON.Execute ("DELETE from treport where  " & stringyear & "")
    
    If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
        MsgBox "invalid date", , title1
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
    OPENINGSUBLEDGERS
    DoEvents
    
''    If Trim(Alpha.Text) <> "" And Alpha.Visible = True Then
''
''      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)   SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(Alpha.Text) + "%' AND  VOUCHERS.DebitorCredit='D' and  convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)      ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
''      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )    SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(Alpha.Text) + "%'  AND   VOUCHERS.DebitorCredit='C' and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
''
''      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )  SELECT  INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & " FROM INVOICEA  where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(Alpha.Text) + "%'  and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
''      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )   SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Export Sales' , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(Alpha.Text) + "%'   and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  "
''      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )   SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Export Sales' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where  " & stringyear & " and CASHA.BAA<>0 and  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(Alpha.Text) + "%'   and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
''
''      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid ) SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & " FROM CREDITA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(Alpha.Text) + "%' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   "
''
''      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )  SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, '', " & UId & ",'" & main.session & "'," & main.setupid & "   From DNFA  where    " & stringyear & " and Pgld ='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld like '" + Trim(Alpha.Text) + "%'  and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
''
''      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )  SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '', " & UId & ",'" & main.session & "'," & main.setupid & "  From CNF1A where  " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld like '" + Trim(Alpha.Text) + "%'  and convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
''
''      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )  SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From DNFB  where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld like '" + Trim(Alpha.Text) + "%'   and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
''
''      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '', " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1B where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld like '" + Trim(Alpha.Text) + "%'   and convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
''
''   End If
   
''
''    If Trim(Alpha.Text) = "" And Alpha.Visible = True Then
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid) SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "'  AND   VOUCHERS.DebitorCredit='C' and  convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
''
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid,fyear,setupid) SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3  , " & UId & ",'" & main.session & "'," & main.setupid & " FROM INVOICEA  where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "'  and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
''                                'change By Dinesh
''               CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT Purchasea.GENLEDGER, Purchasea.SUBLEDGER, Purchasea.INVOICEDATE, 'P' AS Expr1, Purchasea.INVOICENO, 'Purchase Invoice' , Purchasea.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM Purchasea  where   genledger ='" + Trim(COMBOGENLEDGER.Text) + "'  and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
''
''
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid,fyear,setupid) SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Export Sales' , CASHA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM CASHA   where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "'    and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   "
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Export Sales' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where   " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "'   and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND CASHA.BAA <>0  "
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid) SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CREDITA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "'    and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
''
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid ) SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, ''  , " & UId & ",'" & main.session & "'," & main.setupid & " From DNFA  where    " & stringyear & " and Pgld ='" + Trim(COMBOGENLEDGER.Text) + "'and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
''
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid) SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1A where  " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.Text) + "'  and convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
''
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid ) SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, '', " & UId & ",'" & main.session & "'," & main.setupid & "  From DNFB  where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "'   and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid ) SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1B where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
''
''End If

If Trim(Alpha.Text) = "" And Alpha.Visible = False And cust <> "" Then

For I = 0 To K - 1
cust = arr(I)
If arr(I) <> "" Then
                
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)   SELECT GenLedger, SubLedger, VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'  AND   VOUCHERS.DebitorCredit='C' and  convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
                
                '-------------------------------------------------------------
                'change by dinesh
                
                If COMBOGENLEDGER.Text = "SALES" Then
                 CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)  SELECT '" & COMBOGENLEDGER.Text & "',INVOICEC.SUBLEDGER, INVOICEC.INVOICEDATE, 'I' AS Expr1, INVOICEC.INVOICENO, 'Sales' , INVOICEC.GAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM INVOICEC  where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  group by INVOICEC.SUBLEDGER,INVOICEC.INVOICEDATE,INVOICEC.INVOICENO,INVOICEC.GAMOUNT"
                Else
                 CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)  " & _
                 "SELECT '" & COMBOGENLEDGER.Text & "',c.SUBLEDGER, c.INVOICEDATE, 'I' AS Expr1, c.INVOICENO, a.subledger, c.AMOUNT, 'C' , '' AS Expr3 ," & _
                 " " & UId & ",'" & main.session & "'," & main.setupid & " FROM INVOICEC as c inner join invoicea as a on c.invoiceno=a.invoiceno " & _
                 "and  c.fyear=a.fyear AND c.setupid = a.setupid  " & _
                 "  where c.fyear='" & main.session & "'  and c.setupid=" & main.setupid & "  and c.genledger ='" + Trim(COMBOGENLEDGER.Text) + "'  and  c.Subledger = '" & Trim(cust) & "' and convert(smalldatetime,c.invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,c.invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  group by c.SUBLEDGER,c.INVOICEDATE,c.INVOICENO,c.AMOUNT,a.subledger"
                End If
                 '''CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)  SELECT '" & COMBOGENLEDGER.Text & "',INVOICEC.SUBLEDGER, INVOICEC.INVOICEDATE, 'I' AS Expr1, INVOICEC.INVOICENO, 'Sales' , INVOICEC.GAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM INVOICEC  where    " & stringyear & " and  genledger = '" & cust & "'  and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  group by INVOICEC.SUBLEDGER,INVOICEC.INVOICEDATE,INVOICEC.INVOICENO,INVOICEC.GAMOUNT"
                 
                 CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid) SELECT '" & COMBOGENLEDGER.Text & "',CASHC.SUBLEDGER, CASHC.INVOICEDATE, 'I' AS Expr1, CASHC.INVOICENO, '', CASHC.GAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM CASHC  where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'  and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  group by CASHC.SUBLEDGER,CASHC.INVOICEDATE,CASHC.INVOICENO,CASHC.GAMOUNT"
                 CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid) SELECT '" & COMBOGENLEDGER.Text & "',CREDITC.SUBLEDGER, CREDITC.INVOICEDATE, 'CI' AS Expr1, CREDITC.INVOICENO, 'Sales Return', CREDITC.GAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM CREDITC  where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'  and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  group by CREDITC.SUBLEDGER,CREDITC.INVOICEDATE,CREDITC.INVOICENO,CREDITC.GAMOUNT"
                 
                
                '-------------------------------------------------------------
                
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)  SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM INVOICEA  where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'  and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
                'change By Dinesh
                'CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT Purchasea.GENLEDGER, Purchasea.SUBLEDGER, Purchasea.INVOICEDATE, 'I' AS Expr1, Purchasea.INVOICENO, 'Purchase Invoice' , Purchasea.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM Purchasea  where   genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
                'CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid,header)  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, ' ' + casha.txt2a + ' @' + casha.CurrencyValue + '  S/Bill No. ' + casha.marka  , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & ",'Export Sales'  FROM CASHA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'   and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   "
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, '' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'   and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND CASHA.BAA <>0  "
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )    SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM CREDITA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'    and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )   SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & "   From DNFA  where    " & stringyear & " and Pgld ='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld = '" & cust & "' and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid,fyear,setupid )   SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1A where  " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld = '" & cust & "'    and convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid ,fyear,setupid)   SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, ''     , " & UId & ",'" & main.session & "'," & main.setupid & " From DNFB  where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld = '" & cust & "'   and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid ,fyear,setupid)   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1B where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld = '" & cust & "' and  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
End If
Next
cust = "a"
End If

''If Trim(Alpha.Text) = "" And cust = "" Then
''
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)   SELECT GenLedger, SubLedger, VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' AND   VOUCHERS.DebitorCredit='C' and  convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM INVOICEA  where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
''                'change By Dinesh
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT Purchasea.GENLEDGER, Purchasea.SUBLEDGER, Purchasea.INVOICEDATE, 'I' AS Expr1, Purchasea.INVOICENO, 'Purchase Invoice' , Purchasea.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM Purchasea  where   genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Export Sales' , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   "
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Export Sales' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND CASHA.BAA <>0  "
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )    SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM CREDITA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )   SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & "   From DNFA  where    " & stringyear & " and Pgld ='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid,fyear,setupid )   SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1A where  " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid ,fyear,setupid)   SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, ''     , " & UId & ",'" & main.session & "'," & main.setupid & " From DNFB  where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
''                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid ,fyear,setupid)   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1B where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
''End If
''

'On Error Resume Next
Dim rss1 As New ADODB.Recordset


CON.Execute "insert into Treport ( Genledger,Subledger,openingbalance,userid,fyear,setupid ) SELECT '" + Trim(COMBOGENLEDGER.Text) + "' as genled, SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId,'" & main.session & "'," & main.setupid & " from subledgertrail where  " & stringyear & " GROUP BY SUBLEDGER;"

If RS.State = 1 Then RS.close
RS.Open "SELECT GenLedger,[subledger],[vdate],[vno],[dorc],vtype,SNO  FROM [ExportData].[dbo].[treport] where (vtype='J' or vtype='P' or vtype='R') order by [vdate],[vno]", CON, adOpenKeyset, adLockReadOnly
While RS.EOF = False
If Not IsNull(RS!dorc) Then
If RS!dorc = "D" Then
 
  'CON.Execute "update [treport] set header  = (SELECT case when ISNULL(subledger, GenLedger) = '' then GenLedger else SubLedger END AS subledger  From VOUCHERS Where  " & stringyear & "  AND VoucherNumber=" & rs("vno") & " and  VOUCHERS.DebitorCredit='C' and convert(smalldatetime,voucherdate,103)=convert(smalldatetime,'" + Trim(rs!vdate) + "',103) AND vOUCHERtYPE='" & rs!vtype & "') Where  SNO=" & rs!sno & ""
  ''CON.Execute "update [treport] set header  = (SELECT case when ISNULL(subledger, GenLedger) = '' then GenLedger else SubLedger END AS subledger  From VOUCHERS Where  " & stringyear & "  AND VoucherNumber=" & rs("vno") & " and  VOUCHERS.DebitorCredit='C' and convert(smalldatetime,voucherdate,103)=convert(smalldatetime,'" + Trim(rs!vdate) + "',103)  AND vOUCHERtYPE='" & rs!vtype & "') Where  SNO=" & rs!sno & ""
 
   ss = "SELECT case when ISNULL(subledger, GenLedger) = '' then GenLedger else SubLedger END AS subledger  From VOUCHERS Where " & stringyear & "  AND VoucherNumber=" & RS("vno") & " and  VOUCHERS.DebitorCredit='C' and convert(smalldatetime,voucherdate,103)=convert(smalldatetime,'" + Trim(RS!vdate) + "',103) AND vOUCHERtYPE='" & RS!vtype & "'"
   If rss1.State = 1 Then rss1.close
   rss1.Open ss, CON
   While rss1.EOF = False
   
    ss11 = rss1!SUBLEDGER
    If Len(rss1(0)) > 0 Then
      CON.Execute "update [treport] set header  = '" & ss11 & "' Where  SNO=" & RS!sno & ""
    End If
    
    If rss1.RecordCount = 2 Then
       rss1.MoveNext
    ElseIf rss1.RecordCount = 3 Then
       rss1.MoveNext
       rss1.MoveNext
    End If
   
   rss1.MoveNext
   Wend


Else
 
   'CON.Execute "update [treport] set header  = (SELECT case when ISNULL(subledger, GenLedger) = '' then GenLedger else SubLedger END AS subledger  From VOUCHERS Where  " & stringyear & "  AND VoucherNumber=" & rs("vno") & " and  VOUCHERS.DebitorCredit='D' and convert(smalldatetime,voucherdate,103)=convert(smalldatetime,'" + Trim(rs!vdate) + "',103) AND vOUCHERtYPE='" & rs!vtype & "') Where  SNO=" & rs!sno & ""
   
   ss = "SELECT case when ISNULL(subledger, GenLedger) = '' then GenLedger else SubLedger END AS subledger  From VOUCHERS Where " & stringyear & "  AND VoucherNumber=" & RS("vno") & " and  VOUCHERS.DebitorCredit='D' and convert(smalldatetime,voucherdate,103)=convert(smalldatetime,'" + Trim(RS!vdate) + "',103) AND vOUCHERtYPE='" & RS!vtype & "'"
   If rss1.State = 1 Then rss1.close
   rss1.Open ss, CON
   While rss1.EOF = False
   
    ss11 = rss1!SUBLEDGER
    If Len(rss1(0)) > 0 Then
      CON.Execute "update [treport] set header  = '" & ss11 & "' Where  SNO=" & RS!sno & ""
    End If
    
    If rss1.RecordCount = 2 Then
       rss1.MoveNext
    ElseIf rss1.RecordCount = 3 Then
       rss1.MoveNext
       rss1.MoveNext
    End If
   
   rss1.MoveNext
   Wend
   
   
End If
End If
RS.MoveNext
Wend



'------------------------------
If RS.State = 1 Then RS.close
RS.Open "SELECT [subledger],[vdate],[vno],[dorc],vtype,SNO  FROM [ExportData].[dbo].[treport] where subledger='Export Sales' order by [vdate],[vno]", CON, adOpenKeyset, adLockReadOnly
While RS.EOF = False
 CON.Execute "update [treport] set header  = (SELECT case when ISNULL(subledger, GenLedger) = '' then GenLedger else SubLedger END AS subledger " & _
 " From casha Where  " & stringyear & " and  invoiceno='" & RS!vno & "') Where  SNO=" & RS!sno & ""
 RS.MoveNext
Wend
'------------------------------


main.reportname = "Sub. Ledger"
If SLEDGERPRINT.Alpha.Visible = True Then
   viewledger.SelectedParty
Else
   viewledger.SelectedParty
End If
Unload PrintOption
Load PrintOption

End Sub
Private Sub Commandreturn_Click()
    Unload Me
    Unload PrintOption
End Sub
Private Sub Commandshow_Click()

'********sub for alpha wise and Partywise according to new fast mathed
cust = ""
For I = 0 To Me.Combosubledger.ListCount - 1
If Me.Combosubledger.Selected(I) = True Then
cust = Me.Combosubledger.List(I)
GoTo abc:
End If
Next

abc:
Commandshow.Enabled = False
DoEvents

CON.Execute "DELETE from subledgertrail where LEN(SUBLEDGER)>0"
CON.Execute "DELETE from TREPORT where LEN(SUBLEDGER)>0"

DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
ALPHAB
Commandshow.Enabled = True
End Sub

Private Sub date1_KeyPress(KeyAscii As Integer)
  
    If KeyAscii = 13 Then
    'date2.SetFocus
End If
End Sub

Private Sub date1_LostFocus()
    If Trim(date1.Text) <> "" Then
        If Not checkdate(Trim(date1.Text), date1) Then
            date1.SetFocus
            End If
    End If
End Sub

Private Sub date2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   ' SendKeys "{TAB}"
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


BackColorFrom Me

Me.Top = 0
Me.Left = 0
Set RS = New ADODB.Recordset
    RS.Open "select * from gledger where  " & stringyear & " and slf=1", CON, adOpenStatic, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        Do While Not RS.EOF
            COMBOGENLEDGER.AddItem Trim(RS!gledger)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    RS.Open "Select * from setup where " & stringyear & "", CON, adOpenStatic, adLockReadOnly, adCmdText
    CNSetup
    date1.Text = RS!yarfrom
    date2.Text = RS!yarto
    RS.close
End Sub
Sub xx()

''    Else
''
''                ' opening Balance Start
''                rs.Open "select * from sledger where gledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "'", con, adopenStatic, adLockReadOnly, adCmdText
''                BALANCE =myround(rs!YEAROPENING, 2)
''                rs.Close
''                ' vouchers opening balance
''                rs.Open "select sum(amount) from vouchers where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DebitorCredit='D'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(amount) from vouchers where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DebitorCredit='C'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE - rs(0)
''                End If
''                rs.Close
''                ' invoice opening balance
''                rs.Open "select sum(NETAMOUNT) from invoicea where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                    ' cash counter opening balance
''                rs.Open "select sum(NETAMOUNT) from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(baa) from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE - rs(0)
''                End If
''                rs.Close
''
''                ' Credit note Item opening balance
''                rs.Open "select sum(NETAMOUNT) from credita where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                'credit note opening balance
''                rs.Open "select sum(NA) from cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DC='D'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(NA) from cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DC='C'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE - rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(A) from cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DC='D'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(A) from cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DC='C'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE - rs(0)
''                End If
''                rs.Close
''
''
''                'debit note opening balance
''                 rs.Open "select sum(NA) from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DC='D'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(NA) from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DC='C'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE - rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(A) from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DC='D'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(A) from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DC='C'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE - rs(0)
''                End If
''                rs.Close
''            ' balance field contain the opening amount upto the date
''                Set rs1 = New ADODB.Recordset
''                CON.Execute ("DELETE from treport")
''                rs1.Open "treport", CON, adopenStatic, adLockOptimistic, adcmdtext
''                rs1.AddNew
''                rs1!Text = "** Opening Balance as on " + Trim(date1.Text)
''                rs1!ad = BALANCE
''                rs1!period = Trim(date1.Text) + "  To  " + Trim(date2.Text)
''                rs1!header = "SUB. LEDGER ACCOUNT"
''                rs1!subledger = Trim(Combosubledger.Text)
''                rs1!dorc = "D"
''                rs1.Update
''               'opening the tables
''                rs.Open "select * from vouchers where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) order by  VoucherDate,vouchertype,vouchernumber", con, adopenStatic, adLockReadOnly, adCmdText
''                rs2.Open "select * from invoicea where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) order by invoicedate", con, adopenStatic, adLockReadOnly, adCmdText
''
''                rs3.Open "select * from CREDITa where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", con, adopenStatic, adLockReadOnly, adCmdText
''                rs4.Open "select * from Cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,cnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", con, adopenStatic, adLockReadOnly, adCmdText
''                rs5.Open "select * from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", con, adopenStatic, adLockReadOnly, adCmdText
''                rs6.Open "select * from Cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,cnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", con, adopenStatic, adLockReadOnly, adCmdText
''                rs7.Open "select * from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", con, adopenStatic, adLockReadOnly, adCmdText
''                rs8.Open "select * from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) order by invoicedate", con, adopenStatic, adLockReadOnly, adCmdText
''
''

End Sub

Sub OPENINGSUBLEDGERS()

          
        
If Trim(Alpha.Text) = "" And cust = "" Then
         
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , YEAROPENING,  0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
''        & " FROM SLEDGER where   " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "'", p, adCmdText
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
''        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER) AND (SLEDGER.fyear = INVOICEA.fyear) AND (SLEDGER.setupid = INVOICEA.setupid))  " _
''        & " where sledger.fyear='" & main.session & "' and invoicea.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "'and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''        & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(Purchasea.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
''        & " FROM (SLEDGER LEFT JOIN Purchasea ON (SLEDGER.SUBLEDGER = Purchasea.SUBLEDGER) AND (SLEDGER.gledger = Purchasea.GENLEDGER))  " _
''        & " where  sledger.fyear='" & main.session & "' and purchasea.setupid=" & main.setupid & " and  genledger='" + Trim(COMBOGENLEDGER.Text) + "'and convert(smalldatetime,INVOICEDATE,103) <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " & _
''        " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
''
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
''        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid)) " _
''        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and  SLEDGER.SUBLEDGER <>'CASH PARTY'    and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''        & " GROUP BY SLEDGER.SUBLEDGER " _
''        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
''        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))" _
''        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''        & " GROUP BY SLEDGER.SUBLEDGER " _
''        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
''
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
''        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER) AND (SLEDGER.fyear = credita.fyear) AND (SLEDGER.setupid = credita.setupid)) " _
''        & " where sledger.fyear='" & main.session & "' and credita.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''        & " GROUP BY SLEDGER.SUBLEDGER " _
''        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
''        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid) " _
''        & " WHERE sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''        & " GROUP BY SLEDGER.SUBLEDGER " _
''        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
''
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
''        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger)  AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid)" _
''        & " WHERE sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''        & " GROUP BY SLEDGER.SUBLEDGER " _
''        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
''         '''ok
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
''        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD)  AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid)" _
''        & " WHERE sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='D' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''        & " GROUP BY SLEDGER.SUBLEDGER " _
''        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
''        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid) " _
''        & " WHERE sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='C' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''        & " GROUP BY SLEDGER.SUBLEDGER " _
''        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
''
''
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
''        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)  AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid)" _
''        & " WHERE sledger.fyear='" & main.session & "' and dnfa.setupid=" & main.setupid & " and DNFA.DC = 'D' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''        & " GROUP BY SLEDGER.SUBLEDGER " _
''        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
''
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
''        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid)" _
''        & " WHERE sledger.fyear='" & main.session & "' and dnfa.setupid=" & main.setupid & " and DNFA.DC='C' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''        & " GROUP BY SLEDGER.SUBLEDGER " _
''        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
''        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid) " _
''        & " WHERE sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC='D' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
''        & " GROUP BY SLEDGER.SUBLEDGER " _
''        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
''
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
''        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD)  AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid)" _
''        & " WHERE sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC='C' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
''        & " GROUP BY SLEDGER.SUBLEDGER " _
''        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
''
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
''        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD)  AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid)" _
''        & " WHERE sledger.fyear='" & main.session & "' and dnfb.setupid=" & main.setupid & " and DNFB.DC='D' and gld = '" + Trim(COMBOGENLEDGER.Text) + "'And convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''        & " GROUP BY SLEDGER.SUBLEDGER " _
''        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
''
''        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
''        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)  AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid)" _
''        & " WHERE sledger.fyear='" & main.session & "' and dnfb.setupid=" & main.setupid & " and DNFB.DC='C' and gld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''        & " GROUP BY SLEDGER.SUBLEDGER " _
''        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
''
''
''
  
  
  
  
''''comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , YEAROPENING,  0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
''''        & " FROM SLEDGER where  gledger='" + Trim(COMBOGENLEDGER.Text) + "'", p, adCmdText
''''
''''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
''''        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER))  " _
''''        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "'and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''''        & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
''''
''''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER)) " _
''''        & " where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' and  SLEDGER.SUBLEDGER <>'CASH PARTY'    and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
''''
''''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
''''        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER))" _
''''        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
''''
''''
''''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER)) " _
''''        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
''''
''''
''''
''''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
''''        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
''''        & " WHERE DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
''''
''''
''''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
''''        & " WHERE DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
''''         ''''ok
''''
''''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
''''        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD)  AND (SLEDGER.fyear = CNF1A.fyear) AND (SLEDGER.setupid = CNF1A.setupid) " _
''''        & " WHERE (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
''''
''''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD)  AND (SLEDGER.fyear = CNF1A.fyear) AND (SLEDGER.setupid = CNF1A.setupid) " _
''''        & " WHERE (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
''''
''''
''''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = dnfa.PSLD)  AND (SLEDGER.fyear = dNFA.fyear) AND (SLEDGER.setupid = dNFA.setupid) " _
''''        & " WHERE ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
''''
''''
''''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = dnfa.PSLD)  AND (SLEDGER.fyear = dNFA.fyear) AND (SLEDGER.setupid = dNFA.setupid) " _
''''        & " WHERE (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
''''
''''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1b.SLD)  AND (SLEDGER.fyear = CNF1b.fyear) AND (SLEDGER.setupid = CNF1b.setupid) " _
''''        & " WHERE (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
''''
''''
''''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1b.SLD)  AND (SLEDGER.fyear = CNF1b.fyear) AND (SLEDGER.setupid = CNF1b.setupid) " _
''''        & " WHERE (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
''''
''''
''''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = dnfb.SLD)  AND (SLEDGER.fyear = dNFb.fyear) AND (SLEDGER.setupid = dNFb.setupid) " _
''''        & " WHERE (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.Text) + "'And convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
''''
''''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
''''        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = dnfb.SLD)  AND (SLEDGER.fyear = dNFb.fyear) AND (SLEDGER.setupid = dNFb.setupid) " _
''''        & " WHERE (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
''''        & " GROUP BY SLEDGER.SUBLEDGER " _
''''        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
  
  End If
  
If Trim(Alpha.Text) = "" And cust <> "" Then
        K = 0
    Dim k1 As Integer
    k1 = 0
    'cust = ""
    ReDim Preserve arr(Me.Combosubledger.ListCount)
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
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , SLEDGER.YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "  FROM SLEDGER  where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER = '" & cust & "' and sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid, p, adCmdText
       
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER) AND (SLEDGER.fyear = INVOICEA.fyear) AND (SLEDGER.setupid = INVOICEA.setupid))  " _
        & " where sledger.fyear='" & main.session & "' and invoicea.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and sledger.SUBLEDGER = '" & cust & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
  
  
        '''''''''''changing by dinesh
        '''''''''''--------------------
        
      If COMBOGENLEDGER.Text = "SALES" Then
      
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  INVOICEC.GAMOUNT AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN INVOICEC ON (SLEDGER.SUBLEDGER = INVOICEC.SUBLEDGER) AND (SLEDGER.gledger = INVOICEC.GENLEDGER) AND (SLEDGER.fyear = INVOICEC.fyear) AND (SLEDGER.setupid = INVOICEC.setupid))" _
        & " where  sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER = '" & cust & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER,INVOICEC.GAMOUNT "
        
     Else
     
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  INVOICEC.AMOUNT AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN INVOICEC ON (SLEDGER.SUBLEDGER = INVOICEC.SUBLEDGER) AND (SLEDGER.gledger = INVOICEC.GENLEDGER) AND (SLEDGER.fyear = INVOICEC.fyear) AND (SLEDGER.setupid = INVOICEC.setupid))" _
        & " where  sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER = '" & cust & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER,INVOICEC.AMOUNT "
        
     End If
        

''       CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  INVOICEC.GAMOUNT AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
''       & " FROM (SLEDGER LEFT JOIN INVOICEC ON (SLEDGER.SUBLEDGER = INVOICEC.SUBLEDGER) AND (SLEDGER.gledger = INVOICEC.GENLEDGER) AND (SLEDGER.fyear = INVOICEC.fyear) AND (SLEDGER.setupid = INVOICEC.setupid))" _
''       & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & "  and SLEDGER.SUBLEDGER = '" & cust & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%' " _
''        & " GROUP BY SLEDGER.SUBLEDGER,INVOICEC.GAMOUNT "

        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  CASHC.GAMOUNT AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHC ON (SLEDGER.SUBLEDGER = CASHC.SUBLEDGER) AND (SLEDGER.gledger = CASHC.GENLEDGER) AND (SLEDGER.fyear = CASHC.fyear) AND (SLEDGER.setupid = CASHC.setupid))" _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER = '" & cust & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER,CASHC.GAMOUNT "
        
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  CREDITC.GAMOUNT AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CREDITC ON (SLEDGER.SUBLEDGER = CREDITC.SUBLEDGER) AND (SLEDGER.gledger = CREDITC.GENLEDGER) AND (SLEDGER.fyear = CREDITC.fyear) AND (SLEDGER.setupid = CREDITC.setupid))" _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER = '" & cust & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER,CREDITC.GAMOUNT "
        

        '-------------------------------
        
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid)) " _
        & " where   sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & "  and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER = '" & cust & "' and SLEDGER.SUBLEDGER=  '" & cust & "' and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText

        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))" _
        & " where   sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & "  and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER = '" & cust & "' and SLEDGER.SUBLEDGER ='" & cust & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText



        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER) AND (SLEDGER.fyear = credita.fyear) AND (SLEDGER.setupid = credita.setupid)) " _
        & " where sledger.fyear='" & main.session & "' and credita.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER='" & cust & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText



        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER= '" & cust & "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText


        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER= '" & cust & "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
         ''''ok

        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='D' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "'and CNF1A.PSLD = '" & cust & "'  and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText

        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD)  AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid)" _
        & " WHERE sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='C' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and CNF1A.PSLD = '" & cust & "'  and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText



        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)  AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid)" _
        & " WHERE sledger.fyear='" & main.session & "' and DNFA.setupid=" & main.setupid & " and DNFA.DC='D' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and   DNFA.PSLD = '" & cust & "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText


        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid)" _
        & " WHERE sledger.fyear='" & main.session & "' and DNFA.setupid=" & main.setupid & " and DNFA.DC='C' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "'  and  DNFA.PSLD = '" & cust & "'   and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText

        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD)  AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid)" _
        & " WHERE sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC='D' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and  CNF1B.SLD = '" & cust & "'    and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText


        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid) " _
        & " WHERE sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC= 'C' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "'  and  CNF1B.SLD = '" & cust & "'      and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText


        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD)  AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid)" _
        & " WHERE sledger.fyear='" & main.session & "' and DNFB.setupid=" & main.setupid & " and DNFB.DC='D' and gld = '" + Trim(COMBOGENLEDGER.Text) + "'  and  DNFB.SLD = '" & cust & "'     And convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText

        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD) AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid)" _
        & " WHERE sledger.fyear='" & main.session & "' and DNFB.setupid=" & main.setupid & " and DNFB.DC='C' and gld = '" + Trim(COMBOGENLEDGER.Text) + "'  and  DNFB.SLD = '" & cust & "'    and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
        End If
  End If
  Next I
  
  cust = "a"

'''
'''
'''      comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , SLEDGER.YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId  FROM SLEDGER  where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER = '" & Combosubledger.Text & "'", p, adCmdText
'''
'''
'''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
'''        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER))  " _
'''        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER=  '" & Combosubledger.Text & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'''        & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
'''
'''
'''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER)) " _
'''        & " where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER=  '" & Combosubledger.Text & "'    and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
'''
'''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
'''        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER))" _
'''        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER ='" & Combosubledger.Text & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
'''
'''
'''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER)) " _
'''        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER='" & Combosubledger.Text & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
'''
'''
'''
'''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
'''        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
'''        & " WHERE DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER= '" & Combosubledger.Text & "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
'''
'''
'''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
'''        & " WHERE DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER= '" & Combosubledger.Text & "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
'''         ''''ok
'''
'''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
'''        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD)  AND (SLEDGER.fyear = CNF1A.fyear) AND (SLEDGER.setupid = CNF1A.setupid) " _
'''        & " WHERE (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "'and CNF1A.PSLD = '" & Combosubledger.Text & "'  and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
'''
'''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD)  AND (SLEDGER.fyear = CNF1A.fyear) AND (SLEDGER.setupid = CNF1A.setupid) " _
'''        & " WHERE (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and CNF1A.PSLD = '" & Combosubledger.Text & "'  and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
'''
'''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = dnfa.PSLD)  AND (SLEDGER.fyear = dNFA.fyear) AND (SLEDGER.setupid = dNFA.setupid) " _
'''        & " WHERE ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and   DNFA.PSLD = '" & Combosubledger.Text & "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
'''
'''
'''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = dnfa.PSLD)  AND (SLEDGER.fyear = dNFA.fyear) AND (SLEDGER.setupid = dNFA.setupid) " _
'''        & " WHERE (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.Text) + "'  and  DNFA.PSLD = '" & Combosubledger.Text & "'   and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
'''
'''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1b.SLD)  AND (SLEDGER.fyear = CNF1b.fyear) AND (SLEDGER.setupid = CNF1b.setupid) " _
'''        & " WHERE (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and  CNF1B.SLD = '" & Combosubledger.Text & "'    and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
'''
'''
'''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1b.SLD)  AND (SLEDGER.fyear = CNF1b.fyear) AND (SLEDGER.setupid = CNF1b.setupid) " _
'''        & " WHERE (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "'  and  CNF1B.SLD = '" & Combosubledger.Text & "'      and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
'''
'''
'''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = dnfb.SLD)  AND (SLEDGER.fyear = dNFb.fyear) AND (SLEDGER.setupid = dNFb.setupid) " _
'''        & " WHERE (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.Text) + "'  and  DNFB.SLD = '" & Combosubledger.Text & "'     And convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
'''
'''        comcon.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'''        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = dnfb.SLD)  AND (SLEDGER.fyear = dNFb.fyear) AND (SLEDGER.setupid = dNFb.setupid) " _
'''        & " WHERE (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.Text) + "'  and  DNFB.SLD = '" & Combosubledger.Text & "'    and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
'''        & " GROUP BY SLEDGER.SUBLEDGER " _
'''        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
'''

  End If
  
End Sub

