VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form SLEDGERPRINT 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5805
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "SLPRINT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Alpha 
      Height          =   315
      Left            =   3060
      MaxLength       =   1
      TabIndex        =   2
      Top             =   2070
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.ComboBox Combosubledger 
      Height          =   315
      Left            =   3060
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1590
      Visible         =   0   'False
      Width           =   4665
   End
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   405
      Left            =   4590
      TabIndex        =   6
      Top             =   3570
      Width           =   1545
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      Height          =   405
      Left            =   2910
      TabIndex        =   5
      Top             =   3570
      Width           =   1545
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   3060
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1170
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
      Left            =   3060
      TabIndex        =   3
      Top             =   2580
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
      Left            =   5610
      TabIndex        =   4
      Top             =   2580
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
      Caption         =   "From The Date"
      Height          =   315
      Left            =   870
      TabIndex        =   11
      Top             =   2610
      Width           =   1995
   End
   Begin VB.Label Label3 
      Caption         =   " - To - "
      Height          =   315
      Left            =   4710
      TabIndex        =   10
      Top             =   2610
      Width           =   585
   End
   Begin VB.Label Label5 
      Caption         =   "Alphabat"
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   2130
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Sub. Ledger Desc."
      Height          =   285
      Left            =   960
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Gen. Ledger Desc."
      Height          =   285
      Left            =   960
      TabIndex        =   7
      Top             =   1170
      Width           =   2055
   End
End
Attribute VB_Name = "SLEDGERPRINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As Connection
Dim rs As Recordset


Private Sub COMBOGENLEDGER_Change()
    If rs.State = 1 Then
        rs.Close
    End If
    rs.Open "select * from sledger where gledger='" + Trim(COMBOGENLEDGER.Text) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
    Combosubledger.Clear
    If Not rs.BOF Then
        Do While Not rs.EOF
            Combosubledger.AddItem Trim(rs!SubLedger)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    
End Sub

Private Sub COMBOGENLEDGER_Click()
If rs.State = 1 Then
        rs.Close
    End If
    rs.Open "select * from sledger where gledger='" + Trim(COMBOGENLEDGER.Text) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
    Combosubledger.Clear
    If Not rs.BOF Then
        Do While Not rs.EOF
            Combosubledger.AddItem Trim(rs!SubLedger)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
End Sub

Private Sub COMBOGENLEDGER_LostFocus()
    If Trim(COMBOGENLEDGER.Text) <> "" Then
        rs.Open "select * from gledger where slf=true", CON, adOpenStatic, adLockReadOnly, adCmdText
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

Private Sub Combosubledger_GotFocus()
    If Trim(COMBOGENLEDGER.Text) = "" Then
        COMBOGENLEDGER.SetFocus
    End If
End Sub

Private Sub Combosubledger_LostFocus()
If Trim(Combosubledger.Text) <> "" Then
    If Trim(COMBOGENLEDGER.Text) <> "" Then
        If rs.State = 1 Then
            rs.Close
        End If
        rs.Open "select * from sledger where gledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
        If rs.BOF Then
            Combosubledger.SetFocus
        End If
        rs.Close
    Else
        Combosubledger.Text = ""
    End If
End If
End Sub

 Sub ALPHAB()
    If rs.State = 1 Then rs.Close
    DoEvents
    CON.Execute ("delete * from treport")
    
    If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
        MsgBox "invalid date"
        Exit Sub
    End If
    Dim rs1 As New ADODB.Recordset
    Dim BALANCE As Double
    Dim OPBALANCE As Double
    Dim SDamount As Double
    Dim SCamount As Double
    Dim RsT As New ADODB.Recordset
    Dim viewsubledger As Boolean
    viewsubledger = False
    BALANCE = 0
    OPBALANCE = 0
    OPENINGSUBLEDGERS
    DoEvents
    If Trim(alpha.Text) <> "" And alpha.Visible = True Then
      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )   SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & " From VOUCHERS Where   genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(alpha.Text) + "%' AND  VOUCHERS.DebitorCredit='D' and  VoucherDate >= cdate('" + Trim(date1.Text) + "')   and VoucherDate <=cdate('" + Trim(date2.Text) + "')      ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )    SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & " From VOUCHERS Where   genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(alpha.Text) + "%'  AND   VOUCHERS.DebitorCredit='C' and VoucherDate >= cdate('" + Trim(date1.Text) + "')   and VoucherDate <=cdate('" + Trim(date2.Text) + "')    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
 
      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT  INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & " FROM INVOICEA  where   genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(alpha.Text) + "%'  and InvoiceDate >= cdate('" + Trim(date1.Text) + "')   and InvoiceDate <=cdate('" + Trim(date2.Text) + "') "
      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )   SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & "  FROM CASHA  where   genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(alpha.Text) + "%'   and InvoiceDate >= cdate('" + Trim(date1.Text) + "')   and InvoiceDate <=cdate('" + Trim(date2.Text) + "')  "
      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )   SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & "  FROM CASHA  where CASHA.BAA<>0 and  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(alpha.Text) + "%'   and InvoiceDate >= cdate('" + Trim(date1.Text) + "')   and InvoiceDate <=cdate('" + Trim(date2.Text) + "') "

      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3, " & UId & " FROM CREDITA  where   genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger like '" + Trim(alpha.Text) + "%' and InvoiceDate >= cdate('" + Trim(date1.Text) + "')   and InvoiceDate <=cdate('" + Trim(date2.Text) + "')   "

      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, '', " & UId & "   From DNFA  where   Pgld ='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld like '" + Trim(alpha.Text) + "%'  and dnd >= cdate('" + Trim(date1.Text) + "')   and dnd <=cdate('" + Trim(date2.Text) + "')  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
 
      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '', " & UId & "  From CNF1A where Pgld='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld like '" + Trim(alpha.Text) + "%'  and cnd >= cdate('" + Trim(date1.Text) + "')   and cnd <=cdate('" + Trim(date2.Text) + "')   ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
 
      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, '' , " & UId & " From DNFB  where gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld like '" + Trim(alpha.Text) + "%'   and dnd >= cdate('" + Trim(date1.Text) + "')   and dnd <=cdate('" + Trim(date2.Text) + "')    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"

      CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '', " & UId & " From CNF1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld like '" + Trim(alpha.Text) + "%'   and cnd >= cdate('" + Trim(date1.Text) + "')   and cnd <=cdate('" + Trim(date2.Text) + "')   ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
   End If
   
   
    If Trim(alpha.Text) = "" And alpha.Visible = True Then
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & " From VOUCHERS Where   genledger ='" + Trim(COMBOGENLEDGER.Text) + "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate >= cdate('" + Trim(date1.Text) + "')   and VoucherDate <=cdate('" + Trim(date2.Text) + "')    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & " From VOUCHERS Where   genledger ='" + Trim(COMBOGENLEDGER.Text) + "'  AND   VOUCHERS.DebitorCredit='C' and  VoucherDate >= cdate('" + Trim(date1.Text) + "')   and VoucherDate <=cdate('" + Trim(date2.Text) + "')   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
 
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid) SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3  , " & UId & " FROM INVOICEA  where   genledger ='" + Trim(COMBOGENLEDGER.Text) + "'  and InvoiceDate >= cdate('" + Trim(date1.Text) + "')   and InvoiceDate <=cdate('" + Trim(date2.Text) + "') "
    
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid) SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & " FROM CASHA   where   genledger='" + Trim(COMBOGENLEDGER.Text) + "'    and InvoiceDate >= cdate('" + Trim(date1.Text) + "')   and InvoiceDate <=cdate('" + Trim(date2.Text) + "')   "
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & "  FROM CASHA  where  genledger='" + Trim(COMBOGENLEDGER.Text) + "'   and InvoiceDate >= cdate('" + Trim(date1.Text) + "')   and InvoiceDate <=cdate('" + Trim(date2.Text) + "')  AND CASHA.BAA <>0  "
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3, " & UId & "  FROM CREDITA  where   genledger='" + Trim(COMBOGENLEDGER.Text) + "'    and InvoiceDate >= cdate('" + Trim(date1.Text) + "') and InvoiceDate <=cdate('" + Trim(date2.Text) + "') "

                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, ''  , " & UId & " From DNFA  where   Pgld ='" + Trim(COMBOGENLEDGER.Text) + "'and dnd >= cdate('" + Trim(date1.Text) + "')   and dnd <=cdate('" + Trim(date2.Text) + "')  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
 
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '' , " & UId & " From CNF1A where Pgld='" + Trim(COMBOGENLEDGER.Text) + "'  and cnd >= cdate('" + Trim(date1.Text) + "')   and cnd <=cdate('" + Trim(date2.Text) + "')  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
 
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, '', " & UId & "  From DNFB  where gld='" + Trim(COMBOGENLEDGER.Text) + "'   and dnd >= cdate('" + Trim(date1.Text) + "')   and dnd <=cdate('" + Trim(date2.Text) + "')    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ) SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & " From CNF1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and  cnd >= cdate('" + Trim(date1.Text) + "')   and cnd <=cdate('" + Trim(date2.Text) + "') ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
End If
If Trim(alpha.Text) = "" And alpha.Visible = False And Combosubledger.Text <> "" Then
       
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )   SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & " From VOUCHERS Where   genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate >= cdate('" + Trim(date1.Text) + "')   and VoucherDate <=cdate('" + Trim(date2.Text) + "')    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
                
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & " From VOUCHERS Where   genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "'  AND   VOUCHERS.DebitorCredit='C' and  VoucherDate >= cdate('" + Trim(date1.Text) + "')   and VoucherDate <=cdate('" + Trim(date2.Text) + "')   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
 
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & " FROM INVOICEA  where   genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "'  and InvoiceDate >= cdate('" + Trim(date1.Text) + "')   and InvoiceDate <=cdate('" + Trim(date2.Text) + "') "
    
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & "  FROM CASHA  where   genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "'   and InvoiceDate >= cdate('" + Trim(date1.Text) + "')   and InvoiceDate <=cdate('" + Trim(date2.Text) + "')   "
                
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & "  FROM CASHA  where   genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "'   and InvoiceDate >= cdate('" + Trim(date1.Text) + "')   and InvoiceDate <=cdate('" + Trim(date2.Text) + "')  AND CASHA.BAA <>0  "
                
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )    SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & " FROM CREDITA  where   genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "'    and InvoiceDate >= cdate('" + Trim(date1.Text) + "') and InvoiceDate <=cdate('" + Trim(date2.Text) + "') "

                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )   SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, '' , " & UId & "   From DNFA  where   Pgld ='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld = '" & Combosubledger.Text & "' and dnd >= cdate('" + Trim(date1.Text) + "')   and dnd <=cdate('" + Trim(date2.Text) + "')  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
 
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid )   SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '' , " & UId & " From CNF1A where Pgld='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld = '" & Combosubledger.Text & "'    and cnd >= cdate('" + Trim(date1.Text) + "')   and cnd <=cdate('" + Trim(date2.Text) + "')  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
 
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid )   SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, ''     , " & UId & " From DNFB  where gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld = '" & Combosubledger.Text & "'   and dnd >= cdate('" + Trim(date1.Text) + "')   and dnd <=cdate('" + Trim(date2.Text) + "')    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
                
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid )   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & " From CNF1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld = '" & Combosubledger.Text & "' and  cnd >= cdate('" + Trim(date1.Text) + "')   and cnd <=cdate('" + Trim(date2.Text) + "') ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
End If
CON.Execute "insert into Treport ( Genledger,Subledger,openingbalance,userid ) SELECT '" + Trim(COMBOGENLEDGER.Text) + "'as genled, SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId from subledgertrail GROUP BY SUBLEDGER;"
main.reportname = "Sub. Ledger"
If SLEDGERPRINT.alpha.Visible = True Then
   viewledger.genreport1
Else
   viewledger.genreport1
End If
PrintOption.Show
End Sub

Private Sub Commandreturn_Click()
    MainMenu.Toolbar1.Visible = True
    Unload Me
End Sub
Private Sub Commandshow_Click()

'********sub for alpha wise and Partywise according to new fast mathed
 DoEvents
 CON.Execute "Delete * from subledgertrail"
 DoEvents
 ALPHAB




'by vk
'
'
'If Trim(COMBOGENLEDGER.Text) <> "" Then
'    If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
'        MsgBox "invalid date"
'        Exit Sub
'    End If
'    'And Trim(Combosubledger.Text) <> ""
'
'    Dim rs1 As ADODB.Recordset
'    Dim rs2 As ADODB.Recordset
'    Dim rs3 As ADODB.Recordset
'    Dim rs4 As ADODB.Recordset
'    Dim rs5 As ADODB.Recordset
'    Dim rs6 As ADODB.Recordset
'    Dim rs7 As ADODB.Recordset
'    Dim rs8 As ADODB.Recordset
'    Dim rsSLD As New ADODB.Recordset
'
'    Dim CON1 As ADODB.Connection
'    Set CON1 = New ADODB.Connection
'     Set rs1 = New ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Set rs3 = New ADODB.Recordset
'    Set rs4 = New ADODB.Recordset
'    Set rs5 = New ADODB.Recordset
'    Set rs6 = New ADODB.Recordset
'    Set rs7 = New ADODB.Recordset
'    Set rs8 = New ADODB.Recordset
'
'    Dim BALANCE As Double
'    BALANCE = 0
'    CON1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + Trim(VB.App.Path) + "\" + Trim(main.directory) + "\tchitra.mdb"
'    CON1.Open
'    If rs.State = 1 Then
'        rs.Close
'    End If
'    CON1.Execute ("delete * from treport")
'    If Trim(alpha.Text) <> "" Then
'        rsSLD.Open "Select * from  sledger where   gledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & Combosubledger.Text & "'", con, adopenStatic, adLockReadOnly, adCmdText
'    Else
'        rsSLD.Open "select * from sledger where gledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "'", con, adopenStatic, adLockReadOnly, adCmdText
'    End If
'
'    If Not rsSLD.BOF Then
'
'         While Not rsSLD.EOF
'
'            '    If rs.State = 1 Then rs.Close
'            '    rs.Open "select * from sledger where gledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger   = '" & RsSLD(1) & "'", con, adopenStatic, adLockReadOnly, adCmdText
'                BALANCE = Round(rsSLD!YEAROPENING, 2)
'             '   rs.Close
'                ' vouchers opening balance
'                If rs.State = 1 Then rs.Close
'                rs.Open "select sum(amount) from vouchers where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger   = '" & rsSLD(1) & "' and VoucherDate<cdate('" + Trim(date1.Text) + "') and DebitorCredit='D'"
'                If rs(0) >= 0 Then
'                    BALANCE = BALANCE + rs(0)
'                End If
'                rs.Close
'                rs.Open "select sum(amount) from vouchers where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger   = '" & rsSLD(1) & "' and VoucherDate<cdate('" + Trim(date1.Text) + "') and DebitorCredit='C'"
'                If rs(0) >= 0 Then
'                    BALANCE = BALANCE - rs(0)
'                End If
'                rs.Close
'                ' invoice opening balance
'                rs.Open "select sum(NETAMOUNT) from invoicea where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger  = '" & rsSLD(1) & "' and INVOICEDATE<cdate('" + Trim(date1.Text) + "')"
'                If rs(0) >= 0 Then
'                    BALANCE = BALANCE + rs(0)
'                End If
'                rs.Close
'                    ' cash counter opening balance
'                rs.Open "select sum(NETAMOUNT) from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger  = '" & rsSLD(1) & "' and INVOICEDATE<cdate('" + Trim(date1.Text) + "')"
'                If rs(0) >= 0 Then
'                    BALANCE = BALANCE + rs(0)
'                End If
'                rs.Close
'                rs.Open "select sum(baa) from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger  = '" & rsSLD(1) & "' and INVOICEDATE<cdate('" + Trim(date1.Text) + "')"
'                If rs(0) >= 0 Then
'                    BALANCE = BALANCE - rs(0)
'                End If
'                rs.Close
'
'                ' Credit note Item opening balance
'                rs.Open "select sum(NETAMOUNT) from credita where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger  = '" & rsSLD(1) & "' and INVOICEDATE<cdate('" + Trim(date1.Text) + "')"
'                If rs(0) >= 0 Then
'                    BALANCE = BALANCE + rs(0)
'                End If
'                rs.Close
'                'credit note opening balance
'                rs.Open "select sum(NA) from cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld  = '" & rsSLD(1) & "' and cnd < cdate('" + Trim(date1.Text) + "') and DC='D'"
'                If rs(0) >= 0 Then
'                    BALANCE = BALANCE + rs(0)
'                End If
'                rs.Close
'                rs.Open "select sum(NA) from cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld  = '" & rsSLD(1) & "' and cnd < cdate('" + Trim(date1.Text) + "') and DC='C'"
'                If rs(0) >= 0 Then
'                    BALANCE = BALANCE - rs(0)
'                End If
'                rs.Close
'                rs.Open "select sum(A) from cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld  = '" & rsSLD(1) & "' and cnd < cdate('" + Trim(date1.Text) + "') and DC='D'"
'                If rs(0) >= 0 Then
'                    BALANCE = BALANCE + rs(0)
'                End If
'                rs.Close
'                rs.Open "select sum(A) from cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld  = '" & rsSLD(1) & "' and cnd < cdate('" + Trim(date1.Text) + "') and DC='C'"
'                If rs(0) >= 0 Then
'                    BALANCE = BALANCE - rs(0)
'                End If
'                rs.Close
'
'
'                'debit note opening balance
'                 rs.Open "select sum(NA) from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld  = '" & rsSLD(1) & "' and dnd < cdate('" + Trim(date1.Text) + "') and DC='D'"
'                If rs(0) >= 0 Then
'                    BALANCE = BALANCE + rs(0)
'                End If
'                rs.Close
'                rs.Open "select sum(NA) from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld  = '" & rsSLD(1) & "' and dnd < cdate('" + Trim(date1.Text) + "') and DC='C'"
'                If rs(0) >= 0 Then
'                    BALANCE = BALANCE - rs(0)
'                End If
'                rs.Close
'                rs.Open "select sum(A) from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld  = '" & rsSLD(1) & "' and dnd < cdate('" + Trim(date1.Text) + "') and DC='D'"
'                If rs(0) >= 0 Then
'                    BALANCE = BALANCE + rs(0)
'                End If
'                rs.Close
'                rs.Open "select sum(A) from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld  = '" & rsSLD(1) & "' and dnd < cdate('" + Trim(date1.Text) + "') and DC='C'"
'                If rs(0) >= 0 Then
'                    BALANCE = BALANCE - rs(0)
'                End If
'                rs.Close
'            ' balance field contain the opening amount upto the date
'                Set rs1 = New ADODB.Recordset
'                rs1.Open "treport", CON1, adopenStatic, adLockOptimistic, adCmdTable
'                rs1.AddNew
'                rs1!Text = "** Opening Balance as on " + Trim(date1.Text)
'                rs1!ad = BALANCE
'                rs1!period = Trim(date1.Text) + "  To  " + Trim(date2.Text)
'                rs1!header = "SUB. LEDGER ACCOUNT"
'                rs1!SUBLEDGER = Trim(Combosubledger.Text)
'                rs1!dorc = "D"
'                rs1.Update
'               'opening the tables
'                If rs.State = 1 Then rs.Close
'                rs.Open "select * from vouchers where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger  = '" & rsSLD(1) & "' and VoucherDate>=cdate('" + Trim(date1.Text) + "') and VoucherDate<=cdate('" + Trim(date2.Text) + "') order by  VoucherDate,vouchertype,vouchernumber", con, adopenStatic, adLockReadOnly, adCmdText
'                 If rs2.State = 1 Then rs2.Close
'                rs2.Open "select * from invoicea where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger  = '" & rsSLD(1) & "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') order by invoicedate", con, adopenStatic, adLockReadOnly, adCmdText
'                If rs3.State = 1 Then rs3.Close
'                rs3.Open "select * from CREDITa where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger  = '" & rsSLD(1) & "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "')", con, adopenStatic, adLockReadOnly, adCmdText
'                If rs4.State = 1 Then rs4.Close
'
'                rs4.Open "select * from Cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld  = '" & rsSLD(1) & "' and cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')", con, adopenStatic, adLockReadOnly, adCmdText
'                If rs5.State = 1 Then rs5.Close
'                rs5.Open "select * from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld  = '" & rsSLD(1) & "' and dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "')", con, adopenStatic, adLockReadOnly, adCmdText
'                If rs6.State = 1 Then rs6.Close
'                rs6.Open "select * from Cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld  = '" & rsSLD(1) & "' and cnd >= cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')", con, adopenStatic, adLockReadOnly, adCmdText
'                If rs7.State = 1 Then rs7.Close
'                rs7.Open "select * from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld  = '" & rsSLD(1) & "' and dnd >= cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "')", con, adopenStatic, adLockReadOnly, adCmdText
'                If rs8.State = 1 Then rs8.Close
'                rs8.Open "select * from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger  = '" & rsSLD(1) & "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') order by invoicedate", con, adopenStatic, adLockReadOnly, adCmdText
''  End If
'    'vouchers data start
'    If Not rs.BOF Then
'        Do While Not rs.EOF
'            rs1.AddNew
'            rs1!Genledger = rs!Genledger
'            rs1!SUBLEDGER = rs!SUBLEDGER
'            rs1!vdate = rs!VoucherDate
'            rs1!vtype = rs!vouchertype
'            rs1!vno = Trim(Str(rs!VoucherNumber))
'            If Len(Trim(rs!DESCRIPTION)) = 0 Then
'                rs1!narration = " "
'            Else
'                rs1!narration = Trim(rs!DESCRIPTION)
'            End If
'            If Trim(UCase(rs!DebitorCredit)) = Trim(UCase("D")) Then
'                rs1!ad = rs!amount
'                rs1!dorc = "D"
'           '     balance = balance + rs!amount
'            '    rs1!balance = balance
'            Else
'                rs1!aC = rs!amount
'                rs1!dorc = "C"
'             '   balance = balance - rs!amount
'              '  rs1!balance = balance
'            End If
'            If Len(Trim(rs!cbnd)) = 0 Then
'                rs1!cbno = " "
'            Else
'                rs1!cbno = rs!cbnd
'            End If
'            rs1.Update
'            If Not rs.EOF Then
'                rs.MoveNext
'            End If
'        Loop
'    End If
'    ' vouchers data end
'
'    'invoice data start
'    If Not rs2.BOF Then
'        Do While Not rs2.EOF
'            rs1.AddNew
'            rs1!Genledger = rs2!Genledger
'            rs1!SUBLEDGER = rs2!SUBLEDGER
'            rs1!vdate = rs2!INVOICEDATE
'            rs1!vtype = Trim("I")
'            rs1!vno = Trim(rs2!INVOICENO)
'            rs1!ad = rs2!netamount
'            rs1!narration = "Sales Invoice"
'            rs1!dorc = "D"
'           ' balance = balance + rs2!netamount
'           ' rs1!balance = balance
'            rs1.Update
'            If Not rs2.EOF Then
'                rs2.MoveNext
'            End If
'        Loop
'    End If
'    ' invoice data end
'
'    If Not rs3.BOF Then
'        Do While Not rs3.EOF
'            rs1.AddNew
'            rs1!Genledger = rs3!Genledger
'            rs1!SUBLEDGER = rs3!SUBLEDGER
'            rs1!vdate = rs3!INVOICEDATE
'            rs1!vtype = Trim("C")
'            rs1!vno = Trim(rs3!INVOICENO)
'            rs1!aC = rs3!netamount
'            rs1!narration = "Credit Note"
'            rs1!dorc = "C"
'           ' balance = balance - rs3!netamount
'           ' rs1!balance = balance
'            If Not rs3.EOF Then
'               rs3.MoveNext
'            End If
'            rs1.Update
'        Loop
'    End If
'    ' Credit not entry start
'    If Not rs4.EOF Then
'        Do While Not rs4.EOF
'            rs1.AddNew
'            rs1!Genledger = rs4!pgld
'            rs1!SUBLEDGER = rs4!psld
'            rs1!vdate = rs4!cnd
'            rs1!vtype = Trim("C")
'            rs1!vno = Trim(rs4!cnn)
'            rs1!aC = rs4!na
'            rs1!dorc = rs4!DC
'            rs1!narration = rs4!n
'            'balance = balance - rs4!na
'            'rs1!balance = balance
'            If Not rs4.EOF Then
'               rs4.MoveNext
'            End If
'            rs1.Update
'        Loop
'    End If
'  ' Debit note entry start
'    If Not rs5.EOF Then
'        Do While Not rs5.EOF
'            rs1.AddNew
'            rs1!Genledger = rs5!pgld
'            rs1!SUBLEDGER = rs5!psld
'            rs1!vdate = rs5!dnd
'            rs1!vtype = Trim("D")
'            rs1!vno = Trim(rs5!dnn)
'            rs1!ad = rs5!na
'            rs1!narration = rs5!n
'            rs1!dorc = rs5!DC
'            'balance = balance + rs5!na
'            'rs1!balance = balance
'            If Not rs5.EOF Then
'               rs5.MoveNext
'            End If
'            rs1.Update
'        Loop
'    End If
'    ' Credit notE  B entry start
'    If Not rs6.EOF Then
'        Do While Not rs6.EOF
'            rs1.AddNew
'            rs1!Genledger = rs6!gld
'            rs1!SUBLEDGER = rs6!Sld
'            rs1!vdate = rs6!cnd
'            rs1!vtype = Trim("C")
'            rs1!vno = Trim(rs6!cnn)
'            rs1!aC = rs6!a
'            rs1!dorc = rs6!DC
'            'balance = balance - rs4!na
'            'rs1!balance = balance
'            If Not rs6.EOF Then
'               rs6.MoveNext
'            End If
'            rs1.Update
'        Loop
'    End If
'  ' Debit note B entry start
'    If Not rs7.EOF Then
'        Do While Not rs7.EOF
'            rs1.AddNew
'            rs1!Genledger = rs7!gld
'            rs1!SUBLEDGER = rs7!Sld
'            rs1!vdate = rs7!dnd
'            rs1!vtype = Trim("D")
'            rs1!vno = Trim(rs7!dnn)
'            rs1!ad = rs7!a
'            rs1!dorc = rs7!DC
'            'balance = balance + rs5!na
'            'rs1!balance = balance
'            If Not rs7.EOF Then
'               rs7.MoveNext
'            End If
'            rs1.Update
'        Loop
'    End If
'    'cash counter sales data start
'    If Not rs8.BOF Then
'        Do While Not rs8.EOF
'            rs1.AddNew
'            rs1!Genledger = rs8!Genledger
'            rs1!SUBLEDGER = rs8!SUBLEDGER
'            rs1!vdate = rs8!INVOICEDATE
'            rs1!vtype = Trim("I")
'            rs1!vno = Trim(rs8!INVOICENO)
'            rs1!ad = rs8!netamount
'            rs1!narration = "Counter Sales Invoice"
'            rs1!dorc = "D"
'           ' balance = balance + rs8!netamount
'           ' rs1!balance = balance
'            rs1.Update
'            ' for cash recept
'            rs1.AddNew
'            rs1!Genledger = rs8!Genledger
'            rs1!SUBLEDGER = rs8!SUBLEDGER
'            rs1!vdate = rs8!INVOICEDATE
'            rs1!vtype = Trim("I")
'            rs1!vno = Trim(rs8!INVOICENO)
'            rs1!aC = rs8!baa
'            rs1!narration = "Cash receipt"
'            rs1!dorc = "C"
'           ' balance = balance + rs8!netamount
'           ' rs1!balance = balance
'            rs1.Update
'            If Not rs8.EOF Then
'                rs8.MoveNext
'            End If
'        Loop
'    End If
'    rs1.Close
'
'     rsSLD.MoveNext
'
'
'    Wend
' End If
'    ' cash counter  data end
'
'    ' making the balance in the output file
'    rs1.Open "select * from treport order by vdate, vtype,vno", CON1, adopenStatic, adLockOptimistic, adCmdText
'    If Not rs1.BOF Then
'            rs1.MoveFirst
'            BALANCE = 0
'            Do While Not rs1.EOF
'                BALANCE = BALANCE + rs1!ad - rs1!aC
'                rs1!BALANCE = BALANCE
'                rs1.Update
'                If Not rs1.EOF Then
'                    rs1.MoveNext
'                End If
'             Loop
'    End If
'    If rs1.State = 1 Then
'        rs1.Close
'    End If
'    'rs1.Open "treport", CON1, adopenStatic, adLockReadOnly, adCmdTable
'    'rs1.Close
'    If rs.State = 1 Then
'        rs.Close
'    End If
'    main.reportname = "Sub. Ledger"
'    viewledger.Top = 0
'    viewledger.Left = 0
'    viewledger.Show
'Else
'    MsgBox "gen. ledger or sub. ledger not selected"
'End If
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
Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu"))
    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
        Unload VB.Screen.ActiveForm
    End If
Loop
Me.Top = 0
Me.Left = 0
'Set CON = New ADODB.Connection
Set rs = New ADODB.Recordset
'    With CON
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
'        .Open
'    End With
    rs.Open "select * from gledger where slf=true", CON, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            COMBOGENLEDGER.AddItem Trim(rs!gledger)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    rs.Open "setup1", CON, adOpenStatic, adLockReadOnly, adCmdTable
    CNSetup
    date1.Text = rs!yarfrom
    date2.Text = rs!yarto
    rs.Close
End Sub
Sub xx()

''    Else
''
''                ' opening Balance Start
''                rs.Open "select * from sledger where gledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "'", con, adopenStatic, adLockReadOnly, adCmdText
''                BALANCE = Round(rs!YEAROPENING, 2)
''                rs.Close
''                ' vouchers opening balance
''                rs.Open "select sum(amount) from vouchers where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and VoucherDate<cdate('" + Trim(date1.Text) + "') and DebitorCredit='D'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(amount) from vouchers where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and VoucherDate<cdate('" + Trim(date1.Text) + "') and DebitorCredit='C'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE - rs(0)
''                End If
''                rs.Close
''                ' invoice opening balance
''                rs.Open "select sum(NETAMOUNT) from invoicea where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and INVOICEDATE<cdate('" + Trim(date1.Text) + "')"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                    ' cash counter opening balance
''                rs.Open "select sum(NETAMOUNT) from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and INVOICEDATE<cdate('" + Trim(date1.Text) + "')"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(baa) from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and INVOICEDATE<cdate('" + Trim(date1.Text) + "')"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE - rs(0)
''                End If
''                rs.Close
''
''                ' Credit note Item opening balance
''                rs.Open "select sum(NETAMOUNT) from credita where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and INVOICEDATE<cdate('" + Trim(date1.Text) + "')"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                'credit note opening balance
''                rs.Open "select sum(NA) from cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') and DC='D'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(NA) from cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') and DC='C'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE - rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(A) from cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') and DC='D'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(A) from cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') and DC='C'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE - rs(0)
''                End If
''                rs.Close
''
''
''                'debit note opening balance
''                 rs.Open "select sum(NA) from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "') and DC='D'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(NA) from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "') and DC='C'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE - rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(A) from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "') and DC='D'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE + rs(0)
''                End If
''                rs.Close
''                rs.Open "select sum(A) from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "') and DC='C'"
''                If rs(0) >= 0 Then
''                    BALANCE = BALANCE - rs(0)
''                End If
''                rs.Close
''            ' balance field contain the opening amount upto the date
''                Set rs1 = New ADODB.Recordset
''                CON1.Execute ("delete * from treport")
''                rs1.Open "treport", CON1, adopenStatic, adLockOptimistic, adCmdTable
''                rs1.AddNew
''                rs1!Text = "** Opening Balance as on " + Trim(date1.Text)
''                rs1!ad = BALANCE
''                rs1!period = Trim(date1.Text) + "  To  " + Trim(date2.Text)
''                rs1!header = "SUB. LEDGER ACCOUNT"
''                rs1!subledger = Trim(Combosubledger.Text)
''                rs1!dorc = "D"
''                rs1.Update
''               'opening the tables
''                rs.Open "select * from vouchers where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and VoucherDate>=cdate('" + Trim(date1.Text) + "') and VoucherDate<=cdate('" + Trim(date2.Text) + "') order by  VoucherDate,vouchertype,vouchernumber", con, adopenStatic, adLockReadOnly, adCmdText
''                rs2.Open "select * from invoicea where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') order by invoicedate", con, adopenStatic, adLockReadOnly, adCmdText
''
''                rs3.Open "select * from CREDITa where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "')", con, adopenStatic, adLockReadOnly, adCmdText
''                rs4.Open "select * from Cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')", con, adopenStatic, adLockReadOnly, adCmdText
''                rs5.Open "select * from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "')", con, adopenStatic, adLockReadOnly, adCmdText
''                rs6.Open "select * from Cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')", con, adopenStatic, adLockReadOnly, adCmdText
''                rs7.Open "select * from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "')", con, adopenStatic, adLockReadOnly, adCmdText
''                rs8.Open "select * from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') order by invoicedate", con, adopenStatic, adLockReadOnly, adCmdText
''
''

End Sub














Sub OPENINGSUBLEDGERS()

If Trim(alpha.Text) <> "" Then
        'CON.Execute "Insert into subledgertrail  SELECT SUBLEDGER , YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId FROM SLEDGER where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' AND subledger like '" + Trim(alpha.Text) + "%'", p, adCmdText
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER))" _
        & " where  gledger='" + Trim(COMBOGENLEDGER.Text) + "'and INVOICEDATE < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER "
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT,  " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER)) " _
        & " where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'  and INVOICEDATE< cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; "
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT ," & UId & " as UserId  " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER))" _
        & " where  gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'   and INVOICEDATE < cdate('" + Trim(date1.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; "
        
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER)) " _
        & " where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; "
        
        
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT ,  " & UId & " as UserId" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE DEBITORCREDIT = 'D' and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE < cdate('" + Trim(date1.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;"
        
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE DEBITORCREDIT = 'C' and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; "
         ''''ok
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "')AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail     SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') AND SLEDGER.SUBLEDGER like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail     SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.Text) + "'And dnd < cdate('" + Trim(date1.Text) + "')  AND SLEDGER.SUBLEDGER like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId  " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "') AND SLEDGER.SUBLEDGER like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
        
        'CON.Execute "Delete * from TemprptTrialBalance"
        'CON.Execute "insert into TemprptTrialBalance ( Subledger,openingbalance,userid ) SELECT  SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId from subledgertrail GROUP BY SUBLEDGER;"
    
  End If
          
          



If Trim(alpha.Text) = "" And Combosubledger.Text = "" Then


        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER))  " _
        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "'and INVOICEDATE < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER)) " _
        & " where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' and  SLEDGER.SUBLEDGER <>'CASH PARTY'    and INVOICEDATE< cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER))" _
        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY' and INVOICEDATE < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER)) " _
        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
        
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
         ''''ok
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "')  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and cnd < cdate('" + Trim(date1.Text) + "')  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.Text) + "'And dnd < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.Text) + "' and dnd < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
  End If
  
If Trim(alpha.Text) = "" And Combosubledger.Text <> "" Then

        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER))  " _
        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER=  '" & Combosubledger.Text & "' and INVOICEDATE < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER)) " _
        & " where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER=  '" & Combosubledger.Text & "'    and INVOICEDATE< cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER))" _
        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER ='" & Combosubledger.Text & "' and INVOICEDATE < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER)) " _
        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER='" & Combosubledger.Text & "' and INVOICEDATE < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
        
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER= '" & Combosubledger.Text & "' and VOUCHERDATE < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER= '" & Combosubledger.Text & "' and VOUCHERDATE < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
         ''''ok
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "'and CNF1A.PSLD = '" & Combosubledger.Text & "'  and cnd < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and CNF1A.PSLD = '" & Combosubledger.Text & "'  and cnd < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and   DNFA.PSLD = '" & Combosubledger.Text & "' and dnd < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.Text) + "'  and  DNFA.PSLD = '" & Combosubledger.Text & "'   and dnd < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and  CNF1B.SLD = '" & Combosubledger.Text & "'    and cnd < cdate('" + Trim(date1.Text) + "')  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "'  and  CNF1B.SLD = '" & Combosubledger.Text & "'      and cnd < cdate('" + Trim(date1.Text) + "')  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.Text) + "'  and  DNFB.SLD = '" & Combosubledger.Text & "'     And dnd < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.Text) + "'  and  DNFB.SLD = '" & Combosubledger.Text & "'    and dnd < cdate('" + Trim(date1.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
  End If
  
  
  
  
  
  
  
  
  
End Sub









