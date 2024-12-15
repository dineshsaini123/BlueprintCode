VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FRMOPTRANS 
   Caption         =   "Opening transfer"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   9195
   WindowState     =   2  'Maximized
   Begin VB.ListBox Combosubledger 
      Appearance      =   0  'Flat
      Height          =   1605
      Left            =   2550
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   630
      Width           =   4665
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   2550
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   240
      Width           =   4665
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Subledger A/c"
      Height          =   405
      Left            =   2460
      TabIndex        =   8
      Top             =   4620
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   465
      Left            =   2430
      TabIndex        =   7
      Top             =   4110
      Width           =   3180
   End
   Begin VB.TextBox Alpha 
      Height          =   315
      Left            =   2550
      MaxLength       =   1
      TabIndex        =   6
      Top             =   990
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Crystal Report"
      Height          =   315
      Left            =   2535
      TabIndex        =   5
      Top             =   3030
      Width           =   1830
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Dos Report"
      Height          =   360
      Left            =   5010
      TabIndex        =   4
      Top             =   2985
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "All"
      Height          =   255
      Left            =   6450
      TabIndex        =   2
      Top             =   2430
      Width           =   795
   End
   Begin VB.CommandButton cmdClosing 
      Caption         =   "&Ledger Closing A/c"
      Height          =   405
      Left            =   2445
      TabIndex        =   1
      Top             =   3660
      Width           =   3165
   End
   Begin VB.CheckBox closingCheck 
      Caption         =   "For Closing Tranf. To  Next Session"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5790
      TabIndex        =   0
      Top             =   3630
      Width           =   2115
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
      Left            =   2550
      TabIndex        =   10
      Top             =   2460
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
      Left            =   5100
      TabIndex        =   11
      Top             =   2460
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Gen. Ledger Desc."
      Height          =   285
      Left            =   450
      TabIndex        =   16
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Sub. Ledger Desc."
      Height          =   285
      Left            =   450
      TabIndex        =   15
      Top             =   630
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Alphabat"
      Height          =   285
      Left            =   450
      TabIndex        =   14
      Top             =   990
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label Label3 
      Caption         =   " - To - "
      Height          =   315
      Left            =   4200
      TabIndex        =   13
      Top             =   2490
      Width           =   585
   End
   Begin VB.Label Label2 
      Caption         =   "From The Date"
      Height          =   315
      Left            =   360
      TabIndex        =   12
      Top             =   2490
      Width           =   1995
   End
End
Attribute VB_Name = "FRMOPTRANS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As Connection
Dim sum1 As Double
 Dim K As Integer
    Dim arr() As String
        Dim cust As String
Dim RS As Recordset

Private Sub Check1_Click()
For I = 0 To Combosubledger.ListCount - 1
   Combosubledger.Selected(I) = False
Next


If Check1.value = 1 Then

For I = 0 To Combosubledger.ListCount - 1
   Combosubledger.Selected(I) = True
Next

End If

End Sub

Private Sub cmdClosing_Click()
slcosing = 1
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
CON.Execute "DELETE from subledgertrail where  " & stringyear & " and len(subledger)>0"
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
Exit Sub
aa:
MsgBox "" & Err.DESCRIPTION
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
    CON.Execute ("DELETE from treport where len(genledger)>0")
    
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
    


If Trim(Alpha.Text) = "" And Alpha.Visible = False And cust <> "" Then

For I = 0 To K - 1
cust = arr(I)
If arr(I) <> "" Then
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)   SELECT GenLedger, SubLedger, VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & ",'" & main.session & "'," & main.setupid & " From VOUCHERS Where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'  AND   VOUCHERS.DebitorCredit='C' and  convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber"
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM INVOICEA  where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'  and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
                'change By Dinesh
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT Purchasea.GENLEDGER, Purchasea.SUBLEDGER, Purchasea.INVOICEDATE, 'I' AS Expr1, Purchasea.billNO, 'Purchase Invoice' , Purchasea.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM Purchasea  where    " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,fyear,setupid)  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'   and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   "
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid)  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & ",'" & main.session & "'," & main.setupid & "  FROM CASHA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'   and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND CASHA.BAA <>0  "
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )    SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & ",'" & main.session & "'," & main.setupid & " FROM CREDITA  where    " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger = '" & cust & "'    and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) "
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,fyear,setupid )   SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & "   From DNFA  where    " & stringyear & " and Pgld ='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld = '" & cust & "' and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid,fyear,setupid )   SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1A where  " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld = '" & cust & "'    and convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid ,fyear,setupid)   SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, ''     , " & UId & ",'" & main.session & "'," & main.setupid & " From DNFB  where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld = '" & cust & "'   and convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
                CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid ,fyear,setupid)   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & ",'" & main.session & "'," & main.setupid & " From CNF1B where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld = '" & cust & "' and  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
End If
Next
cust = "a"
End If


CON.Execute "insert into Treport ( Genledger,Subledger,openingbalance,userid,fyear,setupid ) SELECT '" + Trim(COMBOGENLEDGER.Text) + "' as genled, SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId,'" & main.session & "'," & main.setupid & " from subledgertrail where  " & stringyear & " GROUP BY SUBLEDGER;"

dr = 0
cr = 0
op = 0



If closingCheck.value = 1 Then

Dim fyear1, setupid As String

If RS.State = 1 Then RS.close
RS.Open "select fyear,setupid from financialyear where setupid=" & main.setupid & " order by fyear,setupid", CON, adOpenKeyset, adLockReadOnly
While RS.EOF = False
   fyear1 = RS(0)
   setupid = RS(1)
   RS.MoveNext
Wend



If RS.State = 1 Then RS.close
RS.Open "select subledger from [ExportData].[dbo].[treport] where " & stringyear & " group by subledger", CON, adOpenKeyset, adLockReadOnly
While RS.EOF = False

cr = 0
dr = 0
sum1 = 0
op = 0

If rs1.State = 1 Then rs1.close
rs1.Open "select subledger,OpeningBalance,ad,dorc from [ExportData].[dbo].[treport] where " & stringyear & " and subledger='" & RS(0) & "'", CON, adOpenKeyset, adLockReadOnly
While rs1.EOF = False
    If rs1!dorc = "D" Then
      dr = dr + rs1!aD
    End If
    If rs1!dorc = "C" Then
      cr = cr + rs1!aD
    End If
    If rs1!OpeningBalance <> 0 Then
    op = rs1!OpeningBalance
    End If
rs1.MoveNext
Wend

sum1 = op + (dr - cr)


CON.Execute "update sledger set YEAROPENING = " & sum1 & "  Where  setupid=" & setupid & " and fyear='" & fyear1 & "' and SUBLEDGER='" & RS!SUBLEDGER & "'"
RS.MoveNext
Wend


'------------------------------------------
MsgBox "Opening Transfer ..", vbInformation
Exit Sub
End If


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
    '''MainMenu.Toolbar1.Visible = True
    Unload Me
    Unload PrintOption
End Sub
Private Sub Commandshow_Click()

On Error GoTo aa:

'********sub for alpha wise and Partywise according to new fast mathed
slcosing = 0

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
CON.Execute "DELETE from subledgertrail where  " & stringyear & ""
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
Exit Sub
aa:
MsgBox "" & Err.DESCRIPTION

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

Option1.value = True


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
Sub OPENINGSUBLEDGERS()
  
If Trim(Alpha.Text) = "" And cust <> "" Then
    K = 0
    Dim k1 As Integer
    k1 = 0
    ReDim Preserve arr(Me.Combosubledger.ListCount)
    arr(K) = ""
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
        & " where sledger.fyear='" & main.session & "' and invoicea.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER=  '" & cust & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  0 AS OPAMOUNTDEBIT,sum(Purchasea.NETAMOUNT) AS OPAMOUNTCREDIT  , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN purchasea ON (SLEDGER.SUBLEDGER = purchasea.SUBLEDGER) AND (SLEDGER.gledger = purchasea.GENLEDGER) AND (SLEDGER.fyear = purchasea.fyear) AND (SLEDGER.setupid = purchasea.setupid))  " _
        & " where  sledger.fyear='" & main.session & "' and purchasea.setupid=" & main.setupid & " and  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER=  '" & cust & "' and convert(smalldatetime,INVOICEDATE,103) <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " & _
        " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText



        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid)) " _
        & " where   sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & "  and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER=  '" & cust & "' and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText

        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))" _
        & " where   sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & "  and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER ='" & cust & "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
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

  End If
  
End Sub



