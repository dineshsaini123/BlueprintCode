VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form subtrial 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5805
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "strial.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      Caption         =   "Crystal Report"
      Height          =   435
      Left            =   3120
      TabIndex        =   16
      Top             =   2340
      Width           =   1470
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Dos report"
      Height          =   495
      Left            =   4740
      TabIndex        =   15
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5790
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton export 
      Height          =   345
      Left            =   1470
      Picture         =   "strial.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4740
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "strial.frx":045D
      Left            =   2370
      List            =   "strial.frx":0470
      TabIndex        =   10
      Text            =   "100 %"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton print1 
      Height          =   345
      Left            =   1920
      Picture         =   "strial.frx":0496
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4740
      Visible         =   0   'False
      Width           =   345
   End
   Begin RichTextLib.RichTextBox r1 
      Height          =   4095
      Left            =   840
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   7223
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      MaxLength       =   99999999
      RightMargin     =   20000
      TextRTF         =   $"strial.frx":0608
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
      Left            =   4890
      TabIndex        =   5
      Top             =   3060
      Width           =   1545
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      Height          =   405
      Left            =   3120
      TabIndex        =   4
      Top             =   3060
      Width           =   1545
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   3060
      TabIndex        =   0
      Top             =   1140
      Width           =   3885
   End
   Begin MSMask.MaskEdBox alpha 
      Height          =   345
      Left            =   7650
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   1
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin MSComDlg.CommonDialog c1 
      Left            =   300
      Top             =   5190
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   3030
      TabIndex        =   2
      Top             =   1650
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
      Left            =   5520
      TabIndex        =   3
      Top             =   1650
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
      Left            =   780
      TabIndex        =   14
      Top             =   1710
      Width           =   1995
   End
   Begin VB.Label Label3 
      Caption         =   " - To - "
      Height          =   315
      Left            =   4620
      TabIndex        =   13
      Top             =   1680
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Alphabat"
      Height          =   195
      Left            =   7440
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Gen. Ledger Desc."
      Height          =   195
      Left            =   960
      TabIndex        =   6
      Top             =   1170
      Width           =   1350
   End
End
Attribute VB_Name = "subtrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As Connection
Dim Q1 As String
Dim RS As Recordset

Private Sub alpha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub COMBOGENLEDGER_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub COMBOGENLEDGER_LostFocus()
    If Trim(COMBOGENLEDGER.Text) <> "" Then
       If RS.State = 1 Then RS.Close
    
        
                RS.Open "select * from gledger where  " & stringyear & " and slf=1", CON, adOpenDynamic, adLockReadOnly, adCmdText
        If Not RS.BOF Then
            RS.Find "gledger='" + Trim(COMBOGENLEDGER.Text) + "'"
            If RS.EOF Then
                COMBOGENLEDGER.SetFocus
            End If
        Else
            COMBOGENLEDGER.SetFocus
        End If
        RS.Close
    End If
End Sub

Private Sub Command1_Click()
Unload Me
''MainMenu.Toolbar1.Visible = True
End Sub


Sub SUBLEDGERSBALANCE()
       
If Trim(Alpha.Text) <> "" Then


CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 as YEAROPENING,  Sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER) AND (SLEDGER.fyear = INVOICEA.fyear) AND (SLEDGER.setupid = INVOICEA.setupid))" _
        & " where sledger.fyear='" & main.session & "' and invoicea.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER; "

                
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))" _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'  and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; "
        
        
        CON.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))" _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'  and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; "
   
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER) AND (SLEDGER.fyear = credita.fyear) AND (SLEDGER.setupid = credita.setupid))" _
        & " where sledger.fyear='" & main.session & "' and credita.setupid=" & main.setupid & " and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; "
  
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid)" _
        & " where sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;"
        
   
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid)" _
        & " where sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'C' and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; "
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid)" _
        & " where sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='D' and gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,cnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER  " _
        & " HAVING  Sum(CNF1A.NA) <> 0; "

        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid)" _
        & " where sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='C' and gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; "
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid)" _
        & " where sledger.fyear='" & main.session & "' and dnfa.setupid=" & main.setupid & " and DNFA.DC='D' and gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER" _
        & " HAVING  Sum(DNFA.NA) <> 0; "
        

        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid)" _
        & " where sledger.fyear='" & main.session & "' and dnfa.setupid=" & main.setupid & " and DNFA.DC='C' and gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; "

        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid)" _
        & " where sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC='D' and gledger  = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,cnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; "
        

        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD)  AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid)" _
        & " where sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC='C' and gledger  = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;"

        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid)" _
        & " where sledger.fyear='" & main.session & "' and dnfb.setupid=" & main.setupid & " and DNFB.DC='D' and gledger  = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; "
        

        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD) AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid)" _
        & " where sledger.fyear='" & main.session & "' and dnfb.setupid=" & main.setupid & " and DNFB.DC='C' and gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; "


Else
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 as YEAROPENING,  Sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER) AND (SLEDGER.fyear = INVOICEA.fyear) AND (SLEDGER.setupid = INVOICEA.setupid))" _
        & " where sledger.fyear='" & main.session & "' and invoicea.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER; "
        
       
       If Trim(COMBOGENLEDGER.Text) = "SALES" Then
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 as YEAROPENING,  Sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER) AND (SLEDGER.fyear = INVOICEA.fyear) AND (SLEDGER.setupid = INVOICEA.setupid))" _
        & " where sledger.fyear='" & main.session & "' and invoicea.setupid=" & main.setupid & " and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER; "
       
       
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 as YEAROPENING,  Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))" _
        & " where CurrencyValue>0 and sledger.fyear='" & main.session & "' and CASHA.setupid=" & main.setupid & " and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER; "
      
       End If

        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))" _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "'  and SLEDGER.SUBLEDGER <>'CASH PARTY' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; "
        

        CON.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))" _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "'and SLEDGER.SUBLEDGER <>'CASH PARTY'  and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; " _
        
   
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER) AND (SLEDGER.fyear = credita.fyear) AND (SLEDGER.setupid = credita.setupid))" _
        & " where sledger.fyear='" & main.session & "' and credita.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; " _
        
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid)" _
        & " where sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;"
   
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid)" _
        & " where sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; "
   
        CON.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid)" _
        & " where sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='D' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,cnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER  " _
        & " HAVING  Sum(CNF1A.NA) <> 0; "
        
 
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) and  (SLEDGER.fyear = cnf1a.fyear)" _
        & " where sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='C' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; "
        
        
     
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid)" _
        & " where sledger.fyear='" & main.session & "' and dnfa.setupid=" & main.setupid & " and DNFA.DC='D' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER" _
        & " HAVING  Sum(DNFA.NA) <> 0; "
        
 
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid)" _
        & " where sledger.fyear='" & main.session & "' and DNFA.setupid=" & main.setupid & " and DNFA.DC='C' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; "
        
  
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid)" _
        & " where sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC='D' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,cnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; "
        
    
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid)" _
        & " where sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC='C' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;"
        
 
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid)" _
        & " where sledger.fyear='" & main.session & "' and dnfb.setupid=" & main.setupid & " and DNFB.DC='D' and gld = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; "
        
   
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD) AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid)" _
        & " where sledger.fyear='" & main.session & "' and dnfb.setupid=" & main.setupid & " and DNFB.DC='C' AND gld = '" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; "
        
        
   End If
   DoEvents
   CON.Execute "insert into TemprptTrialBalance ( Subledger, Damount,CAmount,userid,FYEAR,SETUPID) SELECT  SUBLEDGER,  SUM (OPAMOUNTDEBIT) as Damount,  SUM(OPAMOUNTCREDIT)as Camount,userid,'" & main.session & "'," & main.setupid & "  from subledgertrail where  " & stringyear & " and userid=" & main.UId & " GROUP BY SUBLEDGER,userid;"

End Sub


Sub OPENINGSUBLEDGERS()
CON.Execute "Delete  from subledgertrail where " & stringyear & " and userid=" & main.UId
If Trim(Alpha.Text) <> "" Then
          
     ' This Code is Not Use
    
     CON.Execute "Insert into subledgertrail  SELECT SUBLEDGER , YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & " FROM SLEDGER where  " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' AND subledger like '" + Trim(Alpha.Text) + "%'", p, adCmdText
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER) AND (SLEDGER.fyear = INVOICEA.fyear) AND (SLEDGER.setupid = INVOICEA.setupid))  " _
        & " where sledger.fyear='" & main.session & "' and invoicea.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "'and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT,  " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))  " _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'  and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT ," & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))  " _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'   and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER) AND (SLEDGER.fyear = credita.fyear) AND (SLEDGER.setupid = credita.setupid))  " _
        & " where sledger.fyear='" & main.session & "' and credita.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
        
        
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT ,  " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid) " _
        & " where sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'D' and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid) " _
        & " where sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'C' and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
         ''''ok
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid) " _
        & " where sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='D' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid) " _
        & " where sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='C' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid) " _
        & " where sledger.fyear='" & main.session & "' and dnfa.setupid=" & main.setupid & " and DNFA.DC='D' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid) " _
        & " where sledger.fyear='" & main.session & "' and dnfa.setupid=" & main.setupid & " and DNFA.DC='C' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail     SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid) " _
        & " where sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC='D' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid) " _
        & " where sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC='C' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.SUBLEDGER like '" + Trim(Alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail     SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid) " _
        & " where sledger.fyear='" & main.session & "' and dnfb.setupid=" & main.setupid & " and DNFB.DC='D' and gld = '" + Trim(COMBOGENLEDGER.Text) + "'And convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  AND SLEDGER.SUBLEDGER like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD) AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid) " _
        & " where sledger.fyear='" & main.session & "' and dnfb.setupid=" & main.setupid & " and DNFB.DC= 'C' and gld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) AND SLEDGER.SUBLEDGER like '" + Trim(Alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
        
        CON.Execute "DELETE from TemprptTrialBalance where  " & stringyear & " and userid=" & main.UId
        CON.Execute "insert into TemprptTrialBalance ( Subledger,openingbalance,userid,fyear,setupid ) SELECT  SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId,'" & main.session & "'," & main.setupid & " from subledgertrail where  " & stringyear & " and userid=" & main.UId & " GROUP BY SUBLEDGER;"
    
    
Else
  
    CON.Execute "Insert into subledgertrail SELECT SUBLEDGER , YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT, " & UId & " as UserId,'" & main.session & "'," & main.setupid & "  FROM SLEDGER where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "'", p, adCmdText

        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER) AND (SLEDGER.fyear = INVOICEA.fyear) AND (SLEDGER.setupid = INVOICEA.setupid))  " _
        & " where sledger.fyear='" & main.session & "' and invoicea.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "'and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
        
       'Change By Dinesh
       
       If Trim(COMBOGENLEDGER.Text) = "SALES" Then
            
            CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
            & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER) AND (SLEDGER.fyear = INVOICEA.fyear) AND (SLEDGER.setupid = INVOICEA.setupid))  " _
            & " where sledger.fyear='" & main.session & "' and invoicea.setupid=" & main.setupid & " and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
            & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
       
       
            CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
            & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))  " _
            & " where sledger.fyear='" & main.session & "' and CASHA.setupid=" & main.setupid & " and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
            & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText

       
       End If
        
        
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))  " _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and  SLEDGER.SUBLEDGER <>'CASH PARTY'    and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER) AND (SLEDGER.fyear = CASHA.fyear) AND (SLEDGER.setupid = CASHA.setupid))  " _
        & " where sledger.fyear='" & main.session & "' and sledger.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER) AND (SLEDGER.fyear = credita.fyear) AND (SLEDGER.setupid = credita.setupid))  " _
        & " where sledger.fyear='" & main.session & "' and credita.setupid=" & main.setupid & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE <convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid) " _
        & " where sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) AND (SLEDGER.fyear = vouchers.fyear) AND (SLEDGER.setupid = vouchers.setupid) " _
        & " where sledger.fyear='" & main.session & "' and vouchers.setupid=" & main.setupid & " and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
         ''''ok
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid) " _
        & " where sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='D' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) AND (SLEDGER.fyear = cnf1a.fyear) AND (SLEDGER.setupid = cnf1a.setupid) " _
        & " where sledger.fyear='" & main.session & "' and cnf1a.setupid=" & main.setupid & " and CNF1A.DC='C' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid) " _
        & " where sledger.fyear='" & main.session & "' and dnfa.setupid=" & main.setupid & " and DNFA.DC='D' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) AND (SLEDGER.fyear = dnfA.fyear) AND (SLEDGER.setupid = dnfA.setupid) " _
        & " where sledger.fyear='" & main.session & "' and dnfa.setupid=" & main.setupid & " and DNFA.DC='C' and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid) " _
        & " where sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC='D' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) AND (SLEDGER.fyear = cnf1b.fyear) AND (SLEDGER.setupid = cnf1b.setupid) " _
        & " where sledger.fyear='" & main.session & "' and cnf1b.setupid=" & main.setupid & " and CNF1B.DC='C' and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,cnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid) " _
        & " where sledger.fyear='" & main.session & "' and dnfb.setupid=" & main.setupid & " and DNFB.DC='D' and gld = '" + Trim(COMBOGENLEDGER.Text) + "'And convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD) AND (SLEDGER.fyear = dnfb.fyear) AND (SLEDGER.setupid = dnfb.setupid) " _
        & " where sledger.fyear='" & main.session & "' and dnfb.setupid=" & main.setupid & " and DNFB.DC= 'C' and gld = '" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,dnd,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
        CON.Execute "DELETE from TemprptTrialBalance where  " & stringyear & " and userid=" & main.UId, p, adCmdText
        CON.Execute "insert into TemprptTrialBalance (Subledger,openingbalance,userid,fyear,setupid) SELECT  SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR , " & UId & " as UserId,'" & main.session & "'," & main.setupid & "   from subledgertrail where  " & stringyear & " and userid=" & main.UId & " GROUP BY SUBLEDGER;", p, adCmdText
  
  End If
  
  Exit Sub
End Sub


Private Sub Commandreturn_Click()
''MainMenu.Toolbar1.Visible = True
Unload Me
End Sub
Private Sub Commandshow_Click()
If Trim(COMBOGENLEDGER.Text) <> "" Then
    If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
        MsgBox "invalid date", , title1
        Exit Sub
    End If
    
    CON.Execute "DELETE FROM TemprptTrialBalance where  " & stringyear & ""
    CON.Execute "DELETE FROM subledgertrail where  " & stringyear & ""
    OPENINGSUBLEDGERS
    CON.Execute "Delete from subledgertrail where  " & stringyear & ""
    SUBLEDGERSBALANCE
    Genrate
    Unload PrintOption
    Load PrintOption
    
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
Me.Option1.Value = True

Me.r1.TOP = 10
Me.r1.Left = 10
'''''Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu"))
'''''    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
'''''        Unload VB.Screen.ActiveForm
'''''    End If
'''''Loop


Me.TOP = 0
Me.Left = 0
'CON.Execute "DELETE FROM TemprptTrialBalance"
'CON.Execute "DELETE FROM subledgertrail"
'con.Execute "DELETE FROM TemprptTrialBalance"
'con.Execute "DELETE FROM subledgertrail"

'Set CON = New ADODB.Connection
'CON.CursorLocation = adUseClient
Set RS = New ADODB.Recordset

 '   With CON
  '      .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\data.mdb"
  '      .Open
   ' End With
    'rs.Open "select * from gledger where slf=true", CON, adOpenDynamic, adLockReadOnly, adCmdText
       RS.Open "select * from gledger where  " & stringyear & " and slf=1", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        Do While Not RS.EOF
            COMBOGENLEDGER.AddItem Trim(RS!gledger)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    

    RS.Close
    CNSetup
    RS.Open "Select * from setup where " & stringyear & "", CON, adOpenDynamic, adLockReadOnly, adCmdText
    date1.Text = RS!yarfrom
    date2.Text = RS!yarto
    RS.Close
End Sub

Private Sub print_Click()
        
        Rsinvoicea.Open "select GenLedger,  SubLedger , sum(amount) as INVAmount from invoicea where  " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) order by invoicedate", CON, adOpenDynamic, adLockReadOnly, adCmdText
        
        RsCREDITa.Open "select * from CREDITa where  " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenDynamic, adLockReadOnly, adCmdText
        RsCnf1a.Open "select * from Cnf1a where  " & stringyear & " and pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,cnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenDynamic, adLockReadOnly, adCmdText
        Rsdnfa.Open "select * from dnfa where  " & stringyear & " and pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenDynamic, adLockReadOnly, adCmdText
        RsCnf1B.Open "select * from Cnf1B where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,cnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenDynamic, adLockReadOnly, adCmdText
        RsdnfB.Open "select * from dnfB where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenDynamic, adLockReadOnly, adCmdText
        RScasha.Open "select * from casha where  " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) order by invoicedate", CON, adOpenDynamic, adLockReadOnly, adCmdText
        
  
End Sub


Sub Genrate()
    Command1.TOP = r1.TOP + r1.Height + 30
    Combo1.TOP = r1.TOP + r1.Height + 30
    Set RS = New ADODB.Recordset
    main.reportname = "Sub. Ledger Trial"
    Dim called1, called2 As Boolean
    Dim MaxLine As Integer
    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
    Dim paperWidth As Integer
    Dim xtemp As String
    Dim trs As ADODB.Recordset
    Set trs = New ADODB.Recordset
        paperWidth = 150
        T1 = 10
        T2 = 25
        T3 = 40
        T4 = 50
        T5 = 70
        T6 = 85
        T7 = 100
        T8 = 115
        MaxLine = 72
        called1 = False
        called2 = False
        Dim Line As Integer
        Dim rs1 As ADODB.Recordset
        Dim kkk As ADODB.Recordset
        Dim Pno As Integer
        Dim FooterYes As Boolean
        Dim GopenBal As Double
        Dim GopenDr As Double
        Dim GopenCr As Double
        Dim GopenCl As Double
        GopenBal = 0
        GopenDr = 0
        GopenCr = 0
        GopenCl = 0
        FooterYes = False
        Set kkk = New ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        main.reportdata
        main.repors.Find "reportname='" + Trim(main.reportname) + "'"
        If main.repors!comp = True Then
            paperWidth = Int(main.repors!totalcolumn * 1.75)
        Else
            paperWidth = main.repors!totalcolumn
        End If
        Open "" + App.Path + "\vipin.txt" For Output As #1
        Line = 0
        Pno = 1
header:
            If FooterYes = True Then
                Do While Line < 72
                    Print #1, " "
                    Line = Line + 1
                Loop
                FooterYes = False
                Line = 0
            End If
            If kkk.State = 1 Then kkk.Close
            CNSetup
            kkk.Open "Select * from setup where " & stringyear & "", CON, adOpenDynamic, adLockReadOnly, adCmdText
            If Not kkk.BOF Then
                Print #1, ""
                Print #1, ""
                Print #1, Chr(27) + Chr(15) + Chr(14)
                Print #1, Tab(120); "Page No:  " & Pno
                Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!cname))
                Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
            End If
            If trs.State = 1 Then trs.Close
            
            'trs.Open "treport", CON, adOpenDynamic, adLockReadOnly, adcmdtext
             trs.Open "select * from treport where  " & stringyear & " and userid=" & main.UId, CON, adOpenDynamic, adLockReadOnly, adCmdText
            
            Print #1, Tab(((paperWidth - (Len(Trim("Sub Ledger Trial Balance")))) / 2) + LEFTM); "Sub Ledger Trial Balance"
            Print #1, Tab(LEFTM + ((paperWidth - Len(Trim(COMBOGENLEDGER.Text))) / 2)); Trim(COMBOGENLEDGER.Text)
            xstr = date1.Text & " To " & date2.Text
            Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Trim("Period : " + date1.Text & " To " & date2.Text);
            Print #1, Tab(LEFTM); repli("-", paperWidth)
            Print #1, Tab(0); Chr(27) + Chr(71); Tab(8); "Sub. Ledger Description"; Tab(46); "Opening Balance"; Tab(67); "Amount (Dr.)"; Tab(89); "Amount (Cr.)"; Tab(110); "Closing Balance"; Chr(27) + Chr(72)
            Print #1, Tab(LEFTM); repli("-", paperWidth)
            Print #1, ""
            Line = 13
            trs.Close
            If called1 Then
                GoTo printagain1
                called1 = False
            End If
            If RS.State = 1 Then RS.Close
            
            RS.Open "select gledger,subledger,sum(openingbalance)as openingbalance1  ,sum(Damount)as Damount1, sum(Camount) as Camount1 , (sum(openingbalance)+sum(Damount)- sum(Camount))  as ClosingBalance    from TemprptTrialBalance where  " & stringyear & " and userid=" & main.UId & " and openingbalance<>0   or damount<>0 or cAmount<>0 group  by gledger,subledger ", CON, adOpenStatic, adLockReadOnly, adCmdText
            
            Dim CB As Double
            While Not RS.EOF
                CB = 0
                CB = CB + Val(RS(2) & "") + RS(3) - Abs(RS(4))
                Print #1, Tab(1); RS!SUBLEDGER; Tab(46); IIf(RS!openingbalance1 <> 0, rsets(Trim(Format(RS!openingbalance1, "0.00")), 12), ""); Tab(65); IIf(RS(3) <> 0, rsets(Trim(Format(RS(3), "0.00")), 12), ""); Tab(85); IIf(RS(4) <> 0, rsets(Trim(Format(str(RS(4)), "0.00")), 12), ""); Tab(110); IIf(CB <> 0, IIf(CB > 0, rsets(Trim(Format(str(CB), "0.00")), 12) & "   Dr. ", rsets(Trim(Format(str(CB), "0.00")), 12) & "      Cr."), "")
                
                Line = Line + 1
                GopenBal = GopenBal + Val(RS(2) & "")
                GopenDr = GopenDr + RS(3)
                GopenCr = GopenCr + RS(4)
                GopenCl = GopenCl + Val(RS(2) & "") + RS(3) - Abs(RS(4))
                
                
                If Line > MaxLine - 9 Then
                        called1 = True
                        Pno = Pno + 1
                        FooterYes = True
                        GoTo header
printagain1:
                        
                        called1 = False
                End If
                If Not RS.EOF Then RS.MoveNext
          Wend
printfooter:
            Print #1, Tab(LEFTM); repli("-", paperWidth)
            Print #1, Tab(LEFTM); "* * * NET BALANCE * * * "; Tab(46); IIf(GopenBal <> 0, rsets(Format(Trim(GopenBal), "0.00"), 12), ""); Tab(65); IIf(GopenDr <> 0, rsets(Format(Trim(GopenDr), "0.00"), 12), ""); Tab(85); IIf(GopenCr <> 0, rsets(Format(Trim(GopenCr), "0.00"), 12), ""); Tab(110); IIf(GopenCl <> 0, IIf(GopenCl > 0, rsets(Format(Trim(GopenCl), "0.00"), 12) & "   Dr.", rsets(Format(Trim(GopenCl), "0.00"), 12) & "     Cr."), "")
            Print #1, Tab(LEFTM); repli("-", paperWidth)
            Line = Line + 3
            Do While Line < 72
                    Print #1, " "
                    Line = Line + 1
            Loop
            Close #1
End Sub

Private Sub print1_Click()
    c1.PrinterDefault = True
    c1.ShowPrinter
    printnow
 End Sub
