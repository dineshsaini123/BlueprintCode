VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmPartyStatment 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2535
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   5970
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPartyStatment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Begin Crystal.CrystalReport cr 
      Left            =   375
      Top             =   1425
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox cboItem 
      Height          =   315
      Left            =   1575
      TabIndex        =   14
      Top             =   450
      Width           =   3750
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Crystal Report"
      Height          =   495
      Left            =   6975
      TabIndex        =   13
      Top             =   975
      Width           =   1470
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Dos report"
      Height          =   495
      Left            =   7425
      TabIndex        =   12
      Top             =   975
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5790
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton export 
      Height          =   345
      Left            =   1470
      Picture         =   "frmPartyStatment.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4740
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmPartyStatment.frx":045D
      Left            =   2370
      List            =   "frmPartyStatment.frx":0470
      TabIndex        =   8
      Text            =   "100 %"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton print1 
      Height          =   345
      Left            =   1920
      Picture         =   "frmPartyStatment.frx":0496
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4740
      Visible         =   0   'False
      Width           =   345
   End
   Begin RichTextLib.RichTextBox r1 
      Height          =   2145
      Left            =   840
      TabIndex        =   6
      Top             =   5790
      Visible         =   0   'False
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   3784
      _Version        =   393217
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      MaxLength       =   99999999
      RightMargin     =   20000
      TextRTF         =   $"frmPartyStatment.frx":0608
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
      Left            =   3225
      TabIndex        =   5
      Top             =   1545
      Width           =   1545
   End
   Begin VB.CommandButton frmPartyStatment 
      Caption         =   "&Show"
      Height          =   405
      Left            =   1620
      TabIndex        =   4
      Top             =   1545
      Width           =   1545
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   6750
      TabIndex        =   0
      Top             =   375
      Width           =   360
   End
   Begin MSMask.MaskEdBox alpha 
      Height          =   345
      Left            =   7650
      TabIndex        =   1
      Top             =   1050
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      _Version        =   393216
      MaxLength       =   1
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin MSComDlg.CommonDialog c1 
      Left            =   675
      Top             =   4800
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
      Left            =   7200
      TabIndex        =   2
      Top             =   2400
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
      Left            =   7050
      TabIndex        =   3
      Top             =   1500
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
      Caption         =   "Item Name"
      Height          =   255
      Left            =   525
      TabIndex        =   15
      Top             =   450
      Width           =   990
   End
   Begin VB.Label Label3 
      Caption         =   " - To - "
      Height          =   315
      Left            =   7350
      TabIndex        =   11
      Top             =   2025
      Width           =   585
   End
End
Attribute VB_Name = "frmPartyStatment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As Connection
Dim Q1 As String
Dim rs As Recordset

Private Sub alpha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub COMBOGENLEDGER_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub COMBOGENLEDGER_LostFocus()
    If Trim(COMBOGENLEDGER.Text) <> "" Then
       If rs.State = 1 Then rs.Close
    
        
                rs.Open "select * from gledger where slf=true", CON, adOpenDynamic, adLockReadOnly, adCmdText
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
MainMenu.Toolbar1.Visible = True
End Sub


Sub SUBLEDGERSBALANCE()
       
If Trim(alpha.Text) <> "" Then


CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 as YEAROPENING,  Sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER))" _
        & " where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER; "

                
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER)) " _
        & " where  gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'  and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "')AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; "
        
        
        CON.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER))" _
        & " where  gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'  and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; "
   
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER)) " _
        & " where  gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; "
  
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE>=cdate('" + Trim(date1.Text) + "') and VOUCHERDATE<=cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;"
        
   
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE DEBITORCREDIT = 'C' and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE>=cdate('" + Trim(date1.Text) + "') and VOUCHERDATE<=cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; "
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE (((CNF1A.DC)='D')) and gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER  " _
        & " HAVING  Sum(CNF1A.NA) <> 0; "

        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE (((CNF1A.DC)='C')) and gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  cnd >= cdate('" + Trim(date1.Text) + "') and cnd <= cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; "
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE ((( DNFA.DC) = 'D' )) and  gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER" _
        & " HAVING  Sum(DNFA.NA) <> 0; "
        

        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE (((DNFA.DC)='C')) and   gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; "

        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE (((CNF1B.DC)='D')) and gledger  = '" + Trim(COMBOGENLEDGER.Text) + "' and  cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; "
        

        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE (((CNF1B.DC)= 'C')) and gledger  = '" + Trim(COMBOGENLEDGER.Text) + "' and  cnd >= cdate('" + Trim(date1.Text) + "') and cnd <= cdate('" + Trim(date2.Text) + "')  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;"

        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE (((DNFB.DC)='D')) and   gledger  = '" + Trim(COMBOGENLEDGER.Text) + "' and  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; "
        

        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE (((DNFB.DC)= 'C')) and  gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; "


Else
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 as YEAROPENING,  Sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER))" _
        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER; "
        

        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER)) " _
        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "'  and SLEDGER.SUBLEDGER <>'CASH PARTY' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; "
        

        CON.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER))" _
        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "'and SLEDGER.SUBLEDGER <>'CASH PARTY'  and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; " _
        
   
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER)) " _
        & " where  genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; " _
        
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE>=cdate('" + Trim(date1.Text) + "') and VOUCHERDATE<=cdate('" + Trim(date2.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;"
   
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE>=cdate('" + Trim(date1.Text) + "') and VOUCHERDATE<=cdate('" + Trim(date2.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; "
   
        CON.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and  cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER  " _
        & " HAVING  Sum(CNF1A.NA) <> 0; "
        
 
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and  cnd >= cdate('" + Trim(date1.Text) + "') and cnd <= cdate('" + Trim(date2.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; "
        
        
     
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId  " _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER" _
        & " HAVING  Sum(DNFA.NA) <> 0; "
        
 
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.Text) + "' and  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; "
        
  
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and  cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; "
        
    
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT  , " & UId & " as UserId" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.Text) + "' and  cnd >= cdate('" + Trim(date1.Text) + "') and cnd <= cdate('" + Trim(date2.Text) + "')  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;"
        
 
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.Text) + "' and  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; "
        
   
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.Text) + "' and  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "') " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; "
        
        
   End If
   DoEvents
   CON.Execute "insert into TemprptTrialBalance ( Subledger, Damount,CAmount,userid) SELECT  SUBLEDGER,  SUM (OPAMOUNTDEBIT) as Damount,  SUM(OPAMOUNTCREDIT)as Camount,userid  from subledgertrail GROUP BY SUBLEDGER,userid;"

End Sub
Sub OPENINGSUBLEDGERS()
CON.Execute "Delete  from subledgertrail"
If Trim(alpha.Text) <> "" Then
    
     CON.Execute "Insert into subledgertrail  SELECT SUBLEDGER , YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId FROM SLEDGER where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' AND subledger like '" + Trim(alpha.Text) + "%'", p, adCmdText
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER))" _
        & " where  gledger='" + Trim(COMBOGENLEDGER.Text) + "'and INVOICEDATE < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT,  " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER)) " _
        & " where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'  and INVOICEDATE< cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
        
        CON.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT ," & UId & " as UserId  " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) AND (SLEDGER.gledger = CASHA.GENLEDGER))" _
        & " where  gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'   and INVOICEDATE < cdate('" + Trim(date1.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) AND (SLEDGER.gledger =CREDITA.GENLEDGER)) " _
        & " where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
        
        
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT ,  " & UId & " as UserId" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE DEBITORCREDIT = 'D' and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE < cdate('" + Trim(date1.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
        
        
        CON.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE DEBITORCREDIT = 'C' and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE < cdate('" + Trim(date1.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
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
        
        CON.Execute "Delete * from TemprptTrialBalance"
        CON.Execute "insert into TemprptTrialBalance ( Subledger,openingbalance,userid ) SELECT  SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId from subledgertrail GROUP BY SUBLEDGER;"
    
    
Else
  
    CON.Execute "Insert into subledgertrail SELECT SUBLEDGER , YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT, " & UId & " as UserId  FROM SLEDGER where  gledger='" + Trim(COMBOGENLEDGER.Text) + "'", p, adCmdText

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
        CON.Execute "Delete * from TemprptTrialBalance", p, adCmdText
        CON.Execute "insert into TemprptTrialBalance (Subledger,openingbalance,userid) SELECT  SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR , " & UId & " as UserId   from subledgertrail GROUP BY SUBLEDGER;", p, adCmdText
  
  End If
  
  Exit Sub
End Sub


Private Sub Commandreturn_Click()
MainMenu.Toolbar1.Visible = True
Unload Me
End Sub
Private Sub Commandshow_Click()
If Trim(COMBOGENLEDGER.Text) <> "" Then
    If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
        MsgBox "invalid date", , title1
        Exit Sub
    End If
    
    CON.Execute "DELETE * FROM TemprptTrialBalance"
    CON.Execute "DELETE * FROM subledgertrail"
    OPENINGSUBLEDGERS
    CON.Execute "Delete from subledgertrail"
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

Me.r1.Top = 10
Me.r1.Left = 10
'''''Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu"))
'''''    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
'''''        Unload VB.Screen.ActiveForm
'''''    End If
'''''Loop


Me.Top = 0
Me.Left = 0
'CON.Execute "DELETE * FROM TemprptTrialBalance"
'CON.Execute "DELETE * FROM subledgertrail"
'con.Execute "DELETE * FROM TemprptTrialBalance"
'con.Execute "DELETE * FROM subledgertrail"

'Set CON = New ADODB.Connection
'CON.CursorLocation = adUseClient
Set rs = New ADODB.Recordset

 '   With CON
  '      .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + VB.App.Path + "\" + Trim(main.directory) + "\data.mdb"
  '      .Open
   ' End With
    'rs.Open "select * from gledger where slf=true", CON, adOpenDynamic, adLockReadOnly, adCmdText
       rs.Open "select * from gledger where slf=true", CON, adOpenDynamic, adLockReadOnly, adCmdText
    If Not rs.BOF Then
        Do While Not rs.EOF
            COMBOGENLEDGER.AddItem Trim(rs!gledger)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    

    rs.Close
    CNSetup
    rs.Open "setup1", CON, adOpenDynamic, adLockReadOnly, adCmdTable
    date1.Text = rs!yarfrom
    date2.Text = rs!yarto
    rs.Close
    
    Me.COMBOGENLEDGER.Text = "SUNDRY DEBTORS"
    
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "select distinct(district) from INVOICEA", CON
    While rs.EOF = False
      cboItem.AddItem rs(0) & ""
      rs.MoveNext
    Wend
    
End Sub

Private Sub print_Click()
  
Rsinvoicea.Open "select GenLedger,  SubLedger , sum(amount) as INVAmount from invoicea where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') order by invoicedate", CON, adOpenDynamic, adLockReadOnly, adCmdText
   
       RsCREDITa.Open "select * from CREDITa where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "')", CON, adOpenDynamic, adLockReadOnly, adCmdText
       RsCnf1a.Open "select * from Cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')", CON, adOpenDynamic, adLockReadOnly, adCmdText
       Rsdnfa.Open "select * from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "' and psld='" + Trim(Combosubledger.Text) + "' and dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "')", CON, adOpenDynamic, adLockReadOnly, adCmdText
       RsCnf1B.Open "select * from Cnf1B where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')", CON, adOpenDynamic, adLockReadOnly, adCmdText
       RsdnfB.Open "select * from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "' and dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "')", CON, adOpenDynamic, adLockReadOnly, adCmdText
       RScasha.Open "select * from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and subledger='" + Trim(Combosubledger.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') order by invoicedate", CON, adOpenDynamic, adLockReadOnly, adCmdText
  
  
End Sub


Sub Genrate()
    Command1.Top = r1.Top + r1.Height + 30
    Combo1.Top = r1.Top + r1.Height + 30
    Set rs = New ADODB.Recordset
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
        Open "" + VB.App.Path + "\vipin.txt" For Output As #1
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
            kkk.Open "select * from setup1", CON, adOpenDynamic, adLockReadOnly, adCmdText
            If Not kkk.BOF Then
                Print #1, ""
                Print #1, ""
                Print #1, Chr(27) + Chr(15) + Chr(14)
                Print #1, Tab(120); "Page No:  " & Pno
                Print #1, Tab(((paperWidth - (Len(Trim(kkk!CNAME)) * 2)) / 2) - 15); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!CNAME))
                Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
            End If
            If trs.State = 1 Then trs.Close
            
            'trs.Open "treport", CON1, adOpenDynamic, adLockReadOnly, adCmdTable
             trs.Open "treport", CON, adOpenDynamic, adLockReadOnly, adCmdTable
            
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
            If rs.State = 1 Then rs.Close
            
            rs.Open "select gledger,subledger,sum(openingbalance)as openingbalance1  ,sum(Damount)as Damount1, sum(Camount) as Camount1 , (sum(openingbalance)+sum(Damount)- sum(Camount))  as ClosingBalance    from TemprptTrialBalance where openingbalance<>0   or damount<>0 or cAmount<>0 and userid = " & UId & " group  by gledger,subledger ", CON, adOpenStatic, adLockReadOnly, adCmdText
            
            Dim CB As Double
            While Not rs.EOF
                CB = 0
                CB = CB + rs(2) + rs(3) - Abs(rs(4))
                Print #1, Tab(1); rs!subledger; Tab(46); IIf(rs!openingbalance1 <> 0, rsets(Trim(Format(rs!openingbalance1, "0.00")), 12), ""); Tab(65); IIf(rs(3) <> 0, rsets(Trim(Format(rs(3), "0.00")), 12), ""); Tab(85); IIf(rs(4) <> 0, rsets(Trim(Format(Str(rs(4)), "0.00")), 12), ""); Tab(110); IIf(CB <> 0, IIf(CB > 0, rsets(Trim(Format(Str(CB), "0.00")), 12) & "   Dr. ", rsets(Trim(Format(Str(CB), "0.00")), 12) & "      Cr."), "")
                Line = Line + 1
                GopenBal = GopenBal + rs(2)
                GopenDr = GopenDr + rs(3)
                GopenCr = GopenCr + rs(4)
                GopenCl = GopenCl + rs(2) + rs(3) - Abs(rs(4))
                If Line > MaxLine - 9 Then
                        called1 = True
                        Pno = Pno + 1
                        FooterYes = True
                        GoTo header
printagain1:
                        
                        called1 = False
                End If
                If Not rs.EOF Then rs.MoveNext
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

Private Sub frmPartyStatment_Click()
    Dim rs1 As New ADODB.Recordset
    Dim cramt As Double
    Dim opamt As Double
    Dim amt As Double
    cramt = 0
    opamt = 0
    amt = 0
    CON.Execute "DELETE * FROM TemprptTrialBalance"
    CON.Execute "DELETE * FROM subledgertrail"
    CON.Execute "update INVOICEA set RecAmt=0"
    
    OPENINGSUBLEDGERS
    SUBLEDGERSBALANCE
    
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select SUBLEDGER from SLEDGER where DISTCODE='" & Me.cboItem.Text & "'", CON
    'rs1.Open "select SUBLEDGER from SLEDGER where subledger='AB002 M/S GOPAL SURGICAL'", CON
    
    While rs1.EOF = False
    If rs.State = 1 Then rs.Close
    rs.Open "select sum(OpeningBalance),sum(CAmount) from TemprptTrialBalance where Subledger='" & rs1(0) & "'", CON
    If rs.RecordCount > 0 Then
       opamt = rs(0)
       cramt = rs(1)
       CON.Execute "update invoicea set BAA=" & rs(0) & " where Subledger='" & rs1(0) & "'"
    End If
    If opamt <= cramt Then
        amt = cramt - opamt
        CON.Execute "update invoicea set BAA=" & 0 & " where Subledger='" & rs1(0) & "'"
        CON.Execute "update SLEDGER set RecAmt=" & opamt & " where Subledger='" & rs1(0) & "'"
    Else
        amt = opamt - cramt
        CON.Execute "update SLEDGER set RecAmt=" & amt & " where Subledger='" & rs1(0) & "'"
        CON.Execute "update invoicea set BAA=" & amt & " where Subledger='" & rs1(0) & "'"
    amt = 0
    End If

    If amt > 0 Then
    If rs.State = 1 Then rs.Close
    rs.Open "select  NETAMOUNT,invoiceno from INVOICEA where  Subledger='" & rs1(0) & "'", CON
    While rs.EOF = False
    If rs!netamount <= amt Then
    amt = amt - rs!netamount
    CON.Execute "update INVOICEA set RecAmt=" & rs!netamount & " where invoiceno=" & rs(1) & ""
    Else
    CON.Execute "update INVOICEA set RecAmt=" & amt & " where invoiceno=" & rs(1) & ""
    'amt = rs!netamount - amt
    amt = 0

    End If

    If amt <= 0 Then
    GoTo aa11
    End If

    rs.MoveNext
    Wend
    End If

aa11:

    rs1.MoveNext
    Wend

    
If MsgBox("Want to Print !!", vbQuestion + vbYesNo) = vbYes Then
    
    CR.Reset
    CR.ReportFileName = App.Path & "\reports\PartyWiseStm.rpt"
    CR.DataFiles(0) = App.Path & "\" + Trim(main.directory) + "\data.mdb"
    CR.DataFiles(1) = App.Path & "\" + Trim(main.directory) + "\data.mdb"
    CR.DataFiles(2) = App.Path & "\" + Trim(main.directory) + "\data.mdb"
    CR.DataFiles(3) = App.Path & "\" + Trim(main.directory) + "\data.mdb"
    CR.DataFiles(4) = App.Path & "\" + Trim(main.directory) + "\data.mdb"
    If cboItem.Text <> "" Then
        CR.SelectionFormula = "(({invoicea.netAmount}-{invoicea.recamt})>0) and {invoicea.district} = '" & cboItem.Text & "'"
    End If
    CR.WindowShowRefreshBtn = True
    CR.WindowShowPrintBtn = True
    CR.WindowShowPrintSetupBtn = True
    CR.WindowState = crptMaximized
    CR.Action = 1
    
 End If

End Sub

Private Sub print1_Click()
    c1.PrinterDefault = True
    c1.ShowPrinter
    printnow
 End Sub
