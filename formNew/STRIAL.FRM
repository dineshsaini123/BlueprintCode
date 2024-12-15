VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form subtrial 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4476
   ClientLeft      =   252
   ClientTop       =   1416
   ClientWidth     =   9708
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4476
   ScaleWidth      =   9708
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1_drcr 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Print Balance Only (Dr/Cr) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3060
      TabIndex        =   18
      Top             =   2475
      Width           =   2670
   End
   Begin VB.ComboBox cboDRCr 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "STRIAL.frx":0000
      Left            =   5760
      List            =   "STRIAL.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2475
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.ComboBox AgCombo 
      Height          =   315
      Left            =   3045
      Sorted          =   -1  'True
      TabIndex        =   15
      Top             =   285
      Width           =   3885
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
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4740
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2370
      TabIndex        =   10
      Text            =   "100 %"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton print1 
      Height          =   345
      Left            =   1920
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
      _ExtentX        =   15854
      _ExtentY        =   7218
      _Version        =   393217
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      MaxLength       =   99999999
      RightMargin     =   20000
      TextRTF         =   $"STRIAL.frx":001F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Commandreturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   630
      Left            =   4800
      Picture         =   "STRIAL.frx":009F
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3075
      Width           =   1545
   End
   Begin VB.CommandButton Commandshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Show"
      Height          =   630
      Left            =   3060
      Picture         =   "STRIAL.frx":0C83
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3075
      Width           =   1545
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   3045
      TabIndex        =   0
      Top             =   960
      Width           =   3885
   End
   Begin MSMask.MaskEdBox alpha 
      Height          =   345
      Left            =   3045
      TabIndex        =   1
      Top             =   1380
      Width           =   405
      _ExtentX        =   720
      _ExtentY        =   614
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
      Left            =   3090
      TabIndex        =   2
      Top             =   1935
      Width           =   1155
      _ExtentX        =   2032
      _ExtentY        =   572
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
      Left            =   5580
      TabIndex        =   3
      Top             =   1935
      Width           =   1155
      _ExtentX        =   2032
      _ExtentY        =   572
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Agent Name"
      Height          =   195
      Left            =   960
      TabIndex        =   16
      Top             =   315
      Width           =   885
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From The Date"
      Height          =   315
      Left            =   960
      TabIndex        =   14
      Top             =   1905
      Width           =   1995
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " - To - "
      Height          =   315
      Left            =   4680
      TabIndex        =   13
      Top             =   1965
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alphabat"
      Height          =   195
      Left            =   960
      TabIndex        =   7
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gen. Ledger Desc."
      Height          =   195
      Left            =   960
      TabIndex        =   6
      Top             =   990
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
If KeyAscii = 13 Then sendkeys "{tab}"
End Sub

Private Sub Check1_drcr_Click()
If Check1_drcr.value = 0 Then
   cboDRCr.Visible = False
Else
   cboDRCr.Visible = True
   cboDRCr.ListIndex = 2
End If
End Sub

Private Sub COMBOGENLEDGER_KeyPress(KeyAscii As Integer)
If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
   sendkeys "{tab}"
   Exit Sub
End If
If KeyAscii = 13 Then
      sendkeys "{Down}"
      sendkeys "{tab}"
End If

End Sub

Private Sub COMBOGENLEDGER_LostFocus()
    If Trim(COMBOGENLEDGER.text) <> "" Then
       If RS.State = 1 Then RS.close
    
        RS.Open "select * from gledger where " & stringyear & " and slf=1", con, adOpenDynamic, adLockReadOnly, adCmdText
        If Not RS.BOF Then
            RS.Find "gledger='" + Trim(COMBOGENLEDGER.text) + "'"
            If RS.EOF Then
                COMBOGENLEDGER.SetFocus
            End If
        Else
            COMBOGENLEDGER.SetFocus
        End If
        RS.close
    End If
End Sub

Private Sub Command1_Click()
Unload Me
''MainMenu.Toolbar1.Visible = True
End Sub
Sub SUBLEDGERSBALANCE()
       
If Trim(Alpha.text) <> "" Then

'        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 as YEAROPENING,  Sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER))" _
'        & " where  " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
'        & " GROUP BY SLEDGER.SUBLEDGER; "
'
'
'        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER)) " _
'        & " where " & stringyear & " and  gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'  and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "')AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; "
'
'
'
'
'        con.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER))" _
'        & " where  " & stringyear & " and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'  and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(CASHA.baa) <> 0; "
'
'
'
'        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER)) " _
'        & " where  " & stringyear & " and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=cdate('" + Trim(date1.Text) + "') and INVOICEDATE<=cdate('" + Trim(date2.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; "
'
'
'
'        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
'        & " WHERE " & stringyear & " and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE>=cdate('" + Trim(date1.Text) + "') and VOUCHERDATE<=cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;"
'
'
'        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
'        & " WHERE " & stringyear & " and DEBITORCREDIT = 'C' and gledger ='" + Trim(COMBOGENLEDGER.Text) + "' and VOUCHERDATE>=cdate('" + Trim(date1.Text) + "') and VOUCHERDATE<=cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; "
'
'
'
'        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
'        & " WHERE " & stringyear & " and (((CNF1A.DC)='D')) and gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
'        & " GROUP BY SLEDGER.SUBLEDGER  " _
'        & " HAVING  Sum(CNF1A.NA) <> 0; "
'
'
'        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
'        & " WHERE (((CNF1A.DC)='C')) and gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  cnd >= cdate('" + Trim(date1.Text) + "') and cnd <= cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(CNF1A.NA) <> 0; "
'
'
'
'        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
'        & " WHERE " & stringyear & " and ((( DNFA.DC) = 'D' )) and  gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
'        & " GROUP BY SLEDGER.SUBLEDGER" _
'        & " HAVING  Sum(DNFA.NA) <> 0; "
'
'
'        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
'        & " WHERE " & stringyear & " and (((DNFA.DC)='C')) and   gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "') AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%' " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(DNFA.NA) <> 0; "
'
'
'        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
'        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
'        & " WHERE " & stringyear & " and (((CNF1B.DC)='D')) and gledger  = '" + Trim(COMBOGENLEDGER.Text) + "' and  cnd>=cdate('" + Trim(date1.Text) + "') and cnd<=cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(CNF1B.A) <> 0; "
'
'
'
'        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT  , " & UId & " as UserId " _
'        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
'        & " WHERE " & stringyear & " and (((CNF1B.DC)= 'C')) and gledger  = '" + Trim(COMBOGENLEDGER.Text) + "' and  cnd >= cdate('" + Trim(date1.Text) + "') and cnd <= cdate('" + Trim(date2.Text) + "')  " _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(CNF1B.A) <> 0;"
'
'
'        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
'        & " WHERE " & stringyear & " and (((DNFB.DC)='D')) and   gledger  = '" + Trim(COMBOGENLEDGER.Text) + "' and  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(DNFB.A) <> 0; "
'
'
'        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId " _
'        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
'        & " WHERE " & stringyear & " and (((DNFB.DC)= 'C')) and  gledger = '" + Trim(COMBOGENLEDGER.Text) + "' and  dnd>=cdate('" + Trim(date1.Text) + "') and dnd<=cdate('" + Trim(date2.Text) + "')  AND SLEDGER.subledger like '" + Trim(alpha.Text) + "%'" _
'        & " GROUP BY SLEDGER.SUBLEDGER " _
'        & " HAVING  Sum(DNFB.A) <> 0; "


        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 as YEAROPENING,  Sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) )" _
        & " where  INVOICEA.setupid=" & setupid & " and INVOICEA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER; "
        

        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER)) " _
        & " where  cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "'  and SLEDGER.SUBLEDGER <>'CASH PARTY' and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; "
        

        
        
        con.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER))" _
        & " where  cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "'and SLEDGER.SUBLEDGER <>'CASH PARTY'  and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; " _
        
   
        
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT  , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) ) " _
        & " where  CREDITA.setupid=" & setupid & " and CREDITA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; " _
        
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE VOUCHERS.setupid=" & setupid & " and VOUCHERS.fyear='" & session & "' and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VOUCHERDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,VOUCHERDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;"
   
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE VOUCHERS.setupid=" & setupid & " and VOUCHERS.fyear='" & session & "' and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VOUCHERDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,VOUCHERDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; "
   
        con.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE CNF1A.setupid=" & setupid & " and CNF1A.fyear='" & session & "' and (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER  " _
        & " HAVING  Sum(CNF1A.NA) <> 0; "
        
 
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE CNF1A.setupid=" & setupid & " and CNF1A.fyear='" & session & "' and (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,cnd,103) <= convert(smalldatetime,'" + Trim(date2.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; "
        
        
     
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE dnfa.setupid=" & setupid & " and dnfa.fyear='" & session & "' and ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER" _
        & " HAVING  Sum(DNFA.NA) <> 0; "
        
 
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE dNFA.setupid=" & setupid & " and dNFA.fyear='" & session & "' and (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; "
        
  
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1b.setupid=" & setupid & " and CNF1b.fyear='" & session & "' and (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; "
        
    
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT  , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1b.setupid=" & setupid & " and CNF1b.fyear='" & session & "' and (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,cnd,103) <= convert(smalldatetime,'" + Trim(date2.text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;"
        
 
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE DNFB.setupid=" & setupid & " and DNFB.fyear='" & session & "' and (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; "
        
   
        
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT  , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE DNFB.setupid=" & setupid & " and DNFB.fyear='" & session & "' and (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; "

Else
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 as YEAROPENING,  Sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) )" _
        & " where  INVOICEA.setupid=" & setupid & " and INVOICEA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER; "
        

        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER)) " _
        & " where  cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "'  and SLEDGER.SUBLEDGER <>'CASH PARTY' and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; "
        

        
        
        con.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER))" _
        & " where  cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "'and SLEDGER.SUBLEDGER <>'CASH PARTY'  and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; " _
        
   
        
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT  , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER) ) " _
        & " where  CREDITA.setupid=" & setupid & " and CREDITA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; " _
        
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE VOUCHERS.setupid=" & setupid & " and VOUCHERS.fyear='" & session & "' and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VOUCHERDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,VOUCHERDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;"
   
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE VOUCHERS.setupid=" & setupid & " and VOUCHERS.fyear='" & session & "' and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VOUCHERDATE,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,VOUCHERDATE,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; "
   
        con.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE CNF1A.setupid=" & setupid & " and CNF1A.fyear='" & session & "' and (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER  " _
        & " HAVING  Sum(CNF1A.NA) <> 0; "
        
 
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE CNF1A.setupid=" & setupid & " and CNF1A.fyear='" & session & "' and (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,cnd,103) <= convert(smalldatetime,'" + Trim(date2.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; "
        
        
     
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE dnfa.setupid=" & setupid & " and dnfa.fyear='" & session & "' and ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER" _
        & " HAVING  Sum(DNFA.NA) <> 0; "
        
 
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE dNFA.setupid=" & setupid & " and dNFA.fyear='" & session & "' and (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; "
        
  
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1b.setupid=" & setupid & " and CNF1b.fyear='" & session & "' and (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; "
        
    
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT  , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1b.setupid=" & setupid & " and CNF1b.fyear='" & session & "' and (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,cnd,103) <= convert(smalldatetime,'" + Trim(date2.text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;"
        
 
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE DNFB.setupid=" & setupid & " and DNFB.fyear='" & session & "' and (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; "
        
   
        
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT  , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE DNFB.setupid=" & setupid & " and DNFB.fyear='" & session & "' and (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; "
        
        
   End If
   DoEvents
   con.Execute "insert into TemprptTrialBalance ( Subledger, Damount,CAmount,userid,setupid,fyear) SELECT  SUBLEDGER,  SUM (OPAMOUNTDEBIT) as Damount,  SUM(OPAMOUNTCREDIT)as Camount,userid," & setupid & ",'" & session & "'  from subledgertrail  where " & stringyear & " GROUP BY SUBLEDGER,userid;"




End Sub


Sub OPENINGSUBLEDGERS()
con.Execute "Delete  from subledgertrail where " & stringyear
If Trim(Alpha.text) <> "" Then
        con.Execute "Insert into subledgertrail  SELECT SUBLEDGER , YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'FROM SLEDGER where  " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "' AND subledger like '" + Trim(Alpha.text) + "%'", p, adCmdText
        
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) AND (SLEDGER.gledger = INVOICEA.GENLEDGER))" _
        & " where INVOICEA.setupid=" & setupid & " and INVOICEA.fyear='" & session & "' and gledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER ", p, adCmdText
        ''AND (SLEDGER.gledger = INVOICEA.GENLEDGER)
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT,  " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) ) " _
        & " where  cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and gledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'  and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
        
        'AND (SLEDGER.gledger = CASHA.GENLEDGER)
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT ," & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) )" _
        & " where cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and gledger ='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'   and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
        
        'AND (SLEDGER.gledger = CASHA.GENLEDGER)
        
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER)) " _
        & " where  CREDITA.setupid=" & setupid & " and CREDITA.fyear='" & session & "' and gledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
        
        'AND (SLEDGER.gledger =CREDITA.GENLEDGER)
        
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT ,  " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE VOUCHERS.setupid=" & setupid & " and VOUCHERS.fyear='" & session & "' and DEBITORCREDIT = 'D' and gledger ='" + Trim(COMBOGENLEDGER.text) + "' and  convert(smalldatetime,VOUCHERDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE VOUCHERS.setupid=" & setupid & " and VOUCHERS.fyear='" & session & "' and DEBITORCREDIT = 'C' and gledger ='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VOUCHERDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
         ''''ok
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and CONVERT(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE CNF1A.setupid=" & setupid & " and CNF1A.fyear='" & session & "' and (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE DNFA.setupid=" & setupid & " and DNFA.fyear='" & session & "' and ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE DNFA.setupid=" & setupid & " and DNFA.fyear='" & session & "' and (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail     SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1B.setupid=" & setupid & " and CNF1B.fyear='" & session & "' and (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1B.setupid=" & setupid & " and CNF1B.fyear='" & session & "' and (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.SUBLEDGER like '" + Trim(Alpha.text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail     SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE DNFB.setupid=" & setupid & " and DNFB.fyear='" & session & "' and (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.text) + "'And convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)  AND SLEDGER.SUBLEDGER like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE DNFB.setupid=" & setupid & " and DNFB.fyear='" & session & "' and (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.SUBLEDGER like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
        
        con.Execute "Delete from TemprptTrialBalance where " & stringyear
        con.Execute "insert into TemprptTrialBalance ( Subledger,openingbalance,userid ) SELECT  SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId from subledgertrail where " & stringyear & "  GROUP BY SUBLEDGER;"
    
          
          
          




Else
        con.Execute "Insert into subledgertrail SELECT SUBLEDGER , YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT, " & UId & " as UserId," & setupid & ",'" & session & "'  FROM SLEDGER where  " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "'", p, adCmdText

        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER))  " _
        & " where  INVOICEA.setupid=" & setupid & " and INVOICEA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "'and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
        DoEvents
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER)) " _
        & " where  cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and gledger='" + Trim(COMBOGENLEDGER.text) + "' and  SLEDGER.SUBLEDGER <>'CASH PARTY'  and convert(smalldatetime,INVOICEDATE,103)< convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
        DoEvents
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER))" _
        & " where  cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER)) " _
        & " where  CREDITA.setupid=" & setupid & " and CREDITA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
        
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE VOUCHERS.setupid=" & setupid & " and VOUCHERS.fyear='" & session & "' and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VOUCHERDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE VOUCHERS.setupid=" & setupid & " and VOUCHERS.fyear='" & session & "' and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VOUCHERDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
         ''''ok
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE CNF1A.setupid=" & setupid & " and CNF1A.fyear='" & session & "' and (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE CNF1A.setupid=" & setupid & " and CNF1A.fyear='" & session & "' and (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE dNFA.setupid=" & setupid & " and dNFA.fyear='" & session & "' and ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE dNFA.setupid=" & setupid & " and dNFA.fyear='" & session & "' and (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1b.setupid=" & setupid & " and CNF1b.fyear='" & session & "' and (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1b.setupid=" & setupid & " and CNF1b.fyear='" & session & "' and (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE dnfb.setupid=" & setupid & " and dnfb.fyear='" & session & "' and (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.text) + "'And convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE dnfb.setupid=" & setupid & " and dnfb.fyear='" & session & "' and (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
        con.Execute "Delete  from TemprptTrialBalance where " & stringyear, p, adCmdText
        con.Execute "insert into TemprptTrialBalance ( Subledger,openingbalance,userid,setupid,fyear) SELECT  SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR , " & UId & " as UserId," & setupid & ",'" & session & "'  from subledgertrail where " & stringyear & "  GROUP BY SUBLEDGER;", p, adCmdText
  End If
  
  
  
  Exit Sub
End Sub
Private Sub CommandReturn_Click()
''MainMenu.Toolbar1.Visible = True
Unload Me
End Sub
Private Sub Commandshow_Click()
Commandshow.Enabled = False
If Trim(COMBOGENLEDGER.text) <> "" Then
    If DateDiff("d", Trim(date1.text), Trim(date2.text)) <= 0 Then
        MsgBox "invalid date"
        Exit Sub
    End If
    
    con.Execute "DELETE  FROM TemprptTrialBalance where " & stringyear
    con.Execute "DELETE  FROM TemprptTrialBalance1"
    
    con.Execute "DELETE  FROM subledgertrail where " & stringyear
    OPENINGSUBLEDGERS
    con.Execute "Delete from subledgertrail where " & stringyear
    SUBLEDGERSBALANCE
    
    If AgCombo.text <> "" Then
      DoEvents
      con.Execute "DELETE  FROM TemprptTrialBalance where  subledger not in (SELECT SUBLEDGER FROM SLEDGER left JOIN DISTRICTS ON SLEDGER.DISTCODE = DISTRICTS.DISTRICTNAME where  SLEDGER.setupid=" & setupid & " and SLEDGER.fyear='" & session & "' and Agentname ='" & AgCombo.text & "')"
      DoEvents
      con.Execute "DELETE  FROM subledgertrail where  subledger not in (SELECT SUBLEDGER FROM SLEDGER left JOIN DISTRICTS ON SLEDGER.DISTCODE = DISTRICTS.DISTRICTNAME where   SLEDGER.setupid=" & setupid & " and SLEDGER.fyear='" & session & "' and Agentname ='" & AgCombo.text & "')"
      DoEvents
    End If
    con.Execute "insert into TemprptTrialBalance1 select * from TemprptTrialBalance"
    
    Genrate
    
    
    If Check1_drcr.value = 1 Then
       con.Execute "delete  from winrpt where (Closing=0 and Closing1=0)"
    End If
    
    PrintOption.Show
End If
Commandshow.Enabled = True
End Sub

Private Sub date1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then sendkeys "{tab}"

End Sub

Private Sub date1_LostFocus()
    If Trim(date1.text) <> "" Then
        If Not checkdate(Trim(date1.text), date1) Then
            date1.SetFocus
        End If
    End If
End Sub

Private Sub date2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then sendkeys "{tab}"

End Sub

Private Sub date2_LostFocus()
    If Trim(date2.text) <> "" Then
        If Not checkdate(Trim(date2.text), date2) Then
            date2.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
Me.r1.top = 10
Me.r1.Left = 10

On Error GoTo acc1

Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> (Trim(UCase("MainMenu") Or Trim(UCase("frmbilllist"))))
    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
        Unload VB.Screen.ActiveForm
    End If
Loop

acc1:


Me.top = 0
Me.Left = 0
con.Execute "DELETE  FROM TemprptTrialBalance where " & stringyear
con.Execute "DELETE  FROM subledgertrail where " & stringyear
Set RS = New ADODB.Recordset

    RS.Open "select * from gledger where " & stringyear & " and slf=1", con, adOpenDynamic, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        Do While Not RS.EOF
            COMBOGENLEDGER.AddItem Trim(RS!gledger)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    
     RS.close
   RS.Open "select  Agentname  from AgentMaster where " & stringyear & " order by AgentNAME", con, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not RS.EOF Then
        Do While Not RS.EOF
              If IsNull(RS!agentname) = False Then
                Me.AgCombo.AddItem RS!agentname
            End If
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
   
    RS.close
    CNSetup
    
    If RS.State = 1 Then RS.close
    RS.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly
    If RS.EOF = False Then
        date1.text = RS!yarfrom
        date2.text = RS!yarto
    End If
    
    RS.close
    
    
BackColorFrom Me
    
End Sub

Private Sub print_Click()
Rsinvoicea.Open "select GenLedger,  SubLedger , sum(amount) as INVAmount from invoicea where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and subledger='" + Trim(Combosubledger.text) + "' and INVOICEDATE>=cdate('" + Trim(date1.text) + "') and INVOICEDATE<=cdate('" + Trim(date2.text) + "') order by invoicedate", con, adOpenDynamic, adLockReadOnly, adCmdText
   
       RsCREDITa.Open "select * from CREDITa where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and subledger='" + Trim(Combosubledger.text) + "' and INVOICEDATE>=cdate('" + Trim(date1.text) + "') and INVOICEDATE<=cdate('" + Trim(date2.text) + "')", con, adOpenDynamic, adLockReadOnly, adCmdText
       RsCnf1a.Open "select * from Cnf1a where " & stringyear & " and pgld='" + Trim(COMBOGENLEDGER.text) + "' and psld='" + Trim(Combosubledger.text) + "' and cnd>=cdate('" + Trim(date1.text) + "') and cnd<=cdate('" + Trim(date2.text) + "')", con, adOpenDynamic, adLockReadOnly, adCmdText
       Rsdnfa.Open "select * from dnfa where " & stringyear & " and pgld='" + Trim(COMBOGENLEDGER.text) + "' and psld='" + Trim(Combosubledger.text) + "' and dnd>=cdate('" + Trim(date1.text) + "') and dnd<=cdate('" + Trim(date2.text) + "')", con, adOpenDynamic, adLockReadOnly, adCmdText
       RsCnf1B.Open "select * from Cnf1B where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "' and sld='" + Trim(Combosubledger.text) + "' and cnd>=cdate('" + Trim(date1.text) + "') and cnd<=cdate('" + Trim(date2.text) + "')", con, adOpenDynamic, adLockReadOnly, adCmdText
       RsdnfB.Open "select * from dnfB where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "' and sld='" + Trim(Combosubledger.text) + "' and dnd>=cdate('" + Trim(date1.text) + "') and dnd<=cdate('" + Trim(date2.text) + "')", con, adOpenDynamic, adLockReadOnly, adCmdText
       RScasha.Open "select * from casha where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and subledger='" + Trim(Combosubledger.text) + "' and INVOICEDATE>=cdate('" + Trim(date1.text) + "') and INVOICEDATE<=cdate('" + Trim(date2.text) + "') order by invoicedate", con, adOpenDynamic, adLockReadOnly, adCmdText
  
End Sub
Sub Genrate()
    Command1.top = r1.top + r1.Height + 30
    Combo1.top = r1.top + r1.Height + 30
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
        
        ''con.Execute "delete from Winrpt where uid=" & UId & ""
        con.Execute "delete from Winrpt"
        
        If cboDRCr.text = "DR" Then
           con.Execute "delete from TemprptTrialBalance where Subledger in(SELECT [Subledger] FROM TemprptTrialBalance group by [Subledger] having sum([OpeningBalance]) +sum([DAmount])-sum([CAmount])<0)"
        ElseIf cboDRCr.text = "CR" Then
           con.Execute "delete from TemprptTrialBalance where Subledger in(SELECT [Subledger] FROM TemprptTrialBalance group by [Subledger] having sum([OpeningBalance]) +sum([DAmount])-sum([CAmount])>0)"
        End If
        
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
            If kkk.State = 1 Then kkk.close
            CNSetup
            kkk.Open "select * from setup1 where " & stringyear, con, adOpenDynamic, adLockReadOnly, adCmdText
            If Not kkk.BOF Then
                Print #1, ""
                Print #1, ""
                Print #1, Chr(27) + Chr(15) + Chr(14)
                Print #1, Tab(120); "Page No:  " & Pno
                Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2)); Chr(27) + Chr(15) + Chr(14); Trim(kkk!cname)
                Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(15); dspace(Trim(kkk!add1))
            End If
            If trs.State = 1 Then trs.close
            trs.Open "select * from treport where " & stringyear, con, adOpenDynamic, adLockReadOnly
            Print #1, Tab(((paperWidth - (Len(Trim("Sub Ledger Trial Balance")))) / 2) + LEFTM); "Sub Ledger Trial Balance"
            Print #1, Tab(LEFTM + ((paperWidth - Len(Trim(COMBOGENLEDGER.text))) / 2)); Trim(COMBOGENLEDGER.text)
            xstr = date1.text & " To " & date2.text
            Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("Period : " + Trim(xstr)))) / 2)); Trim("Period : " + date1.text & " To " & date2.text);
            Print #1, Tab(LEFTM); repli("-", paperWidth)
            Print #1, Tab(0); Chr(27) + Chr(71); Tab(8); "Sub. Ledger Description"; Tab(46); "Opening Balance"; Tab(67); "Amount (Dr.)"; Tab(89); "Amount (Cr.)"; Tab(110); "Closing Balance"; Chr(27) + Chr(72)
            Print #1, Tab(LEFTM); repli("-", paperWidth)
            Print #1, ""
            Line = 13
            trs.close
            If called1 Then
                GoTo printagain1
                called1 = False
            End If
            If RS.State = 1 Then RS.close
            RS.Open "select gledger,subledger,sum(openingbalance)as openingbalance1  ,sum(Damount)as Damount1, sum(Camount) as Camount1 , (sum(openingbalance)+sum(Damount)- sum(Camount))  as ClosingBalance    from TemprptTrialBalance where " & stringyear & " and openingbalance<>0   or damount<>0 or cAmount<>0 and userid = " & UId & " group  by gledger,subledger ", con, adOpenStatic, adLockReadOnly, adCmdText
            Dim CB As Double
            
            While Not RS.EOF
                CB = 0
                CB = CB + RS(2) + RS(3) - Abs(RS(4))
                Print #1, Tab(1); RS!subledger; Tab(46); IIf(RS!openingbalance1 <> 0, rsets(Trim(Format(RS!openingbalance1, "0.00")), 12), ""); Tab(65); IIf(RS(3) <> 0, rsets(Trim(Format(RS(3), "0.00")), 12), ""); Tab(85); IIf(RS(4) <> 0, rsets(Trim(Format(Str(RS(4)), "0.00")), 12), ""); Tab(110); IIf(CB <> 0, IIf(CB > 0, rsets(Trim(Format(Str(CB), "0.00")), 12) & "   Dr. ", rsets(Trim(Format(Str(CB), "0.00")), 12) & "      Cr."), "")
                
                con.Execute "insert into winrpt(Party,op,Receipt,Payment,closing,closing1,dr,Description,FromDate,toDate,uid) values('" & Trim(RS!subledger) & "'," & RS!openingbalance1 & "," & RS(3) & "," & RS(4) & "," & IIf(RS!ClosingBalance > 0, RS!ClosingBalance, 0) & "," & IIf(RS!ClosingBalance < 0, RS!ClosingBalance, 0) & ",'" & IIf(RS!ClosingBalance < 0, "Cr", "Dr") & "','" & COMBOGENLEDGER.text & "','" & Format(date1.text, "MM/dd/yyyy") & "','" & Format(date2.text, "MM/dd/yyyy") & "'," & UId & ")"
                
                Line = Line + 1
                GopenBal = GopenBal + RS(2)
                GopenDr = GopenDr + RS(3)
                GopenCr = GopenCr + RS(4)
                GopenCl = GopenCl + RS(2) + RS(3) - Abs(RS(4))
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

