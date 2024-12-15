VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Gentrial 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3276
   ClientLeft      =   252
   ClientTop       =   1416
   ClientWidth     =   7884
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Gtrial1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3276
   ScaleWidth      =   7884
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option1 
      BackColor       =   &H0078CFE9&
      Caption         =   "Crystal Report"
      Height          =   285
      Left            =   2595
      TabIndex        =   19
      Top             =   1500
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0078CFE9&
      Caption         =   "Dos Report"
      Height          =   300
      Left            =   4380
      TabIndex        =   18
      Top             =   1500
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   2100
      Begin VB.OptionButton Option3 
         Caption         =   "Balance Sheet"
         Height          =   300
         Left            =   135
         TabIndex        =   17
         Top             =   225
         Width           =   1500
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Trial Balance"
         Height          =   210
         Left            =   165
         TabIndex        =   16
         Top             =   705
         Width           =   1350
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5820
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton export 
      Height          =   345
      Left            =   1440
      Picture         =   "Gtrial1.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4740
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Gtrial1.frx":045D
      Left            =   2370
      List            =   "Gtrial1.frx":0470
      TabIndex        =   12
      Text            =   "100 %"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton print1 
      Height          =   345
      Left            =   1920
      Picture         =   "Gtrial1.frx":0496
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4710
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   465
      Left            =   4260
      TabIndex        =   2
      Top             =   2040
      Width           =   1545
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      Height          =   465
      Left            =   2580
      TabIndex        =   1
      Top             =   2040
      Width           =   1545
   End
   Begin MSMask.MaskEdBox date1 
      Height          =   345
      Left            =   7830
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2074
      _ExtentY        =   614
      _Version        =   393216
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   7860
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSMask.MaskEdBox alpha 
      Height          =   345
      Left            =   7110
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   585
      _ExtentX        =   1016
      _ExtentY        =   614
      _Version        =   393216
      MaxLength       =   1
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin MSComDlg.CommonDialog c1 
      Left            =   45
      Top             =   4050
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox r1 
      Height          =   2655
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   8985
      _ExtentX        =   15854
      _ExtentY        =   4678
      _Version        =   393217
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      MaxLength       =   99999999
      RightMargin     =   20000
      TextRTF         =   $"Gtrial1.frx":0608
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
      Height          =   345
      Left            =   3480
      TabIndex        =   0
      Top             =   900
      Width           =   1125
      _ExtentX        =   1990
      _ExtentY        =   614
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Alphabat"
      Height          =   195
      Left            =   7140
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "  As On :"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "From The Date"
      Height          =   195
      Left            =   7920
      TabIndex        =   7
      Top             =   1170
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Gen. Ledger Desc."
      Height          =   195
      Left            =   6420
      TabIndex        =   6
      Top             =   270
      Visible         =   0   'False
      Width           =   1350
   End
End
Attribute VB_Name = "Gentrial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As Connection
Dim RS As Recordset
Private Sub alpha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then sendkeys "{tab}"
End Sub
Private Sub COMBOGENLEDGER_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then sendkeys "{tab}"
End Sub
Private Sub COMBOGENLEDGER_LostFocus()
If Trim(COMBOGENLEDGER.text) <> "" Then
    RS.Open "select * from gledger where " & stringyear & "", con, adOpenKeyset, adLockReadOnly, adCmdText
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
    r1.Visible = False
    Me.print1.Visible = False
    Me.export.Visible = False
    Me.Combo1.Visible = False
    Me.Command1.Visible = False
End Sub
Private Sub CommandReturn_Click()
   Unload Me
End Sub
Private Sub Commandshow_Click()
'If Trim(COMBOGENLEDGER.Text) <> "" Then
    If DateDiff("d", Trim(date1.text), Trim(date2.text)) <= 0 Then
        MsgBox "invalid date", , title1
        Exit Sub
    End If
    
    Commandshow.Enabled = False
    
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
    Dim Balance As Double
    Dim OPBALANCE As Double
    Dim SDamount As Double
    Dim SCamount As Double
    Set rs1 = New ADODB.Recordset
    Dim viewsubledger As Boolean
    viewsubledger = False
    Balance = 0
    OPBALANCE = 0
    If RS.State = 1 Then
        RS.close
    End If
    
    
    con.Execute "Delete from TemprptTrialBalance"
 
    con.Execute "INSERT INTO TemprptTrialBalance (Gledger, OpeningBalance,userid,fyear,setupid)  SELECT Gledger.gledger, Gledger.YEAROPENING," & UId & " as Userid,'" & main.session & "'," & main.setupid & " from Gledger where  " & stringyear & ""
    
    RsVDr.Open "SELECT GenLedger,sum(amount) as DAmount  FROM VOUCHERS   WHERE " & stringyear & " and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" & Trim(date2.text) & "',103) and DebitorCredit='D' GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
    RsVCr.Open "SELECT GenLedger,sum(amount) as CAmount  FROM VOUCHERS  where   " & stringyear & " and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" & Trim(date2.text) & "',103) and DebitorCredit='C' GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    rs1.Open "select GenLedger,sum(Netamount) as SAmount from invoicea where " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)  GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    rs2.Open "select GenLedger,  sum(Netamount) as  SAmount from CREDITa where " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)  GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    rs3.Open "select GenLedger,  sum(NETamount) as SAmount  from casha where  " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)and SUBLEDGER <>'CASH PARTY'  GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs3cash.Open "select GenLedger,   sum(baa) as SAmount  from casha where   " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) and SUBLEDGER <>'CASH PARTY'  GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs4.Open "select pgld, sum(na) as SAmount  from Cnf1a where    " & stringyear & " and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)  and dc ='D' GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs4c.Open "select pgld, sum(na) as SAmount  from Cnf1a where  " & stringyear & " and  convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)  and dc ='C' GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs5.Open "select pgld, sum(na) as SAmount from dnfa where  " & stringyear & " and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)  and dc ='D'  GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs5C.Open "select pgld, sum(na) as SAmount from dnfa where   " & stringyear & " and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) and  dc ='C'  GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs6.Open "select gld, sum(a) as SAmount from Cnf1B where   " & stringyear & " and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & Trim(date2.text) & "',103) and  dc ='D'  GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs6C.Open "select gld, sum(a) as SAmount from Cnf1B where   " & stringyear & " and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & Trim(date2.text) & "',103) and dc ='C' GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs7.Open "select gld, sum(a) as SAmount from dnfB where    " & stringyear & " and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & Trim(date2.text) & "',103) and dc ='D' GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs7C.Open "select gld, sum(a) as SAmount from dnfB where    " & stringyear & " and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & Trim(date2.text) & "',103) and dc ='C' GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    
    RS.Open "select * from TemprptTrialBalance where " & stringyear & " and userid=" & main.UId, con, adOpenKeyset, adLockReadOnly, adCmdText
    a = con.Execute("Select sum(Gamount) as aa from InvoiceA  WHERE   " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" & Trim(date2.text) & "',103) ")(0).value
    If IsNull(a) Then a = 0
    con.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & a & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = 'SALES'"
    
    
    a = con.Execute("Select sum(Gamount) as aa from CreditA where " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)")(0).value
    If IsNull(a) Then a = 0
    con.Execute "Update TemprptTrialBalance set DAmount = (DAmount + " & a & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = 'SALES RETURN'"


    a = con.Execute("Select sum(Gamount) as aa from CashA  where   " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)")(0).value
    If IsNull(a) Then a = 0
    
    con.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & a & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = 'SALES'"
    a = con.Execute("Select sum(BAA) as aa from CashA  where   " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)")(0).value
    If IsNull(a) Then a = 0
    con.Execute "Update TemprptTrialBalance set dAmount = (dAmount + " & a & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = 'CASH-IN-HAND'"
    Dim tRS1 As New ADODB.Recordset
    
    '*******For Invoicec
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from INVOICEC  where  " & stringyear & " and Debitorcredit='Debit' AND convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)  GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    If tRS1.RecordCount > 0 Then
    While Not tRS1.EOF
    con.Execute "Update TemprptTrialBalance set DAmount = (DAmount + " & tRS1!amt & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = '" & tRS1!Genledger & "'"
    DoEvents
    DoEvents
    tRS1.MoveNext
    Wend
    End If
   
    If tRS1.State = 1 Then tRS1.close
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from INVOICEC where  " & stringyear & " and Debitorcredit='Credit'  AND convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    If tRS1.RecordCount > 0 Then
        While Not tRS1.EOF
            con.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & tRS1!amt & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = '" & tRS1!Genledger & "'"
            DoEvents
            tRS1.MoveNext
        Wend
    End If
    
    
    '*******For start creditc
    
    If tRS1.State = 1 Then tRS1.close
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CreditC WHERE " & stringyear & " and Debitorcredit='Debit' AND convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)  GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    If tRS1.RecordCount > 0 Then
       While Not tRS1.EOF
          con.Execute "Update TemprptTrialBalance set DAmount = (DAmount + " & tRS1!amt & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = '" & tRS1!Genledger & "'"
          tRS1.MoveNext
       Wend
    End If
    If tRS1.State = 1 Then tRS1.close
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CreditC WHERE   " & stringyear & " and Debitorcredit='Credit' AND convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)   GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    If tRS1.RecordCount > 0 Then
        While Not tRS1.EOF
                    con.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & tRS1!amt & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = '" & tRS1!Genledger & "'"
            tRS1.MoveNext
        Wend
    End If
    
    '******* End of CreditC
    
    
    '*******For CashC
    If tRS1.State = 1 Then tRS1.close
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CashC WHERE   " & stringyear & " and Debitorcredit='Debit' AND convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    If tRS1.RecordCount > 0 Then
       While Not tRS1.EOF
               con.Execute "Update TemprptTrialBalance set DAmount = (DAmount + " & tRS1!amt & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = '" & tRS1!Genledger & "'"
          tRS1.MoveNext
       Wend
    End If
    If tRS1.State = 1 Then tRS1.close
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CashC WHERE   " & stringyear & " and Debitorcredit='Credit' AND convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    If tRS1.RecordCount > 0 Then
        While Not tRS1.EOF
                   con.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & tRS1!amt & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = '" & tRS1!Genledger & "'"
            tRS1.MoveNext
        Wend
    End If
    
    '*******For End of CashC
    If Not RS.BOF Then
           Do While Not RS.EOF
               
 
               
               OPBALANCE = RS!OpeningBalance
               SDamount = 0
               SCamount = 0
               
               
               If RS!gledger = "REBATE & DISCOUNT" Then
               'MsgBox "a"
               End If
               
               
            
            If RsVDr.RecordCount > 0 Then
                  RsVDr.Find "Genledger='" + Trim(RS!gledger) + "'"
                  If Not RsVDr.EOF Then
                        SDamount = SDamount + RsVDr!damount
                 End If
            End If

            If RsVCr.RecordCount > 0 Then
                  RsVCr.Find "Genledger='" + Trim(RS!gledger) + "'"
                  If Not RsVCr.EOF Then
                        SCamount = SCamount + RsVCr!camount
                 End If
            End If

            If rs1.RecordCount > 0 Then
                  rs1.Find "Genledger='" + Trim(RS!gledger) + "'"
                  If Not rs1.EOF Then
                      SDamount = SDamount + rs1!samount
                  End If
            End If
            
            If rs2.RecordCount > 0 Then
                  rs2.Find "Genledger='" + Trim(RS!gledger) + "'"
                  If Not rs2.EOF Then
                      SCamount = SCamount + rs2!samount
                  End If
            End If

            If rs3.RecordCount > 0 Then
                  rs3.Find "Genledger='" + Trim(RS!gledger) + "'"
                  If Not rs3.EOF Then
                      SDamount = SDamount + rs3!samount
                  End If
            End If
            
            If rs3cash.RecordCount > 0 Then
                  rs3cash.Find "Genledger='" + Trim(RS!gledger) + "'"
                  If Not rs3cash.EOF Then
                      SCamount = SCamount + rs3cash!samount
                  End If
            End If
            
            If rs4.RecordCount > 0 Then
                  rs4.Find "pgld ='" + Trim(RS!gledger) + "'"
                  If Not rs4.EOF Then
                      SDamount = SDamount + rs4!samount
                  End If
            End If
            
            If rs4c.RecordCount > 0 Then
                  rs4c.Find "pgld= '" + RS!gledger + "'"
                  If Not rs4c.EOF Then
                      SCamount = SCamount + rs4c!samount
                  End If
            End If
            
            
            If rs5.RecordCount > 0 Then
                  rs5.Find "pgld='" + Trim(RS!gledger) + "'"
                  If Not rs5.EOF Then
                      SDamount = SDamount + rs5!samount
                  End If
            End If

            If rs5C.RecordCount > 0 Then
                  rs5C.Find "pgld='" + Trim(RS!gledger) + "'"
                  If Not rs5C.EOF Then
                      SCamount = SCamount + rs5C!samount
                  End If
            End If
            
            If rs6.RecordCount > 0 Then
                  rs6.Find "gld ='" + Trim(RS!gledger) + "'"
                  If Not rs6.EOF Then
                      SDamount = SDamount + rs6!samount
                  End If
            End If
            
            If rs6C.RecordCount > 0 Then
                  rs6C.Find "gld='" + Trim(RS!gledger) + "'"
                  If Not rs6C.EOF Then
                      SCamount = SCamount + rs6C!samount
                  End If
            End If
  
            If rs7.RecordCount > 0 Then
                  rs7.Find "gld='" + Trim(RS!gledger) + "'"
                  If Not rs7.EOF Then
                      SDamount = SDamount + rs7!samount
                  End If
            End If
            If rs7C.RecordCount > 0 Then
                  rs7C.Find "gld='" + Trim(RS!gledger) + "'"
                  If Not rs7C.EOF Then
                      SCamount = SCamount + rs7C!samount
                  End If
            End If
  
  
            
            
            con.Execute "UPDATE TemprptTrialBalance SET TemprptTrialBalance.DAmount = TemprptTrialBalance.DAmount +  " & SDamount & " ,TemprptTrialBalance.CAmount = TemprptTrialBalance.CAmount+  " & SCamount & "  where  " & stringyear & " and userid=" & main.UId & " and Gledger ='" + Trim(RS!gledger) + "'"
            
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
            
            If Not RS.EOF Then
                RS.MoveNext
            End If
      
      Loop
    End If
   
   
   
   
   
   Genrate
   
   con.Execute "delete from Winrpt where (Closing=0 and Closing1=0)"
   
   PrintOption.Show
   
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

BackColorFrom Me

Me.Option1.value = True
Me.Option4.value = True
Me.r1.top = 10
Me.r1.Left = 10

Me.top = 0
Me.Left = 0

Set RS = New ADODB.Recordset

    RS.Open "select * from gledger where  " & stringyear & "", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    If Not RS.BOF Then
        Do While Not RS.EOF
            COMBOGENLEDGER.AddItem Trim(RS!gledger)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    

    RS.close
    CNSetup
    RS.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockReadOnly
    date1.text = RS!yarfrom
    date2.text = RS!yarto
    RS.close
End Sub

Private Sub print_Click()
Rsinvoicea.Open "select GenLedger,  SubLedger , sum(amount) as INVAmount from invoicea where  " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and subledger='" + Trim(Combosubledger.text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) order by invoicedate", con, adOpenKeyset, adLockReadOnly, adCmdText
  RsCREDITa.Open "select * from CREDITa where  " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and subledger='" + Trim(Combosubledger.text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)", con, adOpenKeyset, adLockReadOnly, adCmdText
       RsCnf1a.Open "select * from Cnf1a where  " & stringyear & " and pgld='" + Trim(COMBOGENLEDGER.text) + "' and psld='" + Trim(Combosubledger.text) + "' and convert(smalldatetime,cnd,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)", con, adOpenKeyset, adLockReadOnly, adCmdText
       Rsdnfa.Open "select * from dnfa where  " & stringyear & " and pgld='" + Trim(COMBOGENLEDGER.text) + "' and psld='" + Trim(Combosubledger.text) + "' and convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)", con, adOpenKeyset, adLockReadOnly, adCmdText
       RsCnf1B.Open "select * from Cnf1B where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "' and sld='" + Trim(Combosubledger.text) + "' and convert(smalldatetime,cnd,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)", con, adOpenKeyset, adLockReadOnly, adCmdText
       RsdnfB.Open "select * from dnfB where  " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "' and sld='" + Trim(Combosubledger.text) + "' and convert(smalldatetime,dnd,103) >=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103)", con, adOpenKeyset, adLockReadOnly, adCmdText
       RScasha.Open "select * from casha where  " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and subledger='" + Trim(Combosubledger.text) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.text) + "',103) order by invoicedate", con, adOpenKeyset, adLockReadOnly, adCmdText
  
End Sub
Sub Genrate()
    print1.top = r1.top + r1.Height + 30
    export.top = r1.top + r1.Height + 30
    Command1.top = r1.top + r1.Height + 30
    Combo1.top = r1.top + r1.Height + 30
    Set RS = New ADODB.Recordset
    main.reportname = "Gen. Ledger Trial"
    Dim called1, called2 As Boolean
    Dim MaxLine As Integer
    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
    Dim paperWidth As Integer
    Dim xtemp As String
    Dim trs As ADODB.Recordset
    Set trs = New ADODB.Recordset
    
    con.Execute "delete from Winrpt where UID=" & UId & ""
    
    paperWidth = 146
        T1 = 10
        T2 = 25
        T3 = 40
        T4 = 55
        T5 = 70
        T6 = 85
        T7 = 100
        T8 = 115
        MaxLine = 50
        called1 = False
        called2 = False
        Dim Line As Integer
        Dim rs1 As ADODB.Recordset
        Dim kkk As ADODB.Recordset
        Dim Pno As Integer
        Dim GopenBal As Double
        Dim GopenDr As Double
        Dim GopenCr As Double
        Dim GopenCl As Double
        
        Dim GSumDr  As Double
        Dim GSumCr  As Double
        Dim FooterYes As Boolean
        GopenBal = 0
        GopenDr = 0
        GopenCr = 0
        GopenCl = 0
        GSumDr = 0
        GSumCr = 0
        Set kkk = New ADODB.Recordset
        Set rs1 = New ADODB.Recordset
        main.reportdata
        main.repors.Find "reportname='" + Trim(main.reportname) + "'"
        MaxLine = main.repors!totalline
        If main.repors!comp = True Then
            paperWidth = Int(main.repors!totalcolumn * 1.75)
        Else
            paperWidth = main.repors!totalcolumn
        End If
        Open "" + App.Path + "\vipin.txt" For Output As #1
        Line = 0
        Pno = 1
        FooterYes = False
header:
        Dim I As Integer
        For I = 1 To main.repors!TopMargin
            Print #1, ""
            Line = Line + 1
        Next
        If FooterYes = True Then
           Do While Line <= 72
               Print #1, ""
               Line = Line + 1
           Loop
           Line = 0
           FooterYes = False
        End If
        
        If kkk.State = 1 Then kkk.close
        CNSetup
        kkk.Open "Select * from setup1 where " & stringyear & "", con, adOpenKeyset, adLockReadOnly, adCmdText
        If Not kkk.BOF Then
           Print #1, ""
           Print #1, Chr(27) + Chr(15) + Chr(14)
           Print #1, Tab(115); "Page No:  " & Pno
           Print #1, Tab((((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15) + LEFTM); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!cname))
           Print #1, Tab(((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15); dspace(Trim(kkk!add1)); Chr(27) + Chr(14)
           Line = Line + 5
        End If
        If trs.State = 1 Then trs.close
        
        trs.Open "select * from treport where userid=" & UId & "", con, adOpenKeyset, adLockReadOnly, adCmdText
        Print #1, Tab(((paperWidth - (Len(Trim("Gen. Ledger Trial Balance")) * 2)) / 2) + LEFTM); Chr(27) + Chr(15) + Chr(14); "Gen. Ledger Trial Balance"
        Print #1, Tab(LEFTM + ((paperWidth - Len(Trim("As on : " & Gentrial.date2.text))) / 2)); Trim("As on : " & Gentrial.date2.text)
        Print #1, ""
        Line = Line + 3
        Print #1, Tab(LEFTM); Chr(27) + Chr(71); repli("-", paperWidth)
        Print #1, Tab(8); "Gen. Ledger Description"; Tab(67); "Amount (Dr.)"; Tab(105); "Amount (Cr.)";
        Print #1, Tab(LEFTM); repli("-", paperWidth); Chr(27) + Chr(72)
        Print #1, ""
        Line = Line + 4
        If trs.State = 1 Then trs.close
        If called1 Then
           called1 = False
           GoTo printagain1
        End If
        If RS.State = 1 Then RS.close
        Dim DbB As Double
        Dim CrB As Double
        DbB = 0
        DrB = 0
        If RS.State = 1 Then RS.close
        RS.Open "select * from TemprptTrialBalance where   " & stringyear & " and DAmount<>0 or CAmount<>0 or openingbalance <>0 and userid = " & UId & " order by gledger ", con, adOpenKeyset, adLockReadOnly, adCmdText
        
        Dim CB As Double
           While Not RS.EOF
              DbB = 0
              DbB = IIf(IsNull(RS!damount), 0, RS!damount) - IIf(IsNull(RS!camount), 0, RS!camount) + RS!OpeningBalance
              Print #1, Tab(1); RS!gledger; Tab(65); IIf(DbB > 0, rsets(Trim(Format(DbB, "0.00")), 13), ""); Tab(102); IIf(DbB < 0, rsets(Trim(Format(Str(Abs(DbB)), "0.00")), 13), "")
              con.Execute "insert into winrpt(Party,Narration,op,Receipt,Payment,closing,closing1,Description,FromDate,toDate,uid,OpDes) values('" & RS!gledger & "','" & RS!gledger & "'," & RS!OpeningBalance & "," & 0 & "," & 0 & "," & IIf(DbB > 0, DbB, 0) & "," & IIf(DbB < 0, DbB, 0) & ",'" & "Gen. Ledger Trial Balance" & "','" & Format(date1.text, "MM/dd/yyyy") & "','" & Format(date2.text, "MM/dd/yyyy") & "'," & UId & ",'" & date2.text & "')"
              Line = Line + 1
              If Line > MaxLine - 7 Then
                        called1 = True
                        FooterYes = True
                        Pno = Pno + 1
                        GoTo header
printagain1:
                       
                        called1 = False
              End If
                
              If Not RS.EOF Then
                    RS.MoveNext
              End If
              If DbB > 0 Then
                   GSumDr = GSumDr + Abs(DbB)
              Else
                  GSumCr = GSumCr + Abs(DbB)
              End If
            Wend
printfooter:
            bal = GSumDr - GSumCr
            If bal < 0 And bal <> 0 Then
               Print #1, ""
               'Print #1, Tab(LEFTM); "NET DIFFERENCE "; Tab(65); rsets(Format(Trim(Abs(bal)), "0.00"), 12)
               Line = Line + 1
            Else
              If bal <> 0 Then

                 Print #1, ""
                'Print #1, Tab(LEFTM); "NET DIFFERENCE "; Tab(102); rsets(Format(Trim(Abs(bal)), "0.00"), 12)
                Line = Line + 2
              End If
            End If
            If GSumDr > GSumCr Then
               neta = GSumDr
            Else
               neta = GSumCr
            End If
            
            Print #1, Tab(LEFTM); Chr(27) + Chr(71); repli("-", paperWidth)
            Print #1, Tab(LEFTM); "* * * NET TOTAL * * * "; Tab(65); IIf(GSumDr <> 0, rsets(Format(Trim(GSumDr), "0.00"), 13), ""); Tab(102); IIf(GSumCr <> 0, rsets(Format(Trim(GSumCr), "0.00"), 13), "");
            'Print #1, Tab(LEFTM); "* * * NET TOTAL * * * "; Tab(65); IIf(neta <> 0, rsets(Format(Trim(neta), "0.00"), 12), ""); Tab(102); IIf(neta <> 0, rsets(Format(Trim(neta), "0.00"), 12), "");
            
            Print #1, Tab(LEFTM); repli("-", paperWidth); Chr(27) + Chr(72)
            Line = Line + 3
            Do While Line <= 72
               Print #1, " "
               Line = Line + 1
            Loop
            If trs.State = 1 Then trs.close
            Close #1
End Sub

Private Sub print1_Click()
    c1.PrinterDefault = True
    c1.ShowPrinter
    printnow
End Sub

