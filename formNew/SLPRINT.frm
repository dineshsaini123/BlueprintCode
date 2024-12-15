VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form SLEDGERPRINT 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4752
   ClientLeft      =   252
   ClientTop       =   1416
   ClientWidth     =   8856
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4752
   ScaleWidth      =   8856
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1_cash 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cash Party Statement"
      Height          =   504
      Left            =   2052
      Picture         =   "SLPRINT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3816
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CheckBox Check1_all 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   7596
      TabIndex        =   17
      Top             =   864
      Visible         =   0   'False
      Width           =   1164
   End
   Begin VB.ListBox cboacc_multiple 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2136
      Left            =   2016
      Style           =   1  'Checkbox
      TabIndex        =   16
      Top             =   900
      Visible         =   0   'False
      Width           =   5556
   End
   Begin VB.CheckBox Check1_select 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Multiple A/C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5364
      TabIndex        =   15
      Top             =   576
      Visible         =   0   'False
      Width           =   2136
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   372
      Left            =   168
      TabIndex        =   13
      Top             =   4344
      Visible         =   0   'False
      Width           =   8052
      _ExtentX        =   14203
      _ExtentY        =   656
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   645
      Left            =   8916
      TabIndex        =   12
      Top             =   2475
      Width           =   1005
   End
   Begin VB.TextBox Alpha 
      Height          =   315
      Left            =   2028
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1548
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.ComboBox Combosubledger 
      Height          =   288
      Left            =   2028
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   870
      Visible         =   0   'False
      Width           =   5565
   End
   Begin VB.CommandButton Commandreturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   720
      Left            =   3720
      Picture         =   "SLPRINT.frx":0BE4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3096
      Width           =   1590
   End
   Begin VB.CommandButton Commandshow 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Show"
      Height          =   720
      Left            =   2040
      Picture         =   "SLPRINT.frx":17C8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3096
      Width           =   1590
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   288
      Left            =   2028
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   156
      Width           =   5565
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
      Height          =   312
      Left            =   2028
      TabIndex        =   3
      Top             =   2112
      Width           =   1152
      _ExtentX        =   2032
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
      Height          =   312
      Left            =   4548
      TabIndex        =   4
      Top             =   2112
      Width           =   1152
      _ExtentX        =   2032
      _ExtentY        =   550
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Search Ledger A/c"
      Height          =   192
      Left            =   2028
      TabIndex        =   14
      Top             =   636
      Width           =   3348
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From The Date"
      Height          =   312
      Left            =   180
      TabIndex        =   11
      Top             =   2160
      Width           =   1992
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " - To - "
      Height          =   312
      Left            =   3672
      TabIndex        =   10
      Top             =   2160
      Width           =   588
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Alphabat"
      Height          =   288
      Left            =   180
      TabIndex        =   9
      Top             =   1572
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub. Ledger Desc."
      Height          =   288
      Left            =   180
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gen. Ledger Desc."
      Height          =   288
      Left            =   180
      TabIndex        =   7
      Top             =   156
      Width           =   2052
   End
End
Attribute VB_Name = "SLEDGERPRINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As Recordset
Dim CON_next As New ADODB.Connection

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
    sendkeys "{TAB}"
End If
End Sub

Private Sub Check1_all_Click()
  
  If Check1_all.value = 1 Then
     For k1 = 0 To cboacc_multiple.ListCount - 1
         cboacc_multiple.Selected(k1) = True
     Next
  Else
     For k1 = 0 To cboacc_multiple.ListCount - 1
         cboacc_multiple.Selected(k1) = False
     Next

  End If
  
End Sub

Private Sub Check1_select_Click()
 If Check1_select.value = 1 Then
        Check1_all.Visible = True
        cboacc_multiple.Visible = True
        cboacc_multiple.Clear
        
        Screen.MousePointer = vbHourglass
        
        If RS.State = 1 Then
        RS.close
        End If
        RS.Open "select * from sledger where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "' and len(email)>10", con, adOpenStatic, adLockReadOnly, adCmdText
        Combosubledger.Clear
        If Not RS.BOF Then
        Do While Not RS.EOF
            cboacc_multiple.AddItem (Trim(RS!subledger))
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
        End If
        RS.close
        
        Screen.MousePointer = vbDefault
        
       Else
        cboacc_multiple.Visible = False
        Check1_all.Visible = False
 End If
End Sub

Private Sub COMBOGENLEDGER_Change()
   
      
    

    If RS.State = 1 Then
        RS.close
    End If
    RS.Open "select * from sledger where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
    Combosubledger.Clear
    If Not RS.BOF Then
        
      
    
        Do While Not RS.EOF
            Combosubledger.AddItem Trim(RS!subledger)
            
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    
End Sub

Private Sub COMBOGENLEDGER_Click()

If Trim(COMBOGENLEDGER.text) = "IMPREST A/C" Then
   Check1_select.Visible = True
Else
   Check1_select.Visible = False
End If

cboacc_multiple.Clear
If Check1_select.value = 1 Then
   cboacc_multiple.Visible = True
   cboacc_multiple.Clear
  Else
   cboacc_multiple.Visible = False
End If

If RS.State = 1 Then
        RS.close
    End If
    RS.Open "select * from sledger where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
    Combosubledger.Clear
    If Not RS.BOF Then
        Do While Not RS.EOF
            Combosubledger.AddItem Trim(RS!subledger)
            cboacc_multiple.AddItem (Trim(RS!subledger))
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    
    
    
    
End Sub

Private Sub COMBOGENLEDGER_KeyPress(KeyAscii As Integer)
If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
   sendkeys "{tab}"
   Exit Sub
End If
If KeyAscii = 13 Then
   sendkeys "{DOWN}"
   sendkeys "{tab}"
End If

End Sub

Private Sub COMBOGENLEDGER_LostFocus()
    If Trim(COMBOGENLEDGER.text) <> "" Then
        RS.Open "select * from gledger where " & stringyear & " and slf=1", con, adOpenStatic, adLockReadOnly, adCmdText
        If Not RS.BOF Then
            RS.Find "gledger='" + Trim(COMBOGENLEDGER.text) + "'"
            If RS.EOF Then
                COMBOGENLEDGER.SetFocus
            End If
        Else
            COMBOGENLEDGER.SetFocus
        End If
        RS.close
        
       If Trim(COMBOGENLEDGER.text) = "IMPREST A/C" Then
          Check1_select.Visible = True
       Else
          Check1_select.Visible = False
       End If
       
        
    End If
End Sub

Private Sub Combosubledger_GotFocus()
    If Trim(COMBOGENLEDGER.text) = "" Then
        COMBOGENLEDGER.SetFocus
    End If
    
    If PopUpValue1 <> "" Then
       Combosubledger.text = PopUpValue2
       PopUpValue1 = ""
       PopUpValue2 = ""
    End If
    
End Sub
Private Sub Combosubledger_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
   searchType = "party"
   popuplistModel10 "select Party,Subledger from Sledger where gledger='" & COMBOGENLEDGER.text & "' and " & stringyear & "  order by Party", con
   
End If

End Sub

Private Sub Combosubledger_KeyPress(KeyAscii As Integer)
If VB.Screen.ActiveForm.ActiveControl.SelLength > 0 Then
   sendkeys "{tab}"
   Exit Sub
End If
If KeyAscii = 13 Then
   sendkeys "{Down}"
   sendkeys "{tab}"
End If
End Sub

Private Sub Combosubledger_LostFocus()
If Trim(Combosubledger.text) <> "" Then
    If Trim(COMBOGENLEDGER.text) <> "" Then
        If RS.State = 1 Then
            RS.close
        End If
        RS.Open "select * from sledger where " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "' and subledger='" + Trim(Combosubledger.text) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
        If RS.BOF Then
            Combosubledger.SetFocus
        End If
        RS.close
    Else
        Combosubledger.text = ""
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
    con.Execute "Delete  from subledgertrail where userid=" & UId & ""
    con.Execute ("delete  from treport where userid=" & UId & "")
    
    If DateDiff("d", Trim(date1.text), Trim(date2.text)) <= 0 Then
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
    OPENINGSUBLEDGERS
    DoEvents
    If Trim(Alpha.text) <> "" And Alpha.Visible = True Then
      ' vouchers creditors
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & "," & setupid & ",'" & session & "' From VOUCHERS Where   " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger like '" + Trim(Alpha.text) + "%'  AND   VOUCHERS.DebitorCredit='C' and convert(smalldatetime,VoucherDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,VoucherDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
      ' vouchers debtors
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & "," & setupid & ",'" & session & "' From VOUCHERS Where   " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger like '" + Trim(Alpha.text) + "%' AND  VOUCHERS.DebitorCredit='D' and  convert(smalldatetime,VoucherDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,VoucherDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)      ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
      ' invoice
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear )  SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & "," & setupid & ",'" & session & "' FROM INVOICEA  where   " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger like '" + Trim(Alpha.text) + "%'  and convert(smalldatetime,InvoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,InvoiceDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103) "
      ' cash credit
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & "," & setupid & ",'" & session & "'  FROM CASHA  where   " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger like '" + Trim(Alpha.text) + "%'   and convert(smalldatetime,InvoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,InvoiceDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)  "
      ' cash debit
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear )   SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & "," & setupid & ",'" & session & "'  FROM CASHA  where " & stringyear & " and CASHA.BAA<>0 and  genledger='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger like '" + Trim(Alpha.text) + "%'   and convert(smalldatetime,InvoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,InvoiceDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103) "
       ' credit a
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear ) SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note(I)' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3, " & UId & "," & setupid & ",'" & session & "' FROM CREDITA  where   " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger like '" + Trim(Alpha.text) + "%' and convert(smalldatetime,InvoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,InvoiceDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)   "
       ' dnfadr
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear )  SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, 'Debit Note', DNFA.Na, DNFA.DC, '', " & UId & "," & setupid & ",'" & session & "'   From DNFA  where   " & stringyear & " and Pgld ='" + Trim(COMBOGENLEDGER.text) + "' and  Psld like '" + Trim(Alpha.text) + "%'  and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,dnd,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
       'cnf1cr
      '''CON.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid )  SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, CNF1A.N, CNF1A.NA, CNF1A.DC, '', " & UId & "  From CNF1A where " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.Text) + "' and  Psld like '" + Trim(alpha.Text) + "%'  and convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.Text) + "',103)   and convert(smalldatetime,cnd,103) <=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear )  SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, 'Credit Note', CNF1A.NA, CNF1A.DC, '', " & UId & "," & setupid & ",'" & session & "'  From CNF1A where " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.text) + "' and  Psld like '" + Trim(Alpha.text) + "%'  and convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,cnd,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103) ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"  'and ReflectInAcc=0
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear )  SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, '' , " & UId & "," & setupid & ",'" & session & "' From DNFB  where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "' and  sld like '" + Trim(Alpha.text) + "%'   and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,dnd,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
      con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear )   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '', " & UId & "," & setupid & ",'" & session & "' From CNF1B where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "' and  sld like '" + Trim(Alpha.text) + "%'   and convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,cnd,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)   ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
   End If
   
   
    If Trim(Alpha.text) = "" And Alpha.Visible = True Then
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear ) SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & "," & setupid & ",'" & session & "' From VOUCHERS Where   " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.text) + "'  AND   VOUCHERS.DebitorCredit='C' and  convert(smalldatetime,VoucherDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)  and convert(smalldatetime,VoucherDate,103) <= convert(smalldatetime,'" + Trim(date2.text) + "',103)   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND , " & UId & "," & setupid & ",'" & session & "' From VOUCHERS Where   " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.text) + "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VoucherDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)  and convert(smalldatetime,VoucherDate,103) <= convert(smalldatetime,'" + Trim(date2.text) + "',103)   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid,setupid,fyear) SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3  , " & UId & "," & setupid & ",'" & session & "' FROM INVOICEA  where   " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,invoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)  and convert(smalldatetime,invoiceDate,103) <= convert(smalldatetime,'" + Trim(date2.text) + "',103) "
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid,setupid,fyear) SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & "," & setupid & ",'" & session & "' FROM CASHA   where   " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,invoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)  and convert(smalldatetime,invoiceDate,103) <= convert(smalldatetime,'" + Trim(date2.text) + "',103)"
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear )  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & "," & setupid & ",'" & session & "'  FROM CASHA  where  " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "'   and convert(smalldatetime,invoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)  and convert(smalldatetime,invoiceDate,103) <= convert(smalldatetime,'" + Trim(date2.text) + "',103)  AND CASHA.BAA <>0  "
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear ) SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note(I)' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3, " & UId & "," & setupid & ",'" & session & "'  FROM CREDITA  where   " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "'  and convert(smalldatetime,invoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)  and convert(smalldatetime,invoiceDate,103) <= convert(smalldatetime,'" + Trim(date2.text) + "',103)"
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear ) SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, DNFA.N, DNFA.Na, DNFA.DC, ''  , " & UId & "," & setupid & ",'" & session & "' From DNFA  where   " & stringyear & " and Pgld ='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)  and convert(smalldatetime,dnd,103) <= convert(smalldatetime,'" + Trim(date2.text) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear ) SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' , CNF1A.CNN, 'Credit Note', CNF1A.NA, CNF1A.DC, '' , " & UId & "," & setupid & ",'" & session & "' From CNF1A where " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.text) + "'  and convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)  and convert(smalldatetime,cnd,103) <= convert(smalldatetime,'" + Trim(date2.text) + "',103)  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"   'and ReflectInAcc=0
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear ) SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, '', " & UId & "," & setupid & ",'" & session & "'  From DNFB  where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "'  and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)  and convert(smalldatetime,dnd,103) <= convert(smalldatetime,'" + Trim(date2.text) + "',103)   ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear ) SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & "," & setupid & ",'" & session & "' From CNF1B where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)  and convert(smalldatetime,cnd,103) <= convert(smalldatetime,'" + Trim(date2.text) + "',103) ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
End If
If Trim(Alpha.text) = "" And Alpha.Visible = False And Combosubledger.text <> "" Then
                
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid,setupid,fyear )  SELECT VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & "," & setupid & ",'" & session & "' From VOUCHERS Where   " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger = '" & Combosubledger.text & "'  AND   VOUCHERS.DebitorCredit='C' and  convert(smalldatetime,VoucherDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,VoucherDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)   ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid, setupid,fyear)   SELECT GenLedger, SubLedger, VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber, VOUCHERS.DESCRIPTION, VOUCHERS.Amount, VOUCHERS.DebitorCredit, VOUCHERS.CBND, " & UId & "," & setupid & ",'" & session & "' From VOUCHERS Where   " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger = '" & Combosubledger.text & "' AND  VOUCHERS.DebitorCredit='D' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VoucherDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,VoucherDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)    ORDER BY VOUCHERS.GenLedger, VOUCHERS.SubLedger, VOUCHERS.VoucherDate, VOUCHERS.VoucherType, VOUCHERS.VoucherNumber,vsno"
                
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,setupid,fyear)  SELECT INVOICEA.GENLEDGER, INVOICEA.SUBLEDGER, INVOICEA.INVOICEDATE, 'I' AS Expr1, INVOICEA.INVOICENO, 'Sales Invoice' , INVOICEA.NETAMOUNT, 'D' , '' AS Expr3 , " & UId & "," & setupid & ",'" & session & "' FROM INVOICEA  where   " & stringyear & " and genledger ='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger = '" & Combosubledger.text & "'  and convert(smalldatetime,InvoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,InvoiceDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103) "
    
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,setupid,fyear)  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.NETAMOUNT, 'D' , '' AS Expr3, " & UId & "," & setupid & ",'" & session & "'  FROM CASHA  where   " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger = '" & Combosubledger.text & "'   and convert(smalldatetime,InvoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,InvoiceDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)   "
                
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,setupid,fyear)  SELECT CASHA.GENLEDGER, CASHA.SUBLEDGER, CASHA.INVOICEDATE, 'S' AS Expr1, CASHA.INVOICENO, 'Sales C/M' , CASHA.BAA, 'C' , '' AS Expr3, " & UId & "," & setupid & ",'" & session & "'  FROM CASHA  where   " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger = '" & Combosubledger.text & "' and convert(smalldatetime,InvoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,InvoiceDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)  AND CASHA.BAA <>0  "
                
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,setupid,fyear)    SELECT CREDITA.GENLEDGER, CREDITA.SUBLEDGER, CREDITA.INVOICEDATE, 'C' AS Expr1, CREDITA.INVOICENO, 'Credit Note(I)' , CREDITA.NETAMOUNT, 'C' , '' AS Expr3 , " & UId & "," & setupid & ",'" & session & "' FROM CREDITA  where   " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.text) + "' and  Subledger = '" & Combosubledger.text & "'    and convert(smalldatetime,InvoiceDate,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103) and convert(smalldatetime,InvoiceDate,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103) "

                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno,userid ,setupid,fyear)   SELECT DNFA.PGLD, DNFA.PSLD, DNFA.DND, 'D' , DNFA.DNN, 'Debit Note', DNFA.Na, DNFA.DC, '' , " & UId & "," & setupid & ",'" & session & "'   From DNFA  where   " & stringyear & " and Pgld ='" + Trim(COMBOGENLEDGER.text) + "' and  Psld = '" & Combosubledger.text & "' and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,dnd,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)  ORDER BY DNFA.PGLD, DNFA.PSLD, DNFA.DND, DNFA.DNN "
 
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid,setupid,fyear )   SELECT CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND, 'C' ,  CNF1A.CNN, 'Credit Note' , CNF1A.NA, CNF1A.DC, '' , " & UId & "," & setupid & ",'" & session & "' From CNF1A where " & stringyear & " and Pgld='" + Trim(COMBOGENLEDGER.text) + "' and  Psld = '" & Combosubledger.text & "'    and convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,cnd,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)  ORDER BY CNF1A.PGLD, CNF1A.PSLD, CNF1A.CND,CNF1A.CNN"   'and ReflectInAcc=0
 
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno, userid ,setupid,fyear)   SELECT DNFB.GLD, DNFB.SLD, DNFB.DND, 'D' , DNFB.DNN, 'DNB' , DNFB.A, DNFB.DC, ''     , " & UId & "," & setupid & ",'" & session & "' From DNFB  where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "' and  sld = '" & Combosubledger.text & "'   and convert(smalldatetime,dnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,dnd,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103)    ORDER BY DNFB.GLD, DNFB.SLD, DNFB.DND, DNFB.DNN"
                
                con.Execute "INSERT INTO treport ( genledger, subledger, vdate, vtype, vno, narration, ad, dorc, cbno ,userid ,setupid,fyear)   SELECT CNF1B.GLD, CNF1B.SLD, CNF1B.CND, 'C' , CNF1B.CNN, 'CNB' , CNF1B.A, CNF1B.DC, '' , " & UId & "," & setupid & ",'" & session & "' From CNF1B where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.text) + "' and  sld = '" & Combosubledger.text & "' and  convert(smalldatetime,cnd,103) >= convert(smalldatetime,'" + Trim(date1.text) + "',103)   and convert(smalldatetime,cnd,103) <=convert(smalldatetime,'" + Trim(date2.text) + "',103) ORDER BY    CNF1B.GLD, CNF1B.SLD, CNF1B.CND, CNF1B.CNN"
End If
con.Execute "insert into Treport ( Genledger,Subledger,openingbalance,userid,setupid,fyear ) SELECT '" + Trim(COMBOGENLEDGER.text) + "'as genled, SUBLEDGER, (sum(YEAROPENING) + SUM (OPAMOUNTDEBIT) - SUM(OPAMOUNTCREDIT)) AS OPCR ,  " & UId & " as UserId," & setupid & ",'" & session & "' from subledgertrail where userid=" & UId & " and " & stringyear & " GROUP BY SUBLEDGER;"
main.reportname = "Sub. Ledger"



''''''New code
''''
'''''''
''''
'''''kk = 1
''''
''''
''''
''''
''''kk = kk + 1
''''
''''''''
''''
''''If PopUpValue6 = 1 Then
''''
''''con.Execute "update treport set tab='0'"
''''If RS.State = 1 Then RS.Close
''''RS.Open "select subledger from vouchers where [DESCRIPTION] like 'CASH RECD%' and Subledger like '" + Trim(Alpha.text) + "%'  group by subledger", CON_next
''''
'''''RS.Open "select subledger from treport where narration like '" + "CASH REC" + "%'  group by subledger", con
''''
''''
''''While RS.EOF = False
''''
''''con.Execute "update treport set tab='1' where subledger='" & RS(0) & "'"
''''
''''RS.MoveNext
''''Wend
''''
''''con.Execute "delete from treport  where tab='0'"
''''
''''
''''End If
'''''end code



If SLEDGERPRINT.Alpha.Visible = True Then
   viewledger.genreport1
Else
   viewledger.genreport1
End If
PrintOption.Show
End Sub

Private Sub Command1_cash_Click()


'
'
'
'PopUpValue6 = 1
'
'If Check1_select.value = 0 Then
'
' Commandshow.Enabled = False
''********sub for alpha wise and Partywise according to new fast mathed
' DoEvents
'
' ''CON.Execute "Delete  from subledgertrail where " & stringyear
' DoEvents
' DoEvents
' DoEvents
' DoEvents
' DoEvents
' DoEvents
' ALPHAB
' 'CON.CommitTrans
' Commandshow.Enabled = True
'
' Else
'
' PrintOption.Show
'
' End If

End Sub

Private Sub CommandReturn_Click()
    Unload Me
End Sub
Private Sub Commandshow_Click()
PopUpValue6 = 0

If Check1_select.value = 0 Then
 
 Commandshow.Enabled = False
'********sub for alpha wise and Partywise according to new fast mathed
 DoEvents
 
 ''CON.Execute "Delete  from subledgertrail where " & stringyear
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 DoEvents
 ALPHAB
 'CON.CommitTrans
 Commandshow.Enabled = True
 
 Else
 
 PrintOption.Show
 
 End If

End Sub
Private Sub date1_KeyPress(KeyAscii As Integer)
 
If KeyAscii = 13 Then
    date2.SetFocus
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
    sendkeys "{TAB}"
End If

End Sub

Private Sub date2_LostFocus()
    If Trim(date2.text) <> "" Then
        If Not checkdate(Trim(date2.text), date2) Then
            date2.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 '   SendKeys "{TAB}"
End If

End Sub
Private Sub Form_Load()
con.Execute "delete  from treport where " & stringyear
con.Execute "Delete  from subledgertrail where " & stringyear

'Set CON_next = New ADODB.Connection
'next_dbase = "database=" & next_dbase
'If LCase(server_) = "server" Then
'   CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew & "; " & next_dbase & "; UID=" & sql_user & "; PWD=" & sql_pass
'Else
'   CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName & "; " & next_dbase & ";UID=; PWD=" & sql_pass
'End If
'
'
'CON_next.CursorLocation = adUseClient
'CON_next.Open



On Error GoTo ac1

Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> (Trim(UCase("MainMenu") Or Trim(UCase("frmbilllist"))))
    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
        Unload VB.Screen.ActiveForm
    End If
Loop

ac1:


Me.top = 0
Me.Left = 0
Set RS = New ADODB.Recordset
    
    RS.Open "select * from gledger where " & stringyear & " and slf=1", con, adOpenStatic, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        Do While Not RS.EOF
            COMBOGENLEDGER.AddItem Trim(RS!gledger)
            If Not RS.EOF Then
                RS.MoveNext
            End If
        Loop
    End If
    RS.close
    RS.Open "select * from setup1 where " & stringyear, con, adOpenStatic, adLockReadOnly
    CNSetup
    date1.text = RS!yarfrom
    date2.text = RS!yarto
    RS.close
End Sub
Sub xx()
End Sub
Sub OPENINGSUBLEDGERS()

If Trim(Alpha.text) <> "" Then
        'CON.Execute "Insert into subledgertrail  SELECT SUBLEDGER , YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId FROM SLEDGER where  gledger='" + Trim(COMBOGENLEDGER.Text) + "' AND subledger like '" + Trim(alpha.Text) + "%'", p, adCmdText
        'subledger opening start
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER ,YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER  where  SLEDGER.setupid=" & setupid & " and SLEDGER.fyear='" & session & "' and gledger='" + Trim(COMBOGENLEDGER.text) + "'  AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'"
    ' from invoice a
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) )" _
        & " where   INVOICEA.setupid=" & setupid & " and INVOICEA.fyear='" & session & "' and gledger='" + Trim(COMBOGENLEDGER.text) + "'and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER "
   ' from casha
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT,  " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER) ) " _
        & " where   cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and gledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'  and convert(smalldatetime,INVOICEDATE,103)< convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; "
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT ," & UId & " as UserId," & setupid & ",'" & session & "'  " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER))" _
        & " where   cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and gledger ='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY'   and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; "
        
    ' from credita
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER)) " _
        & " where  CREDITA.setupid=" & setupid & " and CREDITA.fyear='" & session & "' and gledger='" + Trim(COMBOGENLEDGER.text) + "' and INVOICEDATE < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; "
        
        
   ' from vouchers
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT ,  " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE VOUCHERS.setupid=" & setupid & " and VOUCHERS.fyear='" & session & "' and DEBITORCREDIT = 'D' and gledger ='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VOUCHERDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)  AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;"
        
        
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE VOUCHERS.setupid=" & setupid & " and VOUCHERS.fyear='" & session & "' and DEBITORCREDIT = 'C' and gledger ='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VOUCHERDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; "
         ''''ok
  'from cnf1a
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE CNF1A.setupid=" & setupid & " and CNF1A.fyear='" & session & "' and (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        ' and ReflectInAcc=0
                
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE CNF1A.setupid=" & setupid & " and CNF1A.fyear='" & session & "' and (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        'and ReflectInAcc=0
        
  ' from dnfa
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE DNFA.setupid=" & setupid & " and DNFA.fyear='" & session & "' and ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE DNFA.setupid=" & setupid & " and DNFA.fyear='" & session & "' and (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
   ' from cnf1b
        con.Execute "Insert into subledgertrail     SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1B.setupid=" & setupid & " and CNF1B.fyear='" & session & "' and (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail    SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1B.setupid=" & setupid & " and CNF1B.fyear='" & session & "' and (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.SUBLEDGER like '" + Trim(Alpha.text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
   ' dnfb
        con.Execute "Insert into subledgertrail     SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE dnFB.setupid=" & setupid & " and dnfb.fyear='" & session & "' and (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.text) + "'And convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)  AND SLEDGER.SUBLEDGER like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail   SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "'  " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE dnFB.setupid=" & setupid & " and dnfb.fyear='" & session & "' and (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.SUBLEDGER like '" + Trim(Alpha.text) + "%'" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
        
     
  End If
          
          



If Trim(Alpha.text) = "" And Combosubledger.text = "" Then
         
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , YEAROPENING,  0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER where  " & stringyear & " and gledger='" + Trim(COMBOGENLEDGER.text) + "'", p, adCmdText

        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER))  " _
        & " where invoiceA.setupid=" & setupid & " and invoiceA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "'and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER)) " _
        & " where  cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and gledger='" + Trim(COMBOGENLEDGER.text) + "' and  SLEDGER.SUBLEDGER <>'CASH PARTY' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER))" _
        & " where cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER <>'CASH PARTY' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER)) " _
        & " where  credita.setupid=" & setupid & " and credita.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
        
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE vouchers.setupid=" & setupid & " and vouchers.fyear='" & session & "' and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VOUCHERDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE vouchers.setupid=" & setupid & " and vouchers.fyear='" & session & "' and DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,VOUCHERDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
         ''''ok
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE cnf1a.setupid=" & setupid & " and cnf1a.fyear='" & session & "' and (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        ' and ReflectInAcc=0
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE cnf1a.setupid=" & setupid & " and cnf1a.fyear='" & session & "' and (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        ' and ReflectInAcc=0
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE dnfa.setupid=" & setupid & " and dnfa.fyear='" & session & "' and ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE dnfa.setupid=" & setupid & " and dnfa.fyear='" & session & "' and (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE  CNF1B.setupid=" & setupid & " and CNF1B.fyear='" & session & "' and (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1B.setupid=" & setupid & " and CNF1B.fyear='" & session & "' and (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE dnfB.setupid=" & setupid & " and dnfB.fyear='" & session & "' and (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE dnfB.setupid=" & setupid & " and dnfB.fyear='" & session & "' and (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.text) + "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
  End If
  

If Trim(Alpha.text) = "" And Combosubledger.text <> "" Then

        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER , SLEDGER.YEAROPENING, 0 AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & " as setupid,'" & session & "' as fyear  FROM SLEDGER  where " & stringyear & "  and  gledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER = '" & Combosubledger.text & "'", p, adCmdText
        
        con.Execute "Insert into subledgertrail  SELECT SLEDGER.SUBLEDGER , 0 as YEAROPENING,  sum(INVOICEA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN INVOICEA ON (SLEDGER.SUBLEDGER = INVOICEA.SUBLEDGER) )" _
        & " where  INVOICEA.setupid=" & setupid & " and INVOICEA.fyear='" & session & "' and gledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER=  '" & Combosubledger.text & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) AND SLEDGER.subledger like '" + Trim(Alpha.text) + "%' " _
        & " GROUP BY SLEDGER.SUBLEDGER "
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING, Sum(CASHA.NETAMOUNT) AS OPAMOUNTDEBIT,0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER)) " _
        & " where   cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and gledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER=  '" & Combosubledger.text & "'    and convert(smalldatetime,INVOICEDATE,103)< convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.NETAMOUNT) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING , 0  AS OPAMOUNTDEBIT, Sum(CASHA.BAA) AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM (SLEDGER LEFT JOIN CASHA ON (SLEDGER.SUBLEDGER = CASHA.SUBLEDGER))" _
        & " where   cashA.setupid=" & setupid & " and cashA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER ='" & Combosubledger.text & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CASHA.baa) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,    0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,SUM(CREDITA.NETAMOUNT)  AS OPAMOUNTCREDIT , " & UId & " as UserId," & setupid & ",'" & session & "' " _
        & " FROM (SLEDGER LEFT JOIN CREDITA ON (SLEDGER.SUBLEDGER = CREDITA.SUBLEDGER)) " _
        & " where  CREDITA.setupid=" & setupid & " and CREDITA.fyear='" & session & "' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER='" & Combosubledger.text & "' and convert(smalldatetime,INVOICEDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CREDITA.NETAMOUNT) <> 0; ", p, adCmdText
        
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,   0 AS YEAROPENING,Sum(VOUCHERS.Amount) AS OPAMOUNTDEBIT ,   0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE VOUCHERS.setupid=" & setupid & " and VOUCHERS.fyear='" & session & "' and DEBITORCREDIT = 'D' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER= '" & Combosubledger.text & "' and convert(smalldatetime,VOUCHERDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, 0 AS OpAmountdebit,  Sum(VOUCHERS.Amount) AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN VOUCHERS ON (SLEDGER.gledger = VOUCHERS.GenLedger) AND (SLEDGER.SUBLEDGER = VOUCHERS.SubLedger) " _
        & " WHERE DEBITORCREDIT = 'C' and genledger='" + Trim(COMBOGENLEDGER.text) + "' and SLEDGER.SUBLEDGER= '" & Combosubledger.text & "' and convert(smalldatetime,VOUCHERDATE,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(VOUCHERS.AMOUNT) <> 0; ", p, adCmdText
         ''''ok
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER,  0 AS YEAROPENING, Sum(CNF1A.NA) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT  , " & UId & " as UserId," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE cnf1a.setupid=" & setupid & " and cnf1a.fyear='" & session & "' and (((CNF1A.DC)='D')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "'and CNF1A.PSLD = '" & Combosubledger.text & "'  and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        ' and ReflectInAcc=0
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(CNF1A.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1A ON (SLEDGER.gledger = CNF1A.PGLD) AND (SLEDGER.SUBLEDGER = CNF1A.PSLD) " _
        & " WHERE cnf1a.setupid=" & setupid & " and cnf1a.fyear='" & session & "' and (((CNF1A.DC)='C')) and pgld = '" + Trim(COMBOGENLEDGER.text) + "' and CNF1A.PSLD = '" & Combosubledger.text & "'  and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)" _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1A.NA) <> 0; ", p, adCmdText
        
        ' and ReflectInAcc=0
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFA.NA)  AS OPAMOUNTDEBIT,  0  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFA ON (SLEDGER.gledger = DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD) " _
        & " WHERE dnfa.setupid=" & setupid & " and dnfa.fyear='" & session & "' and ((( DNFA.DC) = 'D' )) and  pgld = '" + Trim(COMBOGENLEDGER.text) + "' and   DNFA.PSLD = '" & Combosubledger.text & "' and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0  AS OPAMOUNTDEBIT,  Sum(DNFA.NA)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFA  ON (SLEDGER.gledger =DNFA.PGLD) AND (SLEDGER.SUBLEDGER = DNFA.PSLD)" _
        & " WHERE dnfa.setupid=" & setupid & " and dnfa.fyear='" & session & "' and (((DNFA.DC)='C')) and   pgld = '" + Trim(COMBOGENLEDGER.text) + "'  and  DNFA.PSLD = '" & Combosubledger.text & "'   and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFA.NA) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(CNF1B.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1B.setupid=" & setupid & " and CNF1B.fyear='" & session & "' and (((CNF1B.DC)='D')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "' and  CNF1B.SLD = '" & Combosubledger.text & "'    and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0; ", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(CNF1B.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN CNF1B ON (SLEDGER.gledger = CNF1B.GLD) AND (SLEDGER.SUBLEDGER = CNF1B.SLD) " _
        & " WHERE CNF1B.setupid=" & setupid & " and CNF1B.fyear='" & session & "'and (((CNF1B.DC)= 'C')) and gld  = '" + Trim(COMBOGENLEDGER.text) + "'  and  CNF1B.SLD = '" & Combosubledger.text & "'      and convert(smalldatetime,cnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103)  " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(CNF1B.A) <> 0;", p, adCmdText
        
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, Sum(DNFB.A) AS OPAMOUNTDEBIT, 0 AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER = DNFB.SLD) " _
        & " WHERE dnfb.setupid=" & setupid & " and dnfb.fyear='" & session & "' and (((DNFB.DC)='D')) and   gld = '" + Trim(COMBOGENLEDGER.text) + "'  and  DNFB.SLD = '" & Combosubledger.text & "'     And convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0; ", p, adCmdText
        
        con.Execute "Insert into subledgertrail SELECT SLEDGER.SUBLEDGER, 0 AS YEAROPENING, 0 AS OPAMOUNTDEBIT, Sum(DNFB.A)  AS OPAMOUNTCREDIT , " & UId & " as UserId ," & setupid & ",'" & session & "'" _
        & " FROM SLEDGER LEFT JOIN DNFB  ON (SLEDGER.gledger = DNFB.GLD) AND (SLEDGER.SUBLEDGER =DNFB.SLD)" _
        & " WHERE dnfb.setupid=" & setupid & " and dnfb.fyear='" & session & "' and (((DNFB.DC)= 'C')) and  gld = '" + Trim(COMBOGENLEDGER.text) + "'  and  DNFB.SLD = '" & Combosubledger.text & "'    and convert(smalldatetime,dnd,103) < convert(smalldatetime,'" + Trim(date1.text) + "',103) " _
        & " GROUP BY SLEDGER.SUBLEDGER " _
        & " HAVING  Sum(DNFB.A) <> 0 ", p, adCmdText
  End If
  
  
  
  
  
  
  
  
  
End Sub






