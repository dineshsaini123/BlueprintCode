VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form CASHBOOK 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4215
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "cashbook.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Check Cash Book"
      Height          =   405
      Left            =   1980
      TabIndex        =   8
      Top             =   3060
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton Commandreturn 
      Caption         =   "&Return"
      Height          =   405
      Left            =   5370
      TabIndex        =   5
      Top             =   3060
      Width           =   1545
   End
   Begin VB.CommandButton Commandshow 
      Caption         =   "&Show"
      Height          =   405
      Left            =   3690
      TabIndex        =   3
      Top             =   3060
      Width           =   1545
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      Height          =   315
      Left            =   3030
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1140
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
      Left            =   3030
      TabIndex        =   1
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
      Left            =   5580
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
   Begin VB.Label Label3 
      Caption         =   " - To - "
      Height          =   315
      Left            =   4680
      TabIndex        =   7
      Top             =   1680
      Width           =   585
   End
   Begin VB.Label Label2 
      Caption         =   "From The Date"
      Height          =   315
      Left            =   840
      TabIndex        =   6
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Gen. Ledger Desc."
      Height          =   285
      Left            =   870
      TabIndex        =   4
      Top             =   1170
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

Private Sub COMBOGENLEDGER_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   SendKeys "{tab}"

End If


End Sub

Private Sub Combosubledger_GotFocus()
    If Trim(COMBOGENLEDGER.Text) = "" Then
        COMBOGENLEDGER.SetFocus
    End If
End Sub

Private Sub Combosubledger_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"

End If
End Sub

Private Sub Combosubledger_LostFocus()
If Trim(Combosubledger.Text) <> "" Then
    If Trim(COMBOGENLEDGER.Text) <> "" Then
        If RS.State = 1 Then
            RS.Close
        End If
        RS.Open "select * from cbmf where " & stringyear & " and gld='" + Trim(COMBOGENLEDGER.Text) + "' and sld='" + Trim(Combosubledger.Text) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
        If RS.BOF Then
            Combosubledger.SetFocus
        End If
        RS.Close
    Else
        Combosubledger.Text = ""
    End If
    
End If
End Sub

Private Sub Command1_Click()
CheckCash = True
Commandshow_Click
End Sub

Private Sub Commandreturn_Click()
    Unload Me
    
End Sub

Private Sub Commandshow_Click()

If Trim(COMBOGENLEDGER.Text) <> "" Then
    If DateDiff("d", Trim(date1.Text), Trim(date2.Text)) <= 0 Then
        MsgBox "invalid date"
        Exit Sub
    End If
    Dim vflag As Boolean
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim rs4 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Set rs3 = New ADODB.Recordset
    Dim Balance As Double
    Balance = 0
    RS.Open "select * from gledger where " & stringyear & " and  gledger='" + Trim(COMBOGENLEDGER.Text) + "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    Balance = myround(RS!YEAROPENING, 2)
    RS.Close
    RS.Open "select sum(amount) from vouchers where  " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DebitorCredit='D'"
    If RS(0) >= 0 Then
        Balance = Balance + RS(0)
    End If
    RS.Close
    RS.Open "select sum(amount) from vouchers where  " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103) and DebitorCredit='C'"
    If RS(0) >= 0 Then
        Balance = Balance - RS(0)
    End If
    RS.Close
    ' Opening balance from cash counter sale
    If UCase(Genledger) = "CASH-IN-HAND" Then
    RS.Open "select sum(baa) from CASHa where convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)"
    If RS(0) >= 0 Then
        Balance = Balance + RS(0)
    End If
    RS.Close
    End If
    'Opening balance from cash counter sale end
    RS.Open "select sum(AMOUNT) from cashC where  " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,invoicedate,103)<convert(smalldatetime,'" + Trim(date1.Text) + "',103)"
    If RS(0) >= 0 Then
        Balance = Balance + RS(0)
    End If
    RS.Close
    Set rs1 = New ADODB.Recordset
    CON.Execute ("DELETE from treport where " & stringyear)
    rs1.Open "select * from treport where  " & stringyear & "", CON, adOpenStatic, adLockOptimistic, adCmdText
    rs1.addNew
    rs1!Text = "** Opening Balance as on " + Trim(date1.Text)
    rs1!ad = Balance
    rs1!Period = Trim(date1.Text) + "  To  " + Trim(date2.Text)
    rs1!header = "CASH BOOK"
    rs1!dorc = "D"
    rs1!UserId = main.UId
    rs1!FYear = main.session
    rs1!setupid = main.setupid
    rs1.Update
    RS.Open "SELECT * from VOUCHERS WHERE  " & stringyear & " and VOUCHERNUMBER IN (SELECT DISTINCT VOUCHERNUMBER  FROM VOUCHERS where  genledger <> '" + Trim(COMBOGENLEDGER.Text) + "') and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and VoucherType ='R'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    rs2.Open "select * from CASHC where  " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "'  and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenKeyset, adLockReadOnly, adCmdText
    rs3.Open "select * from casha where  " & stringyear & " and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenKeyset, adLockReadOnly, adCmdText
    vflag = True
    Set rs4 = New ADODB.Recordset
    rs4.Open "SELECT * from VOUCHERS WHERE   " & stringyear & " and genledger = '" + Trim(COMBOGENLEDGER.Text) + "'  and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and vouchertype ='J' ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    
voucheroption:
    If Not RS.BOF Then
        Do While Not RS.EOF
            'If rs!VoucherDate <= rs2!INVOICEDATE Then
                rs1.addNew
                rs1!Genledger = RS!Genledger & ""
                rs1!SUBLEDGER = RS!SUBLEDGER & ""
                rs1!vdate = RS!VoucherDate & ""
                rs1!vtype = RS!VoucherType & ""
                rs1!vno = Trim(str(RS!VOUCHERNUMBER))
                If Trim(RS!DESCRIPTION) <> "" Then
                    rs1!narration = Trim(RS!DESCRIPTION)
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
                If Trim(RS!cbnd) <> "" Then
                    rs1!cbno = Trim(RS!cbnd)
                Else
                    rs1!cbno = " "
                End If
                
                If Not RS.EOF Then
                    RS.MoveNext
                End If
                rs1!UserId = main.UId
    rs1!FYear = main.session
    rs1!setupid = main.setupid
                
                rs1.Update
        Loop
    End If
    If RS.State = 1 Then RS.Close
    If vflag = True Then
        RS.Open "SELECT * from VOUCHERS WHERE " & stringyear & " and VOUCHERNUMBER IN (SELECT DISTINCT VOUCHERNUMBER  FROM VOUCHERS where  genledger <> '" + Trim(COMBOGENLEDGER.Text) + "') and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and  VoucherType ='P'", CON, adOpenKeyset, adLockReadOnly, adCmdText
        If RS.RecordCount > 0 Then
            vflag = False
            GoTo voucheroption
        End If
    End If
    If Not rs4.BOF Then
        Do While Not rs4.EOF
            'If rs!VoucherDate <= rs2!INVOICEDATE Then
                rs1.addNew
                rs1!Genledger = rs4!Genledger & ""
                rs1!SUBLEDGER = rs4!SUBLEDGER & ""
                rs1!vdate = rs4!VoucherDate & ""
                rs1!vtype = rs4!VoucherType & ""
                rs1!vno = Trim(str(rs4!VOUCHERNUMBER))
                If Trim(rs4!DESCRIPTION) <> "" Then
                    rs1!narration = Trim(rs4!DESCRIPTION)
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
                If Trim(rs4!cbnd) <> "" Then
                    rs1!cbno = Trim(rs4!cbnd)
                Else
                    rs1!cbno = " "
                End If
                
                If Not rs4.EOF Then
                    rs4.MoveNext
                End If
                    rs1!UserId = main.UId
    rs1!FYear = main.session
    rs1!setupid = main.setupid
    rs1.Update
        Loop
    End If
    
    
    
    '**************** FOR  COUNTER SALE cashc
    If Not rs2.EOF Then
        Do While Not rs2.EOF
            rs1.addNew
           If rs2!amount > 0 Then
            rs1!Genledger = rs2!Genledger & ""
            rs1!SUBLEDGER = rs2!SUBLEDGER & ""
            rs1!vdate = rs2!InvoiceDate & ""
            rs1!vtype = Trim("S")
            rs1!vno = Trim(rs2!INVOICENO)
            rs1!aC = rs2!amount
            rs1!dorc = "C"
             rs1!narration = "Export Sales"
            Balance = Balance - rs2!amount
            rs1!Balance = Balance
                rs1!createdby = main.username
    rs1!createdon = Now
    rs1!FYear = main.session
    rs1!setupid = main.setupid
            rs1.Update
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
            rs1.addNew
            
            rs1!Genledger = "CASH-IN-HAND"
            rs1!SUBLEDGER = rs3!CASHPARTYNAME & ""
            rs1!vdate = rs3!InvoiceDate & ""
            rs1!vtype = Trim("S")
            rs1!vno = Trim(rs3!INVOICENO)
            rs1!aC = rs3!baa
            rs1!narration = "Export Sales"
            rs1!dorc = "C"
            Balance = Balance + rs3!baa
            rs1!Balance = Balance
                rs1!createdby = main.username
    rs1!createdon = Now
    rs1!FYear = main.session
    rs1!setupid = main.setupid
    rs1.Update
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

 ' ROHAN
 viewcash.GenrepNoFooter
 PrintOption.Command1.Enabled = False
 PrintOption.Show
    
    
Else
    MsgBox "gen. ledger or sub. ledger not selected"
End If
End Sub

Private Sub date1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{tab}"

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
    SendKeys "{tab}"
End If
End Sub

Private Sub date2_LostFocus()
    
    If Trim(date2.Text) <> "" Then
        If Not checkdate(Trim(date2.Text), date2) Then
            date2.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
Do While Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu"))
    If Trim(UCase(VB.Screen.ActiveForm.Name)) <> Trim(UCase("MainMenu")) Then
        Unload VB.Screen.ActiveForm
    End If
Loop
Me.TOP = 0
Me.Left = 0
CheckCash = False
'Set CON = New ADODB.Connection
Set RS = New ADODB.Recordset
   ' With CON
  '      .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\data.mdb"
  '      .Open
   ' End With
    RS.Open "select * from gledger where " & stringyear & " and cashbankbook=1", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        Do While Not RS.EOF
            COMBOGENLEDGER.AddItem Trim(RS!gledger)
            If Not RS.EOF Then
                   RS.MoveNext
            End If
        Loop
    End If
    
    RS.Close
    RS.Open "Select * from setup where " & stringyear & "", CON, adOpenKeyset, adLockReadOnly, adCmdText
    CNSetup
    date1.Text = RS!yarfrom
    date2.Text = RS!yarto
    RS.Close
End Sub
Sub TRYY()
Set rs1 = New ADODB.Recordset
    CON.Execute ("DELETE from treport where " & stringyear & "")
    rs1.Open "select * from treport where " & stringyear & "", CON, adOpenKeyset, adLockOptimistic, adCmdText
    rs1.addNew
    rs1!Text = "** Opening Balance as on " + Trim(date1.Text)
    rs1!ad = Balance
    rs1!Period = Trim(date1.Text) + "  To  " + Trim(date2.Text)
    rs1!header = "CASH BOOK"
  ' rs1!SUBLEDGER = Trim(Combosubledger.Text)
    rs1!dorc = "D"
    rs1.Update
    RS.Open "select * from vouchers where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "' and convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenKeyset, adLockReadOnly, adCmdText
    rs2.Open "select * from CASHC where " & stringyear & " and genledger='" + Trim(COMBOGENLEDGER.Text) + "'  and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenKeyset, adLockReadOnly, adCmdText
    
    ' rs2.Open "select * from invoicea where genledger='" + Trim(COMBOGENLEDGER.TEXT) + "' and subledger='" + Trim(Combosubledger.TEXT) + "' and convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" + Trim(date1.Text) + "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        Do While Not RS.EOF 'And Not rs2.EOF
            'If rs!VoucherDate <= rs2!INVOICEDATE Then
                rs1.addNew
                rs1!Genledger = RS!Genledger
                rs1!SUBLEDGER = RS!SUBLEDGER
                rs1!vdate = RS!VoucherDate
                rs1!vtype = RS!VoucherType
                rs1!vno = Trim(str(RS!VOUCHERNUMBER))
                If Trim(RS!DESCRIPTION) <> "" Then
                    rs1!narration = Trim(RS!DESCRIPTION)
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
                If Trim(RS!cbnd) <> "" Then
                    rs1!cbno = Trim(RS!cbnd)
                Else
                    rs1!cbno = " "
                End If
                
                
                If Not RS.EOF Then
                    RS.MoveNext
                End If
            'Else
            '    rs1.AddNew
            '    rs1!Genledger = rs2!Genledger
            '    rs1!subledger = rs2!subledger
            '    rs1!vdate = rs2!INVOICEDATE
            '    rs1!vtype = Trim("I")
            '    rs1!vno = Trim(rs2!INVOICENO)
            '    rs1!ad = rs2!netamount
            '    rs1!dorc = "D"
            '    balance = balance + rs2!netamount
            '    rs1!balance = balance
            '    If Not rs2.EOF Then
            '        rs2.MoveNext
            '    End If
            'End If
                rs1!createdby = main.username
    rs1!createdon = Now
    rs1!FYear = main.session
    rs1!setupid = main.setupid
            
            rs1.Update
        Loop
    End If
    'If Not rs2.EOF Then
    '    Do While Not rs2.EOF
    '        rs1.AddNew
    '        rs1!Genledger = rs2!Genledger
    '        rs1!subledger = rs2!subledger
    '        rs1!vdate = rs2!INVOICEDATE
    '        rs1!vtype = Trim("I")
    '        rs1!vno = Trim(rs2!INVOICENO)
    '        rs1!ad = rs2!netamount
    '        rs1!dorc = "D"
    '        balance = balance + rs2!netamount
    '        rs1!balance = balance
    '        If Not rs2.EOF Then
    '           rs2.MoveNext
    '        End If
    '        rs1.Update
    '    Loop
    'End If
    'If Not rs.EOF Then
    '    Do While Not rs.EOF
    '            rs1.AddNew
    '            rs1!Genledger = rs!Genledger
    '            rs1!subledger = rs!subledger
    '            rs1!vdate = rs!VoucherDate
    '            rs1!vtype = rs!VoucherType
    '            rs1!vno = Trim(Str(rs!VoucherNumber))
    '            rs1!narration = Trim(rs!DESCRIPTION)
    '            If Trim(UCase(rs!DebitorCredit)) = Trim(UCase("D")) Then
    '                rs1!ad = rs!amount
    '                rs1!dorc = "D"
    '                balance = balance + rs!amount
    '                rs1!balance = balance
    '            Else
    '                rs1!ac = rs!amount
    '                rs1!dorc = "C"
    '                balance = balance - rs!amount
    '                rs1!balance = balance
    '            End If
    '            rs1!cbno = rs!CBND
    '            If Not rs.EOF Then
    '                rs.MoveNext
    '            End If
    '            rs1.Update
    '    Loop
    'End If
    
End Sub
