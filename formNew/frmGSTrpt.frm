VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGSTrpt 
   Caption         =   "GST Related Report..."
   ClientHeight    =   7716
   ClientLeft      =   60
   ClientTop       =   396
   ClientWidth     =   7656
   Icon            =   "frmGSTrpt.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7716
   ScaleWidth      =   7656
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboGST 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmGSTrpt.frx":000C
      Left            =   5670
      List            =   "frmGSTrpt.frx":0019
      TabIndex        =   15
      Text            =   "Select GST"
      Top             =   1125
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.TextBox txtAmt 
      Height          =   330
      Left            =   1080
      TabIndex        =   13
      Top             =   1530
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox cboReport 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmGSTrpt.frx":0028
      Left            =   1065
      List            =   "frmGSTrpt.frx":0032
      TabIndex        =   6
      Top             =   270
      Width           =   3525
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   585
      Left            =   2145
      TabIndex        =   5
      Top             =   5850
      Width           =   1020
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   585
      Left            =   4335
      TabIndex        =   4
      Top             =   5850
      Width           =   1005
   End
   Begin VB.ListBox cboState 
      Appearance      =   0  'Flat
      Height          =   2616
      Left            =   1065
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   2115
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Export"
      Height          =   585
      Left            =   3240
      TabIndex        =   2
      Top             =   5850
      Width           =   1020
   End
   Begin VB.ComboBox COMBOGENLEDGER 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmGSTrpt.frx":0056
      Left            =   1065
      List            =   "frmGSTrpt.frx":0060
      TabIndex        =   1
      Top             =   1125
      Visible         =   0   'False
      Width           =   4605
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   585
      Left            =   1080
      TabIndex        =   0
      Top             =   5850
      Width           =   1020
   End
   Begin Crystal.CrystalReport cr 
      Left            =   195
      Top             =   3690
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker toDate 
      Height          =   315
      Left            =   3225
      TabIndex        =   7
      Top             =   735
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   572
      _Version        =   393216
      Format          =   155582465
      CurrentDate     =   38845
   End
   Begin MSComCtl2.DTPicker FromDate 
      Height          =   300
      Left            =   1065
      TabIndex        =   8
      Top             =   735
      Width           =   1350
      _ExtentX        =   2371
      _ExtentY        =   529
      _Version        =   393216
      Format          =   155582465
      CurrentDate     =   38845
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      Height          =   285
      Left            =   135
      TabIndex        =   14
      Top             =   1575
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Type"
      Height          =   315
      Left            =   135
      TabIndex        =   12
      Top             =   330
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "From Date "
      Height          =   240
      Left            =   135
      TabIndex        =   11
      Top             =   795
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      Height          =   255
      Left            =   2505
      TabIndex        =   10
      Top             =   735
      Width           =   1245
   End
   Begin VB.Label lblLeder 
      BackStyle       =   0  'Transparent
      Caption         =   "Exp. Ledger"
      Height          =   315
      Left            =   135
      TabIndex        =   9
      Top             =   1185
      Visible         =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "frmGSTrpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboReport_Click()

cbostate.Visible = False
cmdview.Enabled = True
cmdPrint.Enabled = True
cmdExcel.Enabled = True

Label4.Visible = False
txtAmt.Visible = False
txtAmt.Visible = False
cboGST.Visible = False


If cboReport.Text = "State Wise Sale" Then
   cbostate.Visible = True
   COMBOGENLEDGER.Visible = False
   lblLeder.Visible = False
   
ElseIf cboReport.Text = "Expenses List" Then

   COMBOGENLEDGER.Visible = True
   lblLeder.Visible = True
   Label4.Visible = True
   txtAmt.Visible = True
   
   txtAmt.Visible = True
   cboGST.Visible = True

   
   COMBOGENLEDGER.Clear
   If RS.State = 1 Then RS.close
   RS.Open "select  gledger from GLEDGER where Category='Expences' order by gledger", con
   While RS.EOF = False
   COMBOGENLEDGER.AddItem RS(0)
   RS.MoveNext
   Wend


End If

End Sub
Sub expenceLedger()

If Trim(COMBOGENLEDGER.Text) <> "" Then
    If DateDiff("d", Trim(FromDate.value), Trim(toDate.value)) <= 0 Then
        MsgBox "invalid date"
        Exit Sub
    End If
    
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim rs4 As ADODB.Recordset
    Dim rs5 As ADODB.Recordset
    Dim rs6 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Set rs3 = New ADODB.Recordset
    Set rs4 = New ADODB.Recordset
    Set rs5 = New ADODB.Recordset
    Set rs6 = New ADODB.Recordset
    Dim Balance As Double
    Balance = 0
    Set rs1 = New ADODB.Recordset
    con.Execute "delete from treport"
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "Select * from treport", con, adOpenDynamic, adLockPessimistic

   
   '************* voucher start
   If RS.State = 1 Then RS.close
    RS.Open "select * from vouchers where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and VoucherDate>=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and VoucherDate<=convert(smalldatetime,'" + Trim(toDate.value) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
    If Not RS.BOF Then
        Do While Not RS.EOF
             rs1.AddNew
             rs1!Genledger = Trim(RS!Genledger)
             rs1!subledger = Trim(RS!subledger)
             rs1!vdate = RS!voucherDATE & ""
             rs1!vtype = Trim(RS!vouchertype)
             rs1!vno = Trim(Str(RS!VOUCHERNUMBER))
             If Len(Trim(RS!DESCRIPTION)) = 0 Then
                rs1!narration = " "
             Else
                rs1!narration = Trim(RS!DESCRIPTION)
             End If
             If Trim(UCase(RS!DebitorCredit)) = Trim(UCase("D")) Then
                rs1!ad = RS!amount
                rs1!dorc = "D"
             Else
                rs1!aC = RS!amount
                rs1!dorc = "C"
            End If
            If Len(Trim(RS!CBND)) = 0 Or IsNull(RS!CBND) Then
               rs1!cbno = " "
            Else
               rs1!cbno = RS!CBND & ""
            End If
            rs1!userid = UId
            rs1.update
            If Not RS.EOF Then
               RS.MoveNext
            End If
        Loop
    End If
'***************   voucher end
'sales GL start

 If UCase(Trim(COMBOGENLEDGER.Text)) = "SALES" Then
    rs3.Open "select * from invoiceA where INVOICEDATE>=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and INVOICEDATE<=convert(smalldatetime,'" + Trim(toDate.value) + "',103) order by invoiceno", con, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs3.EOF Then
           Do While Not rs3.EOF
              If rs3!gamount <> 0 Then
                 rs1.AddNew
                 rs1!Genledger = "SALES"
                 rs1!vdate = rs3!invoiceDate & ""
                 rs1!vtype = Trim("I")
                 rs1!vno = Trim(rs3!invoiceNo)
                 rs1!narration = "Sales Invoice"
                 balance1 = rs3!gamount
                 If rs2.State = 1 Then rs2.close
                 rs2.Open "select * from invoiceC where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and INVOICEDATE<=convert(smalldatetime,'" + Trim(toDate.value) + "',103) and invoiceno =" + Str(rs3!invoiceNo), con, adOpenStatic, adLockReadOnly, adCmdText
                 If rs2.RecordCount > 0 Then
                    Do While Not rs2.EOF
                       If Left(rs2!DebitorCredit, 1) = "D" Then
                          balance1 = balance1 - rs2!amount
                       Else
                          balance1 = balance1 + rs2!amount
                       End If
                       rs2.MoveNext
                    Loop
                 End If
                 rs1!aC = balance1
                 rs1!narration = "Sales Invoice"
                 rs1!dorc = "C"
             End If
             If Not rs3.EOF Then rs3.MoveNext
             rs1!userid = UId
             rs1.update
         Loop
    End If
    rs3.close
    
   'CASH COUNTER SALE
    rs3.Open "select * from CASHA where INVOICEDATE>=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and INVOICEDATE<=convert(smalldatetime,'" + Trim(toDate.value) + "',103) order by invoiceno", con, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs3.EOF Then
       Do While Not rs3.EOF
          If rs3!gamount <> 0 Then
             rs1.AddNew
             rs1!Genledger = "SALES"
             rs1!vdate = rs3!invoiceDate & ""
             rs1!vtype = Trim("S")
             rs1!vno = Trim(rs3!invoiceNo)
             rs1!narration = "Sales C/M"
             balance1 = rs3!gamount
             If rs2.State = 1 Then rs2.close
             rs2.Open "select * from CASHC where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and INVOICEDATE<=convert(smalldatetime,'" + Trim(toDate.value) + "',103) and invoiceno =" + Str(rs3!invoiceNo), con, adOpenStatic, adLockReadOnly, adCmdText
             If rs2.RecordCount > 0 Then
                Do While Not rs2.EOF
                   If Left(rs2!DebitorCredit, 1) = "D" Then
                      balance1 = balance1 - rs2!amount
                   Else
                      balance1 = balance1 + rs2!amount
                   End If
                   rs2.MoveNext
                Loop
             End If
             rs1!aC = balance1
             rs1!narration = "Sales C/M"
             rs1!dorc = "C"
          End If
          If Not rs3.EOF Then rs3.MoveNext
          rs1!userid = UId
          rs1.update
       Loop
    End If
    rs3.close
Else
    rs2.Open "select * from invoiceC where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and INVOICEDATE<=convert(smalldatetime,'" + Trim(toDate.value) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not rs2.EOF
       If rs2!amount <> 0 Then
          rs1.AddNew
          rs1!Genledger = rs2!Genledger & ""
          rs1!vdate = rs2!invoiceDate & ""
          rs1!vtype = Trim("I")
          rs1!narration = "Sales Invoice"
          rs1!vno = Trim(rs2!invoiceNo)
          If Left(rs2!DebitorCredit, 1) = "D" Then
             rs1!ad = rs2!amount
             rs1!dorc = "D"
          Else
            rs1!aC = rs2!amount
            rs1!dorc = "C"
          End If
      End If
      If Not rs2.EOF Then rs2.MoveNext
      rs1!userid = UId
      rs1.update
    Loop
    
    If rs2.State = 1 Then rs2.close
    'CASH COUNTER SALE
    rs2.Open "select * from CASHC where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and INVOICEDATE<=convert(smalldatetime,'" + Trim(toDate.value) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not rs2.EOF
       If rs2!amount <> 0 Then
          rs1.AddNew
          rs1!Genledger = rs2!Genledger & ""
          rs1!vdate = rs2!invoiceDate & ""
          rs1!vtype = Trim("S")
          rs1!narration = "Sales C/M"
          rs1!vno = Trim(rs2!invoiceNo)
          If Left(rs2!DebitorCredit, 1) = "D" Then
             rs1!ad = rs2!amount
             rs1!dorc = "D"
          Else
             rs1!aC = rs2!amount
             rs1!dorc = "C"
          End If
       End If
       If Not rs2.EOF Then rs2.MoveNext
       rs1!userid = UId
       rs1.update
    Loop
End If
'Sales RETURNS GL start
If UCase(Trim(COMBOGENLEDGER.Text)) = "SALES RETURN" Then
   rs4.Open "select * from CREDITA where INVOICEDATE>=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and INVOICEDATE<=convert(smalldatetime,'" + Trim(toDate.value) + "',103) order by invoiceno", con, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs4.BOF Then
      Do While Not rs4.EOF
         If rs4!gamount <> 0 Then
            rs1.AddNew
            rs1!Genledger = "SALES RETURN"
            rs1!vdate = rs4!invoiceDate & ""
            rs1!vtype = Trim("C")
            rs1!vno = Trim(rs4!invoiceNo)
            rs1!narration = "Credit Note(I)"
            balance1 = rs4!gamount
            If rs2.State = 1 Then rs2.close
                  rs2.Open "select * from CREDITC where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and INVOICEDATE<=convert(smalldatetime,'" + Trim(toDate.value) + "',103) and invoiceno =" + Str(rs4!invoiceNo), con, adOpenStatic, adLockReadOnly, adCmdText
                    If rs2.RecordCount > 0 Then
                       Do While Not rs2.EOF
                            If Left(rs2!DebitorCredit, 1) = "D" Then
                                balance1 = balance1 + rs2!amount
                            Else
                                balance1 = balance1 - rs2!amount
                            End If
                            rs2.MoveNext
                        Loop
                    End If
                    rs1!ad = balance1
                    rs1!dorc = "D"
   '                 BALANCE = BALANCE - balance1
    '                rs1!BALANCE = BALANCE
                End If
                If Not rs4.EOF Then
                   rs4.MoveNext
                End If
                rs1!userid = UId
                rs1.update
                
            Loop
        End If
        rs4.close
    Else
       If rs2.State = 1 Then rs2.close
       rs2.Open "select * from CREDITC where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and INVOICEDATE>=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and INVOICEDATE<=convert(smalldatetime,'" + Trim(toDate.value) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not rs2.EOF
            If rs2!amount <> 0 Then
                rs1.AddNew
                rs1!Genledger = rs2!Genledger & ""
                rs1!narration = "Credit Note(I)"
                rs1!vdate = rs2!invoiceDate & ""
                rs1!vtype = Trim("I")
                rs1!vno = Trim(rs2!invoiceNo)
                If Left(rs2!DebitorCredit, 1) = "D" Then
                    rs1!ad = rs2!amount
                    rs1!dorc = "D"
'                   BALANCE = BALANCE + rs2!amount
                Else
                    rs1!aC = rs2!amount
                    rs1!dorc = "C"
 '                  BALANCE = BALANCE - rs2!amount
                End If
            End If
            If Not rs2.EOF Then
               rs2.MoveNext
            End If
            rs1!userid = UId
            rs1.update
            
        Loop
    End If

'CREDIT NOTE
    rs3.Open "select * from Cnf1a where pgld ='" + Trim(COMBOGENLEDGER.Text) + "' and cnd >=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and cnd<=convert(smalldatetime,'" + Trim(toDate.value) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs3.BOF Then
        Do While Not rs3.EOF
            rs1.AddNew
            rs1!Genledger = rs3!Pgld & ""
            rs1!vdate = rs3!Cnd & ""
            rs1!vtype = "C"
            rs1!vno = Trim(Str(rs3!cnn))
            If Trim(UCase(rs3!dc)) = Trim(UCase("D")) Then
                rs1!ad = rs3!na
                rs1!dorc = "D"
            Else
                rs1!aC = rs3!na
                rs1!dorc = "C"
            End If
            rs1!narration = "Credit Note"
            rs1!userid = UId
            rs1.update
            
            If Not rs4.EOF Then
                rs3.MoveNext
            End If
        Loop
    End If
    'CREDIT NOTE B
   rs4.Open "select * from Cnf1b where gld ='" + Trim(COMBOGENLEDGER.Text) + "' and cnd >=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and cnd<=convert(smalldatetime,'" + Trim(toDate.value) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs4.BOF Then
        Do While Not rs4.EOF
            rs1.AddNew
            rs1!Genledger = Trim(rs4!gld)
            rs1!vdate = rs4!Cnd & ""
            rs1!vtype = "C"
            rs1!vno = Trim(Str(rs4!cnn))
            If Trim(UCase(rs4!dc)) = Trim(UCase("D")) Then
                rs1!ad = rs4!a
                rs1!dorc = "D"
            Else
                rs1!aC = rs4!a
                rs1!dorc = "C"
            End If
            rs1!narration = "Credit Note"
            rs1!userid = UId
            rs1.update
            
            If Not rs4.EOF Then
                rs4.MoveNext
            End If
        Loop
    End If
'credit note b end
    
    'debit NOTE
    rs5.Open "select * from dnfa where pgld ='" + Trim(COMBOGENLEDGER.Text) + "' and dnd >=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and dnd<=convert(smalldatetime,'" + Trim(toDate.value) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs5.BOF Then
        Do While Not rs5.EOF
            rs1.AddNew
            rs1!Genledger = rs5!Pgld & ""
            rs1!vdate = rs5!dnd & ""
            rs1!vtype = "C"
            rs1!vno = Trim(Str(rs5!dnn))
            If Trim(UCase(rs5!dc)) = Trim(UCase("D")) Then
                rs1!ad = rs5!na
                rs1!dorc = "D"
            Else
                rs1!aC = rs5!na
                rs1!dorc = "C"
            End If
            rs1!narration = "Debit Note"
            rs1!userid = UId
            rs1.update
            
            If Not rs5.EOF Then
                rs5.MoveNext
            End If
        Loop
    End If
    
    'Debit NOTE B
   rs6.Open "select * from dnfb where gld ='" + Trim(COMBOGENLEDGER.Text) + "' and dnd >=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and dnd<=convert(smalldatetime,'" + Trim(toDate.value) + "',103)", con, adOpenStatic, adLockReadOnly, adCmdText
    If Not rs6.BOF Then
        Do While Not rs6.EOF
            rs1.AddNew
            rs1!Genledger = rs6!gld & ""
            rs1!vdate = rs6!dnd & ""
            rs1!vtype = "C"
            rs1!vno = Trim(Str(rs6!dnn))
            If Trim(UCase(rs6!dc)) = Trim(UCase("D")) Then
                rs1!ad = rs6!a
                rs1!dorc = "D"
            Else
                rs1!aC = rs6!a
                rs1!dorc = "C"
            End If
            rs1!narration = "Debit Note"
            rs1!userid = UId
            rs1.update
            
            If Not rs6.EOF Then
                rs6.MoveNext
            End If
        Loop
    End If
'debit note b end

' making the balance in the output file

''Set rs1 = New ADODB.Recordset
''rs1.Open "select * from treport order by vdate, vtype,vno", con, adOpenStatic, adLockOptimistic, adCmdText
''    Balance = 0
''    If Not rs1.BOF Then
''        rs1.MoveFirst
''        Do While Not rs1.EOF
''            Balance = Balance + rs1!aD - rs1!ac
''            rs1!Balance = Balance
''            rs1.update
''            If Not rs1.EOF Then
''                rs1.MoveNext
''            End If
''        Loop
''    End If
''    If rs1.State = 1 Then
''        rs1.close
''    End If
''    If RS.State = 1 Then
''        RS.close
''    End If
    
Else
    MsgBox "gen. ledger not selected"
End If

''''===========================================================



End Sub
Private Sub cmdExcel_Click()

If cboReport.Text = "State Wise Sale" Then
    SaleStatewise
    cmdExcel.Enabled = False
    cmdview.Enabled = True
ElseIf cboReport.Text = "Expenses List" Then
    
    If COMBOGENLEDGER.Text <> "" Then
        Screen.MousePointer = vbHourglass
        expenceLedger
        ExcelExpenceLedger
        Screen.MousePointer = vbDefault
    End If
    
    
    
End If




End Sub
Sub SaleStatewise()
Dim row_, col_ As Integer
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim xlapp As Excel.Application

row_ = 0
col_ = 0

'==============================================================
s = ""
For I = 0 To cbostate.ListCount - 1
If cbostate.Selected(I) = False Then
   con.Execute "delete from tempLedger1 where party ='" & cbostate.List(I) & "'"
End If
Next
'==============================================================


If xlapp Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True

Dim k1 As Integer
Dim sale As Double
Dim saleR As Double
Dim saleN As Double

k1 = 1
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add


xlSheet.Range("a1", "d1").Merge
xlSheet.Range("a1").value = cname_1 & " " & cname_2
xlSheet.Range("a1").Font.Bold = True
xlSheet.Range("a1").Font.Size = 14


 
xlSheet.Range("b3").HorizontalAlignment = xlRight
xlSheet.Range("c3").HorizontalAlignment = xlRight
xlSheet.Range("d3").HorizontalAlignment = xlRight


k1 = k1 + 1

xlSheet.Cells(k1, 1) = FromDate.value
xlSheet.Cells(k1, 2) = "To"
xlSheet.Cells(k1, 3) = Str(toDate.value)

k1 = k1 + 2

xlSheet.Columns("A:F").ColumnWidth = 15

xlSheet.Range("B:D").NumberFormat = "#,##0.00"

xlSheet.Cells(k1, 1).value = "State Name"
xlSheet.Cells(k1, 1).Font.Bold = True
xlSheet.Cells(k1, 2).value = "        Sale Amt"
xlSheet.Cells(k1, 2).Font.Bold = True
xlSheet.Cells(k1, 3).value = "Sale Amt Ret."
xlSheet.Cells(k1, 3).Font.Bold = True
xlSheet.Cells(k1, 4).value = "                Net Sale"
xlSheet.Cells(k1, 4).Font.Bold = True



k1 = k1 + 1


sale = 0
saleR = 0
saleN = 0


If RS.State = 1 Then RS.close
RS.Open "select  Party,sum(Balance),sum(Cr)  from  tempLedger1 group by Party", con
For I = 1 To RS.RecordCount
     If RS.EOF = False Then
            
            's1_ = Str(Format(RS(1), "0.00"))
            's2_ = Str(Format(RS(2), "0.00"))
            's3_ = Str(Format(RS(1) - RS(2), "0.00"))
            
            xlSheet.Cells(k1, 1).value = RS(0)
            xlSheet.Cells(k1, 2).value = Round(RS(1), 2)
            xlSheet.Cells(k1, 3).value = Round(RS(2), 2)
            xlSheet.Cells(k1, 4).value = Round(RS(1) - RS(2), 2)
            
            sale = sale + RS(1)
            saleR = saleR + RS(2)
            saleN = saleN + (RS(1) - RS(2))

            
            k1 = k1 + 1
            RS.MoveNext
      End If
Next

k1 = k1 + 1
xlSheet.Cells(k1, 1).value = "TOTAL"
xlSheet.Cells(k1, 2).value = sale
xlSheet.Cells(k1, 3).value = saleR
xlSheet.Cells(k1, 4).value = saleN

xlSheet.Cells(k1, 1).Interior.Color = vbCyan
xlSheet.Cells(k1, 2).Interior.Color = vbCyan
xlSheet.Cells(k1, 3).Interior.Color = vbCyan
xlSheet.Cells(k1, 4).Interior.Color = vbCyan

End Sub
Sub ExcelExpenceLedger()

Dim row_, col_ As Integer
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim xlapp As Excel.Application


row_ = 0
col_ = 0

If xlapp Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True

Dim k1 As Integer
Dim sale As Double
Dim saleR As Double
Dim saleN As Double

Dim gst_ As Double
Dim a1, a2 As Double
Dim cgst, sgst As Double
Dim fatch As Double
Dim tgst As Double
Dim Netgst As Double



If cboGST.Text <> "" Then
   gst_ = Round(Val(cboGST.Text) / 2, 2)
End If




k1 = 1

Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add


xlSheet.Range("a1", "d1").Merge
xlSheet.Range("a1").value = cname_1 & " " & cname_2
xlSheet.Range("a1").Font.Bold = True
xlSheet.Range("a1").Font.Size = 14


xlSheet.Range("a2", "b2").Merge
xlSheet.Range("a2").value = COMBOGENLEDGER.Text
xlSheet.Range("a2").Font.Bold = True
xlSheet.Range("a2").Font.Size = 12

 
xlSheet.Range("a3").HorizontalAlignment = xlRight
xlSheet.Range("b3").HorizontalAlignment = xlCenter
xlSheet.Range("c3").HorizontalAlignment = xlRight
xlSheet.Range("d3").HorizontalAlignment = xlRight
xlSheet.Range("e3").HorizontalAlignment = xlRight
xlSheet.Range("f3").HorizontalAlignment = xlRight
xlSheet.Range("g3").HorizontalAlignment = xlRight
xlSheet.Range("h3").HorizontalAlignment = xlRight

k1 = k1 + 2

xlSheet.Columns("A").ColumnWidth = 12
xlSheet.Columns("B").ColumnWidth = 45
xlSheet.Columns("C:h").ColumnWidth = 12

xlSheet.Cells(k1, 1).Interior.Color = &H80C0FF
xlSheet.Cells(k1, 2).Interior.Color = &H80C0FF
xlSheet.Cells(k1, 3).Interior.Color = &H80C0FF
xlSheet.Cells(k1, 4).Interior.Color = &H80C0FF
xlSheet.Cells(k1, 5).Interior.Color = &H80C0FF
xlSheet.Cells(k1, 6).Interior.Color = &H80C0FF
xlSheet.Cells(k1, 7).Interior.Color = &H80C0FF
xlSheet.Cells(k1, 8).Interior.Color = &H80C0FF

xlSheet.Cells(k1, 1).value = "Dale"
xlSheet.Cells(k1, 1).Font.Bold = True
xlSheet.Cells(k1, 2).value = "Narration"
xlSheet.Cells(k1, 2).Font.Bold = True
xlSheet.Cells(k1, 3).value = "Amount(Dr)"
xlSheet.Cells(k1, 3).Font.Bold = True

xlSheet.Cells(k1, 4).value = "Amount(Cr)"
xlSheet.Cells(k1, 4).Font.Bold = True
xlSheet.Cells(k1, 4).EntireColumn.Hidden = False

xlSheet.Cells(k1, 5).value = "CGST(" & gst_ & " %)"
xlSheet.Cells(k1, 5).Font.Bold = True

xlSheet.Cells(k1, 6).value = "SGST(" & gst_ & " %)"
xlSheet.Cells(k1, 6).Font.Bold = True

xlSheet.Cells(k1, 7).value = "IGST(" & "0.00  %" & ")"
xlSheet.Cells(k1, 7).Font.Bold = True

xlSheet.Cells(k1, 8).value = "Total GST"
xlSheet.Cells(k1, 8).Font.Bold = True




k1 = k1 + 1

sale = 0
saleR = 0
saleN = 0
tgst = 0
Netgst = 0



If RS.State = 1 Then RS.close
RS.Open "select  vdate,narration,ad ,ac from treport order by vdate", con
For I = 1 To RS.RecordCount
     If RS.EOF = False Then
            If Not IsNull(RS(0)) Then
            
            
                   If txtAmt.Text = "" Then
                        fatch = True
                    Else
                       If (IIf(IsNull(RS(2)), 0, RS(2)) >= Val(txtAmt) Or IIf(IsNull(RS(3)), 0, RS(3)) >= Val(txtAmt)) Then
                           fatch = True
                        Else
                           fatch = False
                       End If
                    End If
            
            
            
                     If fatch = True Then
                     
                        xlSheet.Cells(k1, 1).value = Format(CStr(RS(0)), "mm/dd/yyyy")
                        xlSheet.Cells(k1, 2).value = RS(1)
                        xlSheet.Cells(k1, 3).value = RS(2) & ""
                        xlSheet.Cells(k1, 4).value = RS(3) & ""
                        
                        If Not IsNull(RS(2)) Then
                        sale = sale + RS(2)
                        End If
                        saleN = saleN + IIf(IsNull(RS(3)), 0, RS(3))
                        
                        k1 = k1 + 1
                        
    
                                If a2 > 0 Then
                                
                                 a1 = RS(2)
                                 a2 = (a1 * gst_) / 100

                                   
                                   '-----CGST
                                   a1 = Round(a2 - Int(a2), 2)
                                   If a1 <= 0.49 Then
                                       xlSheet.Cells(k1, 5).value = Int(a2)
                                       cgst = cgst + Int(a2)
                                   Else
                                      xlSheet.Cells(k1, 5).value = Int(a2) + 1
                                      cgst = cgst + Int(a2) + 1
                                   End If
                                   
                                   
                                   '-----SGST
                                   a1 = Round(a2 - Int(a2), 2)
                                   If a1 <= 0.49 Then
                                       xlSheet.Cells(k1, 6).value = Int(a2)
                                       sgst = sgst + Int(a2)
                                   Else
                                      xlSheet.Cells(k1, 6).value = Int(a2) + 1
                                      sgst = sgst + Int(a2) + 1
                                   End If
                                   xlSheet.Cells(k1, 7).value = 0
                                   xlSheet.Cells(k1, 8).value = (xlSheet.Cells(k1, 5).value + xlSheet.Cells(k1, 6).value + xlSheet.Cells(k1, 7).value)
                                   
                                   Netgst = Netgst + Val(xlSheet.Cells(k1, 8).value)
                                  
                                  
                                  
                                
                                
                                
                   
                            End If
            
            
            
            
            End If
            
            End If
            RS.MoveNext
      End If
Next

k1 = k1 + 1
xlSheet.Cells(k1, 1).value = "TOTAL"
xlSheet.Cells(k1, 3).value = sale
xlSheet.Cells(k1, 4).value = saleN

xlSheet.Cells(k1, 5).value = cgst
xlSheet.Cells(k1, 6).value = sgst
xlSheet.Cells(k1, 7).value = 0
xlSheet.Cells(k1, 8).value = Netgst

xlSheet.Cells(k1, 1).Interior.Color = vbCyan
xlSheet.Cells(k1, 2).Interior.Color = vbCyan
xlSheet.Cells(k1, 3).Interior.Color = vbCyan
xlSheet.Cells(k1, 4).Interior.Color = vbCyan
xlSheet.Cells(k1, 5).Interior.Color = vbCyan
xlSheet.Cells(k1, 6).Interior.Color = vbCyan
xlSheet.Cells(k1, 7).Interior.Color = vbCyan
xlSheet.Cells(k1, 8).Interior.Color = vbCyan



End Sub

Private Sub cmdexit_Click()
 Unload Me
End Sub
Sub genLedgerTrial()

    
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
    Dim wheredate_vc As String
    Dim wheredate_inv As String
    
    
    viewsubledger = False
    Balance = 0
    OPBALANCE = 0
    If RS.State = 1 Then
        RS.close
    End If
    
    wheredate_vc = "convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" & Trim(FromDate.value) & "',103) and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" & Trim(toDate.value) & "',103)"
    wheredate_inv = "convert(smalldatetime,invoicedate,103)>=convert(smalldatetime,'" & Trim(FromDate.value) & "',103) and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" & Trim(toDate.value) & "',103)"
    
    
    con.Execute "Delete from TemprptTrialBalance where  " & stringyear & ""
 
    con.Execute "INSERT INTO TemprptTrialBalance (Gledger, OpeningBalance,userid,fyear,setupid)  SELECT Gledger.gledger, Gledger.YEAROPENING," & UId & " as Userid,'" & main.session & "'," & main.setupid & " from Gledger where  " & stringyear & " and Category='Expences'"
    
    'RsVDr.Open "SELECT GenLedger,sum(amount) as DAmount  FROM VOUCHERS   WHERE " & stringyear & " and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103) and DebitorCredit='D' GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
    RsVDr.Open "SELECT GenLedger,sum(amount) as DAmount  FROM VOUCHERS   WHERE " & stringyear & " and " & wheredate_vc & " and DebitorCredit='D' GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    'RsVCr.Open "SELECT GenLedger,sum(amount) as CAmount  FROM VOUCHERS  where   " & stringyear & " and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103) and DebitorCredit='C' GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
    RsVCr.Open "SELECT GenLedger,sum(amount) as CAmount  FROM VOUCHERS  where   " & stringyear & " and " & wheredate_vc & " and DebitorCredit='C' GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText

    'rs1.Open "select GenLedger,sum(Netamount) as SAmount from invoicea where   " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs1.Open "select GenLedger,sum(Netamount) as SAmount from invoicea where   " & stringyear & " and " & wheredate_inv & "  GROUP BY GenLedger", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    'rs2.Open "select GenLedger,  sum(Netamount) as  SAmount from CREDITa where  " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs2.Open "select GenLedger,  sum(Netamount) as  SAmount from CREDITa where  " & stringyear & " and " & wheredate_inv & "  GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    'rs3.Open "select GenLedger,  sum(NETamount) as SAmount  from casha where  " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)and SUBLEDGER <>'CASH PARTY'  GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs3.Open "select GenLedger,  sum(NETamount) as SAmount  from casha where  " & stringyear & " and  " & wheredate_inv & " and SUBLEDGER <>'CASH PARTY'  GROUP BY GenLedger ", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    'rs4.Open "select pgld, sum(na) as SAmount  from Cnf1a where    " & stringyear & " and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  and dc ='D' GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs4.Open "select pgld, sum(na) as SAmount  from Cnf1a where    " & stringyear & " and (convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(toDate.value) + "',103))  and dc ='D' GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    'rs4c.Open "select pgld, sum(na) as SAmount  from Cnf1a where  " & stringyear & " and  convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  and dc ='C' GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs4c.Open "select pgld, sum(na) as SAmount  from Cnf1a where  " & stringyear & " and  (convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" + Trim(toDate.value) + "',103))  and dc ='C' GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    'rs5.Open "select pgld, sum(na) as SAmount from dnfa where  " & stringyear & " and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  and dc ='D'  GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs5.Open "select pgld, sum(na) as SAmount from dnfa where  " & stringyear & " and (convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(toDate.value) + "',103))  and dc ='D'  GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    'rs5C.Open "select pgld, sum(na) as SAmount from dnfa where   " & stringyear & " and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) and  dc ='C'  GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs5C.Open "select pgld, sum(na) as SAmount from dnfa where   " & stringyear & " and (convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" + Trim(FromDate.value) + "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" + Trim(toDate.value) + "',103)) and  dc ='C'  GROUP BY pgld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    'rs6.Open "select gld, sum(a) as SAmount from Cnf1B where   " & stringyear & " and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103) and  dc ='D'  GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs6.Open "select gld, sum(a) as SAmount from Cnf1B where   " & stringyear & " and (convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" & Trim(FromDate.value) & "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & Trim(toDate.value) & "',103)) and  dc ='D'  GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    'rs6C.Open "select gld, sum(a) as SAmount from Cnf1B where   " & stringyear & " and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103) and dc ='C' GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs6C.Open "select gld, sum(a) as SAmount from Cnf1B where   " & stringyear & " and (convert(smalldatetime,cnd,103)>=convert(smalldatetime,'" & Trim(FromDate.value) & "',103) and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & Trim(toDate.value) & "',103)) and dc ='C' GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    'rs7.Open "select gld, sum(a) as SAmount from dnfB where    " & stringyear & " and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103) and dc ='D' GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs7.Open "select gld, sum(a) as SAmount from dnfB where    " & stringyear & " and (convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" & Trim(FromDate.value) & "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & Trim(toDate.value) & "',103) ) and dc ='D' GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    'rs7C.Open "select gld, sum(a) as SAmount from dnfB where   " & stringyear & " and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103) and dc ='C' GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    rs7C.Open "select gld, sum(a) as SAmount from dnfB where   " & stringyear & " and (convert(smalldatetime,dnd,103)>=convert(smalldatetime,'" & Trim(FromDate.value) & "',103) and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & Trim(toDate.value) & "',103)) and dc ='C' GROUP BY gld ", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    
    RS.Open "select * from TemprptTrialBalance where " & stringyear & " and userid=" & main.UId, con, adOpenKeyset, adLockReadOnly, adCmdText
    'a = con.Execute("Select sum(Gamount) as aa from InvoiceA  WHERE   " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" & Trim(date2.Text) & "',103) ")(0).value
    a = con.Execute("Select sum(Gamount) as aa from InvoiceA  WHERE   " & wheredate_inv & " ")(0).value
    
    If IsNull(a) Then a = 0
    
    con.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & a & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = 'SALES'"
    'a = con.Execute("Select sum(Gamount) as aa from CreditA where   " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)")(0).value
    a = con.Execute("Select sum(Gamount) as aa from CreditA WHERE   " & wheredate_inv & " ")(0).value
    
    If IsNull(a) Then a = 0
    
    con.Execute "Update TemprptTrialBalance set DAmount = (DAmount + " & a & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = 'SALES RETURN'"
    'a = con.Execute("Select sum(Gamount) as aa from CashA  where   " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)")(0).value
    a = con.Execute("Select sum(Gamount) as aa from CashA  WHERE   " & wheredate_inv & " ")(0).value
    
    If IsNull(a) Then a = 0
    
    con.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & a & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = 'SALES'"
    'a = con.Execute("Select sum(BAA) as aa from CashA  where   " & stringyear & " and convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)")(0).value
    a = con.Execute("Select sum(BAA) as aa from CashA  WHERE   " & wheredate_inv & " ")(0).value
    
    If IsNull(a) Then a = 0
    con.Execute "Update TemprptTrialBalance set dAmount = (dAmount + " & a & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = 'CASH-IN-HAND'"
    Dim tRS1 As New ADODB.Recordset
    
    '*******For Invoicec
    'tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from INVOICEC  where  " & stringyear & " and Debitorcredit='Debit' AND convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from INVOICEC  where  " & stringyear & " and Debitorcredit='Debit' AND " & wheredate_inv & "  GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    If tRS1.RecordCount > 0 Then
    While Not tRS1.EOF
    con.Execute "Update TemprptTrialBalance set DAmount = (DAmount + " & tRS1!amt & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = '" & tRS1!Genledger & "'"
    DoEvents
    DoEvents
    tRS1.MoveNext
    Wend
    End If
   
    If tRS1.State = 1 Then tRS1.close
    'tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from INVOICEC where  " & stringyear & " and Debitorcredit='Credit'  AND convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from INVOICEC where  " & stringyear & " and Debitorcredit='Credit'  AND " & wheredate_inv & " GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    If tRS1.RecordCount > 0 Then
        While Not tRS1.EOF
            con.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & tRS1!amt & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = '" & tRS1!Genledger & "'"
            DoEvents
            tRS1.MoveNext
        Wend
    End If
    
    
    
    
    If tRS1.State = 1 Then tRS1.close
    'tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CreditC WHERE   " & stringyear & " and Debitorcredit='Debit' AND convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)  GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CreditC WHERE   " & stringyear & " and Debitorcredit='Debit' AND " & wheredate_inv & "  GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    If tRS1.RecordCount > 0 Then
       While Not tRS1.EOF
          con.Execute "Update TemprptTrialBalance set DAmount = (DAmount + " & tRS1!amt & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = '" & tRS1!Genledger & "'"
          tRS1.MoveNext
       Wend
    End If
    If tRS1.State = 1 Then tRS1.close
    'tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CreditC WHERE   " & stringyear & " and Debitorcredit='Credit' AND convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103)   GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CreditC WHERE   " & stringyear & " and Debitorcredit='Credit' AND " & wheredate_inv & "  GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    If tRS1.RecordCount > 0 Then
        While Not tRS1.EOF
                    con.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & tRS1!amt & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = '" & tRS1!Genledger & "'"
            tRS1.MoveNext
        Wend
    End If
    
    '******* End of CreditC
    
    
    '*******For CashC
    If tRS1.State = 1 Then tRS1.close
    'tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CashC WHERE   " & stringyear & " and Debitorcredit='Debit' AND convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CashC WHERE   " & stringyear & " and Debitorcredit='Debit' AND " & wheredate_inv & " GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    
    If tRS1.RecordCount > 0 Then
       While Not tRS1.EOF
               con.Execute "Update TemprptTrialBalance set DAmount = (DAmount + " & tRS1!amt & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = '" & tRS1!Genledger & "'"
          tRS1.MoveNext
       Wend
    End If
    If tRS1.State = 1 Then tRS1.close
    'tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CashC WHERE   " & stringyear & " and Debitorcredit='Credit' AND convert(smalldatetime,invoicedate,103)<=convert(smalldatetime,'" + Trim(date2.Text) + "',103) GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    tRS1.Open "SELECT GENLEDGER, sum(amount) AS amt from CashC WHERE   " & stringyear & " and Debitorcredit='Credit' AND " & wheredate_inv & " GROUP BY GENLEDGER", con, adOpenStatic, adLockReadOnly
    
    If tRS1.RecordCount > 0 Then
        While Not tRS1.EOF
                   con.Execute "Update TemprptTrialBalance set CAmount = (CAmount + " & tRS1!amt & ")  where  " & stringyear & " and userid=" & main.UId & " and Gledger = '" & tRS1!Genledger & "'"
            tRS1.MoveNext
        Wend
    End If
    '*******For End of CashC
    If Not RS.BOF Then
           Do While Not RS.EOF
               
 
               
               '''OPBALANCE = RS!OpeningBalance
               OPBALANCE = 0
               SDamount = 0
               SCamount = 0
            
            If RS!gledger = "REBATE & DISCOUNT" Then
               'MsgBox ("s")
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
            'If rs3cash.RecordCount > 0 Then rs3cash.MoveFirst
            
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


End Sub
Sub Genrate()
    Set RS = New ADODB.Recordset
    main.reportname = "Gen. Ledger Trial"
    Dim called1, called2 As Boolean
    Dim MaxLine As Integer
    Dim T1, T2, T3, T4, T5, T6, T7, T8 As Integer
    Dim paperWidth As Integer
    Dim xtemp As String
    Dim trs As ADODB.Recordset
    Set trs = New ADODB.Recordset
    'con.Execute "delete from Winrpt where UID=" & UId & ""
    con.Execute "delete from Winrpt "
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
           
'           Print #1, ""
'           Print #1, Chr(27) + Chr(15) + Chr(14)
'           Print #1, Tab(115); "Page No:  " & Pno
'           Print #1, Tab((((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2) - 15) + LEFTM); Chr(27) + Chr(15) + Chr(14); dspace(Trim(kkk!cname))
'           Print #1, Tab(((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2) + LEFTM); Chr(27) + Chr(15); dspace(Trim(kkk!add1)); Chr(27) + Chr(14)
'           Line = Line + 5
        End If
        If trs.State = 1 Then trs.close
        
        trs.Open "select * from treport where " & stringyear, con, adOpenKeyset, adLockReadOnly, adCmdText
        

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
              
              If RS!gledger = "REBATE & DISCOUNT" Then
              'MsgBox "a"
              End If
              
              DbB = RS!damount - RS!camount + RS!OpeningBalance
              Print #1, Tab(1); RS!gledger; Tab(65); IIf(DbB > 0, rsets(Trim(Format(DbB, "0.00")), 13), ""); Tab(102); IIf(DbB < 0, rsets(Trim(Format(Str(Abs(DbB)), "0.00")), 13), "")
              con.Execute "insert into winrpt(Party,Narration,op,Receipt,Payment,closing,closing1,Description,FromDate,toDate,uid,OpDes) values('" & RS!gledger & "','" & RS!gledger & "'," & RS!OpeningBalance & "," & 0 & "," & 0 & "," & IIf(DbB > 0, DbB, 0) & "," & IIf(DbB < 0, DbB, 0) & ",'" & "Gen. Ledger Trial Balance" & "','" & Format(FromDate.value, "MM/dd/yyyy") & "','" & Format(toDate.value, "MM/dd/yyyy") & "'," & UId & ",'" & toDate.value & "')"
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
             '  Print #1, Tab(LEFTM); "NET DIFFERENCE "; Tab(65); rsets(Format(Trim(Abs(bal)), "0.00"), 12)
               Line = Line + 1
            Else
              If bal <> 0 Then

                 Print #1, ""
                  Line = Line + 2
              End If
            End If
            If GSumDr > GSumCr Then
               neta = GSumDr
            Else
               neta = GSumCr
            End If
            
'            Print #1, Tab(LEFTM); Chr(27) + Chr(71); repli("-", paperWidth)
'            Print #1, Tab(LEFTM); "* * * NET TOTAL * * * "; Tab(65); IIf(GSumDr <> 0, rsets(Format(Trim(GSumDr), "0.00"), 12), ""); Tab(102); IIf(GSumCr <> 0, rsets(Format(Trim(GSumCr), "0.00"), 12), "");
'            Print #1, Tab(LEFTM); repli("-", paperWidth); Chr(27) + Chr(72)
'            Line = Line + 3
            
            Do While Line <= 72
               Print #1, " "
               Line = Line + 1
            Loop
            If trs.State = 1 Then trs.close
            Close #1
End Sub
Private Sub cmdPrint_Click()

DSNNew

If cboReport.Text = "State Wise Sale" Then
   Dim s As String
    s = ""
    For I = 0 To cbostate.ListCount - 1
    If cbostate.Selected(I) = True Then
        If s = "" Then
           s = "{tempLedger1.party}='" & cbostate.List(I) & "'"
        Else
           s = s & " Or " & "{tempLedger1.party}='" & cbostate.List(I) & "'"
        End If
    End If
    Next
  
  
  
    cr.Reset
    cr.ReportFileName = rptPath & "/GST_StateWiseSales.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    If s <> "" Then
     cr.ReplaceSelectionFormula s
    End If
    
    
    cr.Formulas(0) = "fdate='" & FromDate.value & "'"
    cr.Formulas(1) = "tdate='" & toDate.value & "'"
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowExportBtn = True
    cr.WindowShowRefreshBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1
    

Else
  


    cr.Reset
    cr.ReportFileName = rptPath & "/GST_GLTrialBalance.rpt"
    cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    cr.Formulas(0) = "fdate='" & FromDate.value & "'"
    cr.Formulas(1) = "tdate='" & toDate.value & "'"
    cr.WindowShowPrintSetupBtn = True
    cr.WindowShowExportBtn = True
    cr.WindowShowRefreshBtn = True
    cr.WindowState = crptMaximized
    cr.Action = 1


End If



End Sub
Private Sub cmdView_Click()

Dim str_date As String
Dim str_date1 As String
Dim cnd_date As String
Dim cash_date As String

Dim s As String

str_date1 = "(invoicedate>=convert(smalldatetime,'" & FromDate.value & "',103) and invoicedate<=convert(smalldatetime,'" & toDate.value & "',103))"
str_date = "(INVOICEA.invoicedate>=convert(smalldatetime,'" & FromDate.value & "',103) and INVOICEA.invoicedate<=convert(smalldatetime,'" & toDate.value & "',103))"
cnd_date = "(CNF1A.cnd>=convert(smalldatetime,'" & FromDate.value & "',103) and CNF1A.cnd<=convert(smalldatetime,'" & toDate.value & "',103))"
cash_date = "(invoicedate>=convert(smalldatetime,'" & FromDate.value & "',103) and invoicedate<=convert(smalldatetime,'" & toDate.value & "',103))"


If cboReport.Text = "" Then
   MsgBox "select report type ...", vbCritical
   cboReport.SetFocus
   Exit Sub
End If


If cboReport.Text = "State Wise Sale" Then

    s = ""
    For I = 0 To cbostate.ListCount - 1
    If cbostate.Selected(I) = True Then
        If s = "" Then
           s = "{tempLedger1.party}='" & cbostate.List(I) & "'"
        Else
           s = s & " Or " & "{tempLedger1.party}='" & cbostate.List(I) & "'"
        End If
    End If
    Next

   con.Execute "delete from tempLedger1"
   
   
   con.Execute "insert into tempLedger1(party,Balance,cr)  SELECT  states, sum(GAMOUNT) ,'0' FROM invoiceaQry  where " & str_date1 & " group by states"
   con.Execute "insert into tempLedger1(party,Balance,cr) " & _
   " SELECT      SLEDGER.states , sum(INVOICEC.AMOUNT),'0' FROM INVOICEC INNER JOIN " & _
   "INVOICEA ON INVOICEC.INVOICENO = INVOICEA.INVOICENO INNER JOIN " & _
   "SLEDGER ON INVOICEA.SUBLEDGER = SLEDGER.SUBLEDGER where (INVOICEC.GENLEDGER='SALES' and INVOICEC.DEBITORCREDIT='Credit' and " & str_date & ") group by  SLEDGER.states"
   
   con.Execute "insert into tempLedger1(party,Balance,cr) " & _
   " SELECT      SLEDGER.states , (sum(INVOICEC.AMOUNT)*-1) as amt,'0' FROM INVOICEC INNER JOIN " & _
   "INVOICEA ON INVOICEC.INVOICENO = INVOICEA.INVOICENO INNER JOIN " & _
   "SLEDGER ON INVOICEA.SUBLEDGER = SLEDGER.SUBLEDGER where (INVOICEC.GENLEDGER='SALES' and INVOICEC.DEBITORCREDIT='Debit' and " & str_date & ") group by  SLEDGER.states"
   
   '=Cash=====================================================================================
   
   con.Execute "insert into tempLedger1(party,Balance,cr)  SELECT  states, sum(GAMOUNT) ,'0' FROM casha  where " & str_date1 & " and states is not null  group by states"
   con.Execute "insert into tempLedger1(party,Balance,cr)  SELECT  states , sum(AMOUNT),'0' FROM cashcQry where (states is not null and GENLEDGER='SALES' and DEBITORCREDIT='Credit' and " & cash_date & ") group by  states"
   con.Execute "insert into tempLedger1(party,Balance,cr)  SELECT  states , (sum(AMOUNT)*-1) as amt,'0' FROM cashcQry where (states is not null and GENLEDGER='SALES' and DEBITORCREDIT='Debit' and " & cash_date & ") group by  states"
   
    
   '==========================================================================================
   
   con.Execute "insert into tempLedger1(party,Balance,cr) " & _
   " SELECT      SLEDGER.states , (sum(CNF1B.A)*-1) as amt,'0' FROM CNF1B INNER JOIN " & _
   " CNF1A  ON CNF1B.cnn = CNF1A.cnn INNER JOIN " & _
   " SLEDGER ON CNF1A.PSLD = SLEDGER.SUBLEDGER where (CNF1B.GLD ='SALES' and CNF1B.DC='D' and " & cnd_date & ") group by  SLEDGER.states"
   
   con.Execute "insert into tempLedger1(party,Balance,cr) " & _
   " SELECT      SLEDGER.states , (sum(CNF1B.A)) as amt,'0' FROM CNF1B INNER JOIN " & _
   " CNF1A  ON CNF1B.cnn = CNF1A.cnn INNER JOIN " & _
   " SLEDGER ON CNF1A.PSLD = SLEDGER.SUBLEDGER where (CNF1B.GLD ='SALES' and CNF1B.DC='C' and " & cnd_date & ") group by  SLEDGER.states"

'''=============================================================================================
  con.Execute "insert into tempLedger1(party,cr,Balance)  SELECT  SLEDGER.states, sum(creditA.GAMOUNT) ,'0' FROM SLEDGER INNER JOIN creditA ON SLEDGER.SUBLEDGER=creditA.SUBLEDGER where (credita.invoicedate>=convert(smalldatetime,'" & FromDate.value & "',103) and creditA.invoicedate<=convert(smalldatetime,'" & toDate.value & "',103))  group by SLEDGER.states"
  con.Execute "insert into tempLedger1(party,cr,Balance)  SELECT  SLEDGER.states, sum(CNF1B.A),0 FROM SLEDGER INNER JOIN (CNF1B INNER JOIN CNF1A ON CNF1B.CNN=CNF1A.CNN) ON SLEDGER.SUBLEDGER=CNF1A.PSLD WHERE (CNF1B.GLD='SALES RETURN' and CNF1B.cnd>=convert(smalldatetime,'" & FromDate.value & "',103) and CNF1B.cnd<=convert(smalldatetime,'" & toDate.value & "',103) ) group by SLEDGER.states;"

 con.Execute "insert into tempLedger1(party,cr,Balance)  SELECT  SLEDGER.states,sum(creditc.AMOUNT*-1) ,0 " & _
 " FROM SLEDGER INNER JOIN (creditC INNER JOIN creditA ON creditc.INVOICENO=creditA.INVOICENO) ON SLEDGER.SUBLEDGER=creditA.SUBLEDGER " & _
 " WHERE (((creditC.GENLEDGER)='SALES RETURN') AND ((creditC.DEBITORCREDIT)='Credit') and creditC.AMOUNT>0 and (credita.invoicedate>=convert(smalldatetime,'" & FromDate.value & "',103) and creditA.invoicedate<=convert(smalldatetime,'" & toDate.value & "',103))) group by SLEDGER.states "

 con.Execute "insert into tempLedger1(party,cr,Balance)  SELECT  SLEDGER.states,sum(creditc.AMOUNT) ,0 " & _
 " FROM SLEDGER INNER JOIN (creditC INNER JOIN creditA ON creditc.INVOICENO=creditA.INVOICENO) ON SLEDGER.SUBLEDGER=creditA.SUBLEDGER " & _
 " WHERE (((creditC.GENLEDGER)='SALES RETURN') AND ((creditC.DEBITORCREDIT)='Debit') and creditC.AMOUNT>0 and (credita.invoicedate>=convert(smalldatetime,'" & FromDate.value & "',103) and creditA.invoicedate<=convert(smalldatetime,'" & toDate.value & "',103))) group by SLEDGER.states "

 con.Execute "insert into tempLedger1(party,cr,Balance)  SELECT  SLEDGER.states,sum(DNFB.A*-1) ,0 " & _
" FROM SLEDGER INNER JOIN (DNFB INNER JOIN DNFA ON DNFB.DNN=DNFA.DNN) ON SLEDGER.SUBLEDGER=DNFA.psld " & _
 " WHERE (((DNFB.GLD)='SALES RETURN') AND ((DNFB.DC)='C') and DNFB.A>0 and (DNFA.DND>=convert(smalldatetime,'" & FromDate.value & "',103) and DNFA.DND<=convert(smalldatetime,'" & toDate.value & "',103))) group by SLEDGER.states "

 con.Execute "insert into tempLedger1(party,Balance,cr)  SELECT  SLEDGER.states,sum(DNFB.A) ,0 " & _
" FROM SLEDGER INNER JOIN (DNFB INNER JOIN DNFA ON DNFB.DNN=DNFA.DNN) ON SLEDGER.SUBLEDGER=DNFA.psld " & _
 " WHERE (((DNFB.GLD)='SALES') AND ((DNFB.DC)='C') and DNFB.A>0 and (DNFA.DND>=convert(smalldatetime,'" & FromDate.value & "',103) and DNFA.DND<=convert(smalldatetime,'" & toDate.value & "',103))) group by SLEDGER.states "

  
  
  
  
cmdview.Enabled = False

Else
  
  If COMBOGENLEDGER.Text = "" Then
     genLedgerTrial
     Genrate
  End If
  
  cmdview.Enabled = False

End If

cmdExcel.Enabled = True

End Sub

Private Sub COMBOGENLEDGER_Click()
If COMBOGENLEDGER.Text = "" Then
  cmdPrint.Enabled = True
  cmdExcel.Enabled = False
Else
  cmdPrint.Enabled = False
  cmdExcel.Enabled = True
End If
End Sub

Private Sub Form_Load()

Me.Top = 1500
Me.Left = 100

If RS.State = 1 Then RS.close
RS.Open "select states from SLEDGER group by states", con
While RS.EOF = False

If Not IsNull(RS(0)) Then
 If RS(0) <> "" Then
 cbostate.AddItem RS(0)
 End If
End If
 
 RS.MoveNext
Wend



FromDate.value = Format(from_date, "dd/MM/yyyy")
toDate.value = Format(to_date, "dd/MM/yyyy")

End Sub


