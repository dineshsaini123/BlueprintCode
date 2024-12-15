VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBankRecon 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Bank Reconciliation"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9480
   ScaleWidth      =   16005
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker fromDate 
      Height          =   330
      Left            =   630
      TabIndex        =   19
      Top             =   7155
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   582
      _Version        =   393216
      Format          =   62586881
      CurrentDate     =   40604
   End
   Begin MSComCtl2.DTPicker ToDate 
      Height          =   330
      Left            =   2160
      TabIndex        =   18
      Top             =   7155
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   582
      _Version        =   393216
      Format          =   62586881
      CurrentDate     =   40604
   End
   Begin VB.CheckBox Check_Clear 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Clear Cheque "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5580
      TabIndex        =   15
      Top             =   8160
      Width           =   1815
   End
   Begin VB.ComboBox cboSub 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   7575
      Width           =   2115
   End
   Begin VB.ComboBox cboGen 
      Height          =   315
      ItemData        =   "frmBankRecon.frx":0000
      Left            =   180
      List            =   "frmBankRecon.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   7575
      Width           =   1620
   End
   Begin MSComCtl2.DTPicker txtPostedDate 
      Height          =   315
      Left            =   1260
      TabIndex        =   11
      Top             =   60
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   62586881
      CurrentDate     =   40527
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Exit"
      Height          =   435
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8100
      Width           =   1215
   End
   Begin VB.CheckBox Check_Recon 
      BackColor       =   &H00C0FFC0&
      Caption         =   "All Cheque "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4080
      TabIndex        =   8
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpDate 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Update"
      Height          =   435
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8100
      Width           =   1215
   End
   Begin VB.CommandButton cmdRef 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&View"
      Height          =   435
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8100
      Width           =   1215
   End
   Begin VB.ListBox List_cr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6030
      Left            =   8040
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1020
      Width           =   7875
   End
   Begin VB.ListBox List_dr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6030
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1020
      Width           =   7875
   End
   Begin VB.Label lblOp 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6015
      TabIndex        =   25
      Top             =   7545
      Width           =   1740
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Bal."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4605
      TabIndex        =   24
      Top             =   7650
      Width           =   1485
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Closing:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9270
      TabIndex        =   23
      Top             =   7695
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblClosing 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   10170
      TabIndex        =   22
      Top             =   7650
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1890
      TabIndex        =   21
      Top             =   7200
      Width           =   285
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "From "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   20
      Top             =   7155
      Width           =   645
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12780
      TabIndex        =   17
      Top             =   7605
      Width           =   1155
   End
   Begin VB.Label lblBal 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   13980
      TabIndex        =   16
      Top             =   7545
      Width           =   1755
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Posted Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   12
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   " Bank Reconciliation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   15795
   End
   Begin VB.Label lblCr_Total 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   13980
      TabIndex        =   5
      Top             =   7110
      Width           =   1755
   End
   Begin VB.Label lblDr_Total 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   7125
      Width           =   1755
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "  RECEIPT                                                                                  (Dr.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   660
      Width           =   7875
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "  PAYMENT                                                                                (Cr.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   8040
      TabIndex        =   2
      Top             =   660
      Width           =   7875
   End
End
Attribute VB_Name = "frmBankRecon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub AddAmount()

Dim DR_SUM As Double
Dim CR_SUM As Double
Dim J As Integer

Dim str1 As String


DR_SUM = 0
CR_SUM = 0
J = 0


List_cr.Clear
List_dr.Clear


str1 = ""
str1 = "((convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(fromdate.Value) + "',103)) or (convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(fromdate.Value) + "',103) and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(todate.Value) + "',103)))"

'===================================================================================================================================
'===================================================================================================================================
'===================================================================================================================================



If rs.State = 1 Then rs.Close

If Check_Recon.Value = 1 Then

rs.Open "select VoucherDate,VoucherNumber,SubLedger,DESCRIPTION,amount,vsno,BankRecon,postedDate from VOUCHERS " & _
"where (genledger='BANK  A/C' or genledger='BANK  A/C') and VoucherType='P' and DebitorCredit='C' " & _
" and fyear='" & main.session & "' and setupid=" & main.setupid & " and genledger='" & cboGen.Text & "' and subledger='" & cboSub.Text & "' and " & str1 & _
" order by VoucherDate,VoucherNumber", CON, adOpenKeyset, adLockReadOnly

Else


If Check_Clear.Value = 1 Then


rs.Open "select VoucherDate,VoucherNumber,SubLedger,DESCRIPTION,amount,vsno,BankRecon,postedDate from VOUCHERS " & _
"where (genledger='BANK  A/C' or genledger='BANK  A/C') and VoucherType='P' and DebitorCredit='C' " & _
" and fyear='" & main.session & "' and setupid=" & main.setupid & " and (len(BankRecon)>0) and genledger='" & cboGen.Text & "' and subledger='" & cboSub.Text & "' and " & str1 & _
" order by VoucherDate,VoucherNumber", CON, adOpenKeyset, adLockReadOnly


Else


rs.Open "select VoucherDate,VoucherNumber,SubLedger,DESCRIPTION,amount,vsno,BankRecon,postedDate from VOUCHERS " & _
"where (genledger='BANK  A/C' or genledger='BANK  A/C') and VoucherType='P' and DebitorCredit='C' " & _
" and fyear='" & main.session & "' and setupid=" & main.setupid & " and (len(BankRecon)=0 or BankRecon is null) and genledger='" & cboGen.Text & "' and subledger='" & cboSub.Text & "' and " & str1 & _
" order by VoucherDate,VoucherNumber", CON, adOpenKeyset, adLockReadOnly


End If


End If

While rs.EOF = False
       
       
       
List_cr.AddItem rs!VOUCHERNUMBER & Space(4 - Len(rs!VOUCHERNUMBER)) & Mid(rs!VoucherDate, 1, 5) & " " & Mid(rs!DESCRIPTION, 1, 35) & "-" & Mid(rs!postedDate, 1, 5) & "" & Space(45 - Len(Mid(rs!DESCRIPTION, 1, 35))) & rs!amount & Space(30) & rs!vsno

CR_SUM = CR_SUM + rs!amount
       
If rs!BankRecon = "y" Then
   List_cr.Selected(J) = True
End If
       
J = J + 1
       
rs.MoveNext
Wend



J = 0




str1 = ""
str1 = "((convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(fromdate.Value) + "',103)) or (convert(smalldatetime,voucherdate,103)>=convert(smalldatetime,'" + Trim(fromdate.Value) + "',103) and convert(smalldatetime,voucherdate,103)<=convert(smalldatetime,'" + Trim(todate.Value) + "',103)))"


If rs.State = 1 Then rs.Close

If Check_Recon.Value = 1 Then

rs.Open "select VoucherDate,VoucherNumber,SubLedger,DESCRIPTION,amount,vsno,BankRecon,postedDate from VOUCHERS " & _
"where (genledger='BANK  A/C' or genledger='BANK  A/C') and VoucherType='R' and DebitorCredit='D' " & _
"and genledger='" & cboGen.Text & "' and subledger='" & cboSub.Text & "' and fyear='" & main.session & "' and setupid=" & main.setupid & " and  " & str1 & "  order by VoucherDate,VoucherNumber", CON, adOpenKeyset, adLockReadOnly

Else

If Check_Clear.Value = 1 Then


rs.Open "select VoucherDate,VoucherNumber,SubLedger,DESCRIPTION,amount,vsno,BankRecon,postedDate from VOUCHERS " & _
"where (genledger='BANK  A/C' or genledger='BANK  A/C') and VoucherType='R' and DebitorCredit='D' " & _
"and genledger='" & cboGen.Text & "' and subledger='" & cboSub.Text & "' and fyear='" & main.session & "' and setupid=" & main.setupid & " and (len(BankRecon)>0) and " & str1 & " order by VoucherDate,VoucherNumber", CON, adOpenKeyset, adLockReadOnly

Else

rs.Open "select VoucherDate,VoucherNumber,SubLedger,DESCRIPTION,amount,vsno,BankRecon,postedDate from VOUCHERS " & _
"where (genledger='BANK  A/C' or genledger='BANK  A/C') and VoucherType='R' and DebitorCredit='D' " & _
"and genledger='" & cboGen.Text & "' and subledger='" & cboSub.Text & "' and fyear='" & main.session & "' and setupid=" & main.setupid & " and (len(BankRecon)=0 or BankRecon is null) and " & str1 & " order by VoucherDate,VoucherNumber", CON, adOpenKeyset, adLockReadOnly


End If

End If


While rs.EOF = False
       

List_dr.AddItem rs!VOUCHERNUMBER & Space(4 - Len(rs!VOUCHERNUMBER)) & Mid(rs!VoucherDate, 1, 5) & " " & Mid(rs!DESCRIPTION, 1, 35) & "-" & Mid(rs!postedDate, 1, 5) & "" & Space(45 - Len(Mid(rs!DESCRIPTION, 1, 35))) & rs!amount & Space(30) & rs!vsno
DR_SUM = DR_SUM + rs!amount

If rs!BankRecon = "y" Then
   List_dr.Selected(J) = True
End If

       
J = J + 1
       
rs.MoveNext
Wend



lblCr_Total.Caption = CR_SUM
lblDr_Total.Caption = DR_SUM


End Sub
Sub ClosingCal()

Dim DR_SUM As Double
Dim CR_SUM As Double
Dim J As Integer

DR_SUM = 0
CR_SUM = 0
J = 0


'===================================================================================================================================
'===================================================================================================================================
'===================================================================================================================================

If rs.State = 1 Then rs.Close

If Check_Recon.Value = 1 Then

rs.Open "select VoucherDate,VoucherNumber,SubLedger,DESCRIPTION,amount,vsno,BankRecon,postedDate from VOUCHERS " & _
"where (genledger='BANK  A/C' or genledger='BANK  A/C') and VoucherType='P' and DebitorCredit='C' " & _
" and fyear='" & main.session & "' and setupid=" & main.setupid & " and genledger='" & cboGen.Text & "' and subledger='" & cboSub.Text & "'" & _
" and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(fromdate.Value) + "',103)" & _
" order by VoucherDate,VoucherNumber", CON, adOpenKeyset, adLockReadOnly

Else


If Check_Clear.Value = 1 Then


rs.Open "select VoucherDate,VoucherNumber,SubLedger,DESCRIPTION,amount,vsno,BankRecon,postedDate from VOUCHERS " & _
"where (genledger='BANK  A/C' or genledger='BANK  A/C') and VoucherType='P' and DebitorCredit='C' " & _
" and fyear='" & main.session & "' and setupid=" & main.setupid & " and (len(BankRecon)>0) and genledger='" & cboGen.Text & "' and subledger='" & cboSub.Text & "'" & _
" and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(fromdate.Value) + "',103)" & _
" order by VoucherDate,VoucherNumber", CON, adOpenKeyset, adLockReadOnly


Else


rs.Open "select VoucherDate,VoucherNumber,SubLedger,DESCRIPTION,amount,vsno,BankRecon,postedDate from VOUCHERS " & _
"where (genledger='BANK  A/C' or genledger='BANK  A/C') and VoucherType='P' and DebitorCredit='C' " & _
" and fyear='" & main.session & "' and setupid=" & main.setupid & " and (len(BankRecon)=0 or BankRecon is null) and genledger='" & cboGen.Text & "'" & _
" and subledger='" & cboSub.Text & "' and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(fromdate.Value) + "',103)" & _
" order by VoucherDate,VoucherNumber", CON, adOpenKeyset, adLockReadOnly


End If


End If

While rs.EOF = False
       
       

CR_SUM = CR_SUM + rs!amount
       
       
J = J + 1
       
rs.MoveNext
Wend






'==================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================
'==================================================================================================================================





J = 0

If rs.State = 1 Then rs.Close

If Check_Recon.Value = 1 Then

rs.Open "select VoucherDate,VoucherNumber,SubLedger,DESCRIPTION,amount,vsno,BankRecon,postedDate from VOUCHERS " & _
"where (genledger='BANK  A/C' or genledger='BANK  A/C') and VoucherType='R' and DebitorCredit='D' and genledger='" & cboGen.Text & "' and subledger='" & cboSub.Text & "'" & _
" and fyear='" & main.session & "' and setupid=" & main.setupid & "" & _
" and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(fromdate.Value) + "',103)  order by VoucherDate,VoucherNumber", CON, adOpenKeyset, adLockReadOnly

Else

If Check_Clear.Value = 1 Then


rs.Open "select VoucherDate,VoucherNumber,SubLedger,DESCRIPTION,amount,vsno,BankRecon,postedDate from VOUCHERS " & _
"where (genledger='BANK  A/C' or genledger='BANK  A/C') and VoucherType='R' and DebitorCredit='D' and genledger='" & cboGen.Text & "' and subledger='" & cboSub.Text & "'" & _
" and fyear='" & main.session & "' and setupid=" & main.setupid & " and " & _
" (len(BankRecon)>0) and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(fromdate.Value) + "',103) order by VoucherDate,VoucherNumber", CON, adOpenKeyset, adLockReadOnly

Else

rs.Open "select VoucherDate,VoucherNumber,SubLedger,DESCRIPTION,amount,vsno,BankRecon,postedDate from VOUCHERS " & _
"where (genledger='BANK  A/C' or genledger='BANK  A/C') and VoucherType='R' and DebitorCredit='D' and genledger='" & cboGen.Text & "' and subledger='" & cboSub.Text & "'" & _
" and fyear='" & main.session & "' and setupid=" & main.setupid & " and (len(BankRecon)=0 or BankRecon is null) " & _
" and convert(smalldatetime,voucherdate,103)<convert(smalldatetime,'" + Trim(fromdate.Value) + "',103) order by VoucherDate,VoucherNumber", CON, adOpenKeyset, adLockReadOnly


End If

End If


While rs.EOF = False
       

DR_SUM = DR_SUM + rs!amount
 
       
J = J + 1
       
rs.MoveNext
Wend



lblClosing.Caption = Round((DR_SUM - CR_SUM), 0)

If rs.State = 1 Then rs.Close
rs.Open "select YEAROPENING from SLEDGER where gledger='" & cboGen.Text & "' and SUBLEDGER='" & cboSub.Text & "'", CON
If rs.EOF = False Then
   lblOp.Caption = rs(0)
End If

'lblBal.Caption = Round((Val(lblDr_Total) - Val(lblCr_Total)), 0)

lblBal.Caption = (Val(lblClosing.Caption) + Val(lblOp.Caption))


If Val(lblBal.Caption) < 0 Then
   lblBal.Caption = Abs(lblBal.Caption) & "  CR"
Else
   lblBal.Caption = Abs(lblBal.Caption) & "  DR"
End If


''If Val(lblClosing.Caption) < 0 Then
''   lblClosing.Caption = Abs(lblClosing.Caption) & "  CR"
''Else
''   lblClosing.Caption = Abs(lblClosing.Caption) & "  DR"
''End If


End Sub


Private Sub cboGen_Click()
     
   cboSub.Clear
     
   If rs.State = 1 Then rs.Close
   rs.Open "select subledger from sledger where  gledger='" & cboGen.Text & "'", CON, adOpenKeyset, adLockReadOnly
   While rs.EOF = False
   cboSub.AddItem rs(0)
   rs.MoveNext
   Wend

  cboSub.ListIndex = 0
  cboSub.SetFocus
  
End Sub
Private Sub cboSub_Click()

''If (cboGen.Text <> "" And cboSub.Text <> "") Then
''   cmdRef_Click
''End If

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdRef_Click()

On Error GoTo ss1
    
Screen.MousePointer = vbHourglass
AddAmount
   
ClosingCal
    
Screen.MousePointer = vbDefault

Exit Sub

ss1:

MsgBox "" & err.DESCRIPTION

End Sub
Private Sub cmdUpDate_Click()

Screen.MousePointer = vbHourglass

For i = 0 To List_dr.ListCount - 1
   

sss = Trim(Right(List_dr.List(i), 20))
   
If rs.State = 1 Then rs.Close
rs.Open "select BankRecon from vouchers where vsno=" & sss & " and fyear='" & main.session & "' and BankRecon='y'", CON, adOpenKeyset, adLockReadOnly
If rs.EOF = False Then
   
If List_dr.Selected(i) = False Then
    CON.Execute "update VOUCHERS set BankRecon='',postedDate='' where vsno=" & sss & " and fyear='" & main.session & "'"
End If
  
Else


If List_dr.Selected(i) = True Then
   CON.Execute "update VOUCHERS set BankRecon='y',postedDate='" & txtPostedDate.Value & "' where vsno=" & sss & " and fyear='" & main.session & "'"
'Else
'   CON.Execute "update VOUCHERS set BankRecon='',postedDate='' where vsno=" & sss & " and fyear='" & main.session & "'"
End If



End If

Next


'-----------------------------------------------------------------

'-----------------------------------------------------------------

For i = 0 To List_cr.ListCount - 1


sss = Trim(Right(List_cr.List(i), 20))

If rs.State = 1 Then rs.Close
rs.Open "select BankRecon from vouchers where vsno=" & sss & " and fyear='" & main.session & "' and BankRecon='y'", CON, adOpenKeyset, adLockReadOnly
If rs.EOF = False Then

If List_cr.Selected(i) = False Then
  CON.Execute "update VOUCHERS set BankRecon='',postedDate='' where vsno=" & sss & " and fyear='" & main.session & "'"
End If


Else


If List_cr.Selected(i) = True Then
   CON.Execute "update VOUCHERS set BankRecon='y',postedDate='" & txtPostedDate.Value & "' where vsno=" & sss & " and fyear='" & main.session & "'"
'Else
'   CON.Execute "update VOUCHERS set BankRecon='',postedDate='' where vsno=" & sss & " and fyear='" & main.session & "'"
End If


End If

Next


MsgBox "Updated .....", vbInformation
Call cmdRef_Click


Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
   txtPostedDate.Value = Format(Date, "dd/MM/yyyy")
   
End Sub
