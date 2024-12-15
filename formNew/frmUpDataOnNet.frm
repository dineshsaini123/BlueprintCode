VERSION 5.00
Begin VB.Form frmUpDataOnNet 
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   Icon            =   "frmUpDataOnNet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1_rep 
      Caption         =   "Create User For Representative"
      Height          =   435
      Left            =   4560
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox cboRep 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1140
      TabIndex        =   2
      Top             =   1260
      Width           =   3315
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   435
      Left            =   2400
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create User"
      Height          =   435
      Left            =   1140
      TabIndex        =   4
      Top             =   1800
      Width           =   1155
   End
   Begin VB.TextBox txtpass 
      Height          =   375
      Left            =   1140
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox txtuser 
      Height          =   375
      Left            =   1140
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Data..."
      Height          =   435
      Left            =   1170
      TabIndex        =   5
      Top             =   2475
      Width           =   2355
   End
   Begin VB.Label rep 
      Caption         =   "Rep. Name "
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   9
      Top             =   1260
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   240
      Width           =   1035
   End
End
Attribute VB_Name = "frmUpDataOnNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con_net As ADODB.Connection
Dim rsDest As New ADODB.Recordset

Private Sub Check1_rep_Click()
  If Check1_rep.value = 1 Then
     rep.Enabled = True
     cboRep.Enabled = True
  Else
     rep.Enabled = False
     cboRep.Enabled = False
  End If
End Sub

Private Sub cmdCreate_Click()

If txtuser = "" Then
   MsgBox "Please Enter User...", vbCritical
   txtuser.SetFocus
   Exit Sub
End If

If txtpass = "" Then
   MsgBox "Please Enter Password...", vbCritical
   txtpass.SetFocus
   Exit Sub
End If

If Check1_rep.value = 1 Then
If cboRep.Text = "" Then
   MsgBox "Please Select Reprasentative...", vbCritical
   cboRep.SetFocus
   Exit Sub
End If
End If


If RS.State = 1 Then RS.close
RS.Open "select * from System_Users", con_net, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
   RS.AddNew
   RS!UserName = Trim(txtuser)
   RS!Password = Trim(txtpass)
   RS!RegDate = Date
   RS!email = "email"
   If cboRep.Text <> "" Then
      RS!RepName = cboRep.Text
      RS!LoginType = "Rep"
   Else
      RS!LoginType = "Other"
   End If
   RS.update
   
   txtuser = ""
   txtpass = ""
   
   txtuser.SetFocus
   
   
   MsgBox "User Created...", vbInformation
Else
   MsgBox "User Already Created...", vbInformation
End If

End Sub

Private Sub cmdUpdate_Click()

Screen.MousePointer = vbHourglass

''Set con_net = New ADODB.Connection
''con_net.ConnectionString = "PROVIDER=SQLOLEDB;" _
''         & "SERVER=111.118.213.132;" _
''         & "Database=bluedata_net;" _
''         & "DataTypeCompatibility=80;" _
''         & "User Id=blue;" _
''         & "Password=Qfwa334%;"
''con_net.Open
'invoice a================================================
con_net.Execute "delete from invoicea"

If RS.State = 1 Then RS.close
RS.Open "select * from invoicea order by INVOICENO", con
If rsDest.State = 1 Then rsDest.close
rsDest.Open "select * from invoicea order by INVOICENO", con_net, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   rsDest.AddNew
   rsDest!invoiceNo = RS!invoiceNo
   rsDest!invoiceDate = RS!invoiceDate
   rsDest!Genledger = RS!Genledger
   rsDest!SUBLEDGER = RS!SUBLEDGER
   rsDest!netamount = RS!netamount
   rsDest!gamount = RS!gamount
   rsDest!District = RS!District
   rsDest!agentname = RS!agentname
   rsDest!transportname = RS!transportname
   rsDest!fyear = RS!fyear
   rsDest!setupid = RS!setupid
   rsDest.update
RS.MoveNext
Wend

'INVOICEA_sp================================================
con_net.Execute "delete from INVOICEA_sp"

If RS.State = 1 Then RS.close
RS.Open "select * from INVOICEA_sp order by INVOICENO", con
If rsDest.State = 1 Then rsDest.close
rsDest.Open "select * from INVOICEA_sp order by INVOICENO", con_net, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   rsDest.AddNew
   rsDest!invoiceNo = RS!invoiceNo
   rsDest!invoiceDate = RS!invoiceDate
   rsDest!Genledger = RS!Genledger
   rsDest!SUBLEDGER = RS!SUBLEDGER
   rsDest!netamount = RS!netamount
   rsDest!gamount = RS!gamount
   rsDest!District = RS!District
   rsDest!agentname = RS!agentname
   rsDest!transportname = RS!transportname
   rsDest!fyear = RS!fyear
   rsDest!setupid = RS!setupid
   rsDest.update
RS.MoveNext
Wend

'INVOICEA_spRet================================================
con_net.Execute "delete from INVOICEA_spRet"

If RS.State = 1 Then RS.close
RS.Open "select * from INVOICEA_spRet order by INVOICENO", con
If rsDest.State = 1 Then rsDest.close
rsDest.Open "select * from INVOICEA_spRet order by INVOICENO", con_net, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   rsDest.AddNew
   rsDest!invoiceNo = RS!invoiceNo
   rsDest!invoiceDate = RS!invoiceDate
   rsDest!Genledger = RS!Genledger
   rsDest!SUBLEDGER = RS!SUBLEDGER
   rsDest!netamount = RS!netamount
   rsDest!gamount = RS!gamount
   rsDest!District = RS!District
   rsDest!agentname = RS!agentname
   rsDest!transportname = RS!transportname
   rsDest!fyear = RS!fyear
   rsDest!setupid = RS!setupid
   rsDest.update
RS.MoveNext
Wend

'INVOICEA_spRet================================================
con_net.Execute "delete from CREDITA"

If RS.State = 1 Then RS.close
RS.Open "select * from CREDITA order by INVOICENO", con
If rsDest.State = 1 Then rsDest.close
rsDest.Open "select * from CREDITA order by INVOICENO", con_net, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   rsDest.AddNew
   rsDest!invoiceNo = RS!invoiceNo
   rsDest!invoiceDate = RS!invoiceDate
   rsDest!Genledger = RS!Genledger
   rsDest!SUBLEDGER = RS!SUBLEDGER
   rsDest!netamount = RS!netamount
   rsDest!gamount = RS!gamount
   rsDest!District = RS!District
   rsDest!agentname = RS!agentname
   rsDest!fyear = RS!fyear
   rsDest!setupid = RS!setupid
   rsDest.update
RS.MoveNext
Wend

'===========================================================

'INVOICEB_sp================================================
con_net.Execute "delete from INVOICEB_sp"

If RS.State = 1 Then RS.close
RS.Open "select * from INVOICEB_sp order by INVOICENO", con
If rsDest.State = 1 Then rsDest.close
rsDest.Open "select * from INVOICEB_sp order by INVOICENO", con_net, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   rsDest.AddNew
   rsDest!invoiceNo = RS!invoiceNo
   rsDest!invoiceDate = RS!invoiceDate
   rsDest!Genledger = RS!Genledger
   rsDest!SUBLEDGER = RS!SUBLEDGER
   rsDest!amount = RS!amount
   rsDest!netamount = RS!netamount
   rsDest!agentname = RS!agentname
   rsDest!fyear = RS!fyear
   rsDest!setupid = RS!setupid
   rsDest!Bookcode = RS!Bookcode
   rsDest!QUANTITY = RS!QUANTITY
   rsDest!rate = RS!rate
   rsDest!discount = RS!discount
   rsDest!PRINTORDER = RS!PRINTORDER
   rsDest!GroupName = RS!GroupName
   rsDest!BookDesc = RS!BookDesc
   rsDest.update
RS.MoveNext
Wend

'==sledger===========================================
con_net.Execute "delete from sledger"

If RS.State = 1 Then RS.close
RS.Open "select * from sledger order by SUBLEDGER", con
If rsDest.State = 1 Then rsDest.close
rsDest.Open "select * from sledger order by SUBLEDGER", con_net, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   rsDest.AddNew
   rsDest!gledger = RS!gledger
   rsDest!SUBLEDGER = RS!SUBLEDGER
   rsDest!DESCFORINVOICE = RS!DESCFORINVOICE
   rsDest!YEAROPENING = RS!YEAROPENING
   rsDest!DISCATEGORY = RS!DISCATEGORY
   rsDest!DISTCODE = RS!DISTCODE
   rsDest!address1 = RS!address1
   rsDest!address2 = RS!address2
   rsDest!address3 = RS!address3
   rsDest!Owner = RS!Owner
   rsDest!op = RS!op
   rsDest!drcr = RS!drcr
   rsDest!party = RS!party
   rsDest!Code = RS!Code
   rsDest!contactp = RS!contactp
   rsDest!Remarks = RS!Remarks
   rsDest!states = RS!states
   rsDest!email = RS!email
   rsDest!mobile = RS!mobile
   rsDest!cityId = RS!cityId
   rsDest!cityname = RS!cityname
   rsDest!fyear = RS!fyear
   rsDest!setupid = RS!setupid
   rsDest.update
RS.MoveNext
Wend
       
'voucher===========================================
con_net.Execute "delete from VOUCHERS"
If RS.State = 1 Then RS.close
RS.Open "select * from VOUCHERS order by SUBLEDGER", con
If rsDest.State = 1 Then rsDest.close
rsDest.Open "select * from VOUCHERS order by SUBLEDGER", con_net, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   rsDest.AddNew
   rsDest!vouchertype = RS!vouchertype
   rsDest!voucherDATE = RS!voucherDATE
   rsDest!VOUCHERNUMBER = RS!VOUCHERNUMBER
   rsDest!Genledger = RS!Genledger
   rsDest!SUBLEDGER = RS!SUBLEDGER
   rsDest!amount = RS!amount
   rsDest!DebitorCredit = RS!DebitorCredit
   rsDest!cbnd = RS!cbnd
   rsDest!EntryNumber = RS!EntryNumber
   rsDest!DESCRIPTION = RS!DESCRIPTION
   rsDest!CashCheck = RS!CashCheck
   rsDest!address2 = RS!address2
   rsDest!fyear = RS!fyear
   rsDest!setupid = RS!setupid
RS.MoveNext
Wend
       
Screen.MousePointer = vbDefault
   
End Sub
Private Sub Form_Load()
Screen.MousePointer = vbHourglass

Set con_net = New ADODB.Connection
 
con_net.ConnectionString = "PROVIDER=SQLOLEDB;" _
         & "SERVER=111.118.213.132;" _
         & "Database=bluedata_net;" _
         & "DataTypeCompatibility=80;" _
         & "User Id=blue;" _
         & "Password=Qfwa334%;"
con_net.Open

Screen.MousePointer = vbDefault

    If RS.State = 1 Then RS.close
    RS.Open "select Rep as Representative from SalesRepQry order by Rep", CON_blue
    cboRep.Clear
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
            Me.cboRep.AddItem RS(0)
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If
End Sub

Private Sub Label4_Click()

End Sub
