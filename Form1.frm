VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   5445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "customer transfer"
      Height          =   600
      Left            =   1710
      TabIndex        =   2
      Top             =   2220
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Item Transfer"
      Height          =   435
      Left            =   840
      TabIndex        =   1
      Top             =   3375
      Width           =   1470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agent Transfer"
      Height          =   495
      Left            =   3285
      TabIndex        =   0
      Top             =   3270
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Private Sub Command1_Click()
rs2.Open "agentmaster", CON, adOpenDynamic, adLockPessimistic
rs.Open "select * from agent1", CON, adOpenDynamic, adLockPessimistic
While rs.EOF <> True
rs2.AddNew
rs2!Agentname = rs!Name
rs2!CITY = rs!add1
'rs2!add2 = rs!add2
rs2.Update
rs.MoveNext
Wend
rs2.Close
rs.Close

End Sub

Private Sub Command2_Click()
rs2.Open "books", CON, adOpenDynamic, adLockPessimistic
rs.Open "select * from item1", CON, adOpenDynamic, adLockPessimistic
While rs.EOF <> True
rs2.AddNew
rs2!Bookcode = rs!code
rs2!Bookname = rs!particular
rs2!size1 = rs!size1
rs2!unit1 = rs!unit1
rs2!size2 = rs!size2
rs2!unit2 = rs!unit2
rs2!per = rs!per
rs2!quality = rs!quality
rs2!rate = rs!rate
rs2.Update
rs.MoveNext
Wend
rs.Close
rs2.Close
End Sub

Private Sub Command3_Click()
rs2.Open "sledger", CON, adOpenDynamic, adLockPessimistic
rs.Open "select * from cust1", CON, adOpenDynamic, adLockPessimistic
While rs.EOF <> True
rs2.AddNew
rs2!gledger = "SUNDRY DEBTORS"
rs2!subledger = rs!cust_code & " " & rs!Name
rs2!address1 = rs!add1
rs2!distcode = rs!add2
'rs2!add2 = rs!add2
rs2.Update
rs.MoveNext
Wend
rs2.Close
rs.Close
End Sub
