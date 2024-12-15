VERSION 5.00
Begin VB.Form frmUpdateCCMail 
   Caption         =   "Update CC Mail"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   4935
   Icon            =   "frmUpdateCCMail.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4935
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   420
      Left            =   135
      TabIndex        =   3
      Top             =   1485
      Width           =   1995
   End
   Begin VB.TextBox txtmail 
      Height          =   375
      Left            =   135
      TabIndex        =   2
      Top             =   945
      Width           =   4515
   End
   Begin VB.ComboBox cboRep 
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   540
      Width           =   4560
   End
   Begin VB.Label Label1 
      Caption         =   "RepName :"
      Height          =   330
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   4515
   End
End
Attribute VB_Name = "frmUpdateCCMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboRep_Click()
If RS.State = 1 Then RS.close
 RS.Open "select email from SalesRepQry where rep='" & cboRep & "'", CON_blue
 If RS.EOF = False Then
    txtmail.Text = RS!email
 End If

End Sub
Private Sub cmdOk_Click()
 
 con.Execute "update MailDetails set RepEmail='" & txtmail & "' where mailsended='Bulk Mail...'"
 
 MsgBox "CC Mail Modify...."
End Sub
Private Sub Form_Load()
   Set RS = New ADODB.Recordset
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
