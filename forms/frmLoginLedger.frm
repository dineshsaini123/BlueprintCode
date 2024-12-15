VERSION 5.00
Begin VB.Form frmLoginl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1425
   ClientLeft      =   4200
   ClientTop       =   2070
   ClientWidth     =   3600
   Icon            =   "frmLoginLedger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3600
   Begin VB.TextBox txtPass1 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   555
      Width           =   2040
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   1470
      TabIndex        =   2
      Top             =   1020
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   225
      Width           =   2040
   End
   Begin VB.Label pass 
      Caption         =   "New Password"
      Height          =   225
      Left            =   150
      TabIndex        =   4
      Top             =   630
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Old Password"
      Height          =   165
      Left            =   150
      TabIndex        =   3
      Top             =   270
      Width           =   1065
   End
End
Attribute VB_Name = "frmLoginl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
 
 If change_Pass = "b" Then
    
    If RS.State = 1 Then RS.Close
    RS.Open "select * from pass where pass='" & txtpass.Text & "'", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
       RS.Fields("pass").Value = txtPass1.Text
       RS.Update
       MsgBox "Password changed !!", vbInformation
       Unload frmLoginl
    Else
    
      MsgBox "Invalid Old Password !!", vbInformation
      txtpass.SetFocus
      Exit Sub
    
    End If
 
 
 Else
 
    
    
    If RS.State = 1 Then RS.Close
    RS.Open "select * from pass where pass='" & txtpass.Text & "'", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
       cp = RS.Fields(0).Value
    Else
       cp = "cp1"
    End If
    
    
    
    Screen.MousePointer = vbHourglass
    If LCase(txtpass.Text) <> cp Or txtpass.Text = "" Then
        strledger = cp
    Else
        strledger = cp
    End If
 
    
    
    Unload frmLoginl
    MainMenu.Toolbar1.Visible = False
    
    
    frmBillList.Show
 
 End If
 
End Sub

Private Sub Label2_Click()

End Sub
Private Sub Form_Load()
    If change_Pass = "a" Then
       pass.Visible = False
       txtPass1.Visible = False
       Label1.Caption = "Password"
    Else
       pass.Visible = True
       txtPass1.Visible = True
       Label1.Caption = "Old Password"
    End If

End Sub

Private Sub txtpass_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call cmdOk_Click
   End If
End Sub
