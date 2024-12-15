VERSION 5.00
Begin VB.Form frmChangePass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2910
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1_sp 
      Caption         =   "Change Password(sponsorship)"
      Height          =   435
      Left            =   3900
      TabIndex        =   10
      Top             =   1080
      Width           =   2235
   End
   Begin VB.CheckBox Check1_changeledger 
      Caption         =   "Change  Ledger Password"
      Height          =   435
      Left            =   3900
      TabIndex        =   9
      Top             =   600
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   555
      Left            =   600
      ScaleHeight     =   495
      ScaleWidth      =   3225
      TabIndex        =   5
      Top             =   2160
      Width           =   3285
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   450
         Left            =   1920
         TabIndex        =   4
         Top             =   30
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Save Password"
         Height          =   450
         Left            =   180
         TabIndex        =   3
         Top             =   30
         Width           =   1590
      End
   End
   Begin VB.TextBox txtnewpass 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1455
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1095
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1455
      TabIndex        =   0
      Top             =   195
      Width           =   2325
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1455
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   645
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&New Password"
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   8
      Top             =   1170
      Width           =   1305
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&User Name"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   6
      Top             =   270
      Width           =   1035
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Old Password"
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub Check1_changeledger_Click()

If Check1_changeledger.value = 1 Then
   lblLabels(0).Visible = False
   txtUserName.Visible = False
Else
   lblLabels(0).Visible = True
   txtUserName.Visible = True

End If
End Sub

Private Sub Check1_sp_Click()
If Check1_sp.value = 1 Then
   lblLabels(0).Visible = False
   txtUserName.Visible = False
Else
   lblLabels(0).Visible = True
   txtUserName.Visible = True

End If

End Sub

Private Sub cmdCancel_Click()
 Unload Me
    
End Sub
Private Sub cmdOK_Click()

Dim rsuser As New ADODB.Recordset
Dim ss1 As String
ss1 = 2

If (Check1_changeledger.value = 1) Then
   ss1 = 1
End If

If (Check1_sp.value = 1) Then
   ss1 = 1
End If



If (ss1 = 2) Then
    
    
    Set rsuser = New ADODB.Recordset
    
    '''rsuser.Open "Select * from UsrePermission where  fyear='" & session & "' and UserName = '" & txtUserName.Text & "' and password = '" & txtPassword.Text & "'", con, adOpenDynamic, adLockOptimistic
    rsuser.Open "Select * from UsrePermission where  (UserName = '" & txtUserName.Text & "' and password = '" & txtPassword.Text & "')", coninfo, adOpenDynamic, adLockOptimistic
    If rsuser.EOF = False Then
       rsuser!UserName = txtUserName.Text
       rsuser!Password = txtnewpass.Text
       rsuser.update
       'coninfo.Execute "update UsrePermission set [Password]='" & txtnewpass.Text & "' where  [UserName] = '" & txtUserName.Text & "' and [password] = '" & txtPassword.Text & "'"
       coninfo.Execute "update UsrePermission set [Password]='" & txtnewpass.Text & "' where  [UserName] = '" & txtUserName.Text & "'"
       MsgBox "User Password has been changed"
       txtUserName.Text = ""
       txtPassword.Text = ""
       txtnewpass.Text = ""
       
       
       
    
       Unload Me
    Else
       MsgBox "Invalid User Name/Password, try again!", , "Change Passord"
       txtPassword.SetFocus
       SendKeys "{Home}+{End}"
       Exit Sub
    End If
    
    
    


ElseIf Check1_changeledger.value = 1 Then
    
    Dim dvalue
    dvalue = encrypt(txtPassword.Text)

    Set rsuser = New ADODB.Recordset
    rsuser.Open "Select * from pass  where pass = '" & dvalue & "' and module_ = '" & module_ & "'", con, adOpenDynamic, adLockOptimistic
    If rsuser.RecordCount > 0 Then
       
       rsuser!pass = encrypt(Trim(txtnewpass.Text))
       rsuser.update
       MsgBox "User Password has been changed"
       
       txtPassword.Text = ""
       txtnewpass.Text = ""
    
       Unload Me
    Else
       MsgBox "Invalid Old Password, try again!", , "Change Passord"
       txtPassword.SetFocus
       SendKeys "{Home}+{End}"
       Exit Sub
    End If

ElseIf Check1_sp.value = 1 Then

    
    dvalue = encrypt(txtPassword.Text)


    Set rsuser = New ADODB.Recordset
    rsuser.Open "Select * from pass  where donnation = '" & dvalue & "' and module_ = '" & module_ & "'", con, adOpenDynamic, adLockOptimistic
    If rsuser.RecordCount > 0 Then
       
        rsuser!donnation = encrypt(Trim(txtnewpass.Text))
        rsuser.update
        MsgBox "User Password has been changed"
       
        txtPassword.Text = ""
        txtnewpass.Text = ""
        Unload Me
        
    Else
        
        MsgBox "Invalid Old Password, try again!", , "Change Passord"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
        
    End If



End If


    
       
End Sub

Private Sub txtnewpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"


End If
End Sub
Sub CheckPass()




End Sub
