VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2985
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1763.636
   ScaleMode       =   0  'User
   ScaleWidth      =   4267.509
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   2460
      TabIndex        =   4
      Top             =   2400
      Width           =   1140
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Top             =   990
      Width           =   2025
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1620
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   2025
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   780
      TabIndex        =   3
      Top             =   2400
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Session :"
      Height          =   270
      Index           =   2
      Left            =   510
      TabIndex        =   7
      Top             =   390
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name :"
      Height          =   270
      Index           =   0
      Left            =   510
      TabIndex        =   5
      Top             =   1020
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password :"
      Height          =   270
      Index           =   1
      Left            =   510
      TabIndex        =   6
      Top             =   1620
      Width           =   1080
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CON As ADODB.Connection
Dim rs As ADODB.Recordset
Option Explicit
Public LoginSucceeded As Boolean
Private Sub cmdCancel_Click()
    If rs.State = 1 Then
        rs.close
    End If
    LoginSucceeded = False
    Me.Hide
    End
End Sub
Private Sub cmdOK_Click()
    If rs.State = 1 Then
        rs.close
    End If
    rs.Open "SETUP", CON, adOpenKeyset, adLockReadOnly, adCmdTable
    rs.Find "UNAME='" + Trim(Me.txtUserName) + "'"
    If Not rs.EOF Then
        If Trim(txtPassword) = Trim(rs!Password) Then
            LoginSucceeded = True
            rs.close
            main.UserName = Trim(Me.txtUserName.Text)
            Unload Me
            MainMenu.Show
        Else
            MsgBox "Invalid Password, try again!", , "Login"
            'txtPassword.SetFocus
            'SendKeys "{Home}+{End}"
        End If
    Else
      '  MsgBox "USER  NOT FOUND"
      '  Me.txtUserName.SetFocus
      '      SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtUserName.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    ConOpen
    If rs.State = 1 Then rs.close
    rs.Open "datadir", CON, adOpenKeyset, adLockOptimistic, adCmdTable
    If Not rs.EOF Then
        Do While Not rs.EOF
            Combo1.AddItem rs!Year
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdOK.SetFocus
        KeyAscii = 0
    End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPassword.SetFocus
        KeyAscii = 0
    End If
End Sub
