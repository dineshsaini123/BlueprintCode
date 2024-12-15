VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLogout 
   Caption         =   "Set DSN"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUserName_ 
      Height          =   345
      Left            =   1440
      TabIndex        =   2
      Top             =   900
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Login"
      Height          =   435
      Left            =   1140
      TabIndex        =   4
      Top             =   1500
      Width           =   975
   End
   Begin VB.CommandButton cmdAddbasil 
      Caption         =   "Add"
      Height          =   555
      Left            =   6180
      TabIndex        =   1
      Top             =   300
      Width           =   495
   End
   Begin VB.TextBox txtPath 
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   5955
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   435
      Left            =   60
      TabIndex        =   3
      Top             =   1500
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   6060
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Enter User Name :"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "frmLogout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddbasil_Click()
cd.ShowOpen
txtPath.Text = cd.FileName

End Sub

Private Sub Command1_Click()

If RS.State = 1 Then RS.close
RS.Open "select UserName from UserDSN where UserName='" & com_name & "' and usrename_='" & Trim(txtUserName_) & "'", CCON
If RS.EOF = False Then
CCON.Execute "update UserDSN set Path='" & txtPath.Text & "',usrename_='" & Trim(txtUserName_) & "' where UserName='" & com_name & "' and usrename_='" & Trim(txtUserName_) & "'"
Else
CCON.Execute "insert into UserDSN(UserName,Path,usrename_) values('" & com_name & "','" & txtPath & "','" & Trim(txtUserName_) & "')"
End If


End Sub
Private Sub Command2_Click()
Unload Me
login.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Unload frmLogout
   DoEvents
   DoEvents
   login.Show 1
End If
End Sub

Private Sub Form_Load()
'Set CON = New ADODB.Connection
'Set CON_blue = New ADODB.Connection
Me.Caption = com_name
txtUserName_ = com_user

If RS.State = 1 Then RS.close
RS.Open "select UserName,Path from UserDSN where UserName='" & com_name & "' and usrename_='" & txtUserName_ & "'", CCON
If RS.EOF = False Then
txtPath = RS!Path & ""
End If


End Sub
