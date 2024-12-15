VERSION 5.00
Begin VB.Form frmCurValue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Currency Values"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4695
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   465
      Left            =   2040
      TabIndex        =   5
      Top             =   1980
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   465
      Left            =   780
      TabIndex        =   4
      Top             =   1980
      Width           =   1140
   End
   Begin VB.TextBox txtCvalue 
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   1080
      Width           =   1590
   End
   Begin VB.ComboBox cboCName 
      Height          =   315
      ItemData        =   "frmCurValues.frx":0000
      Left            =   300
      List            =   "frmCurValues.frx":000D
      TabIndex        =   2
      Top             =   1050
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Value"
      Height          =   195
      Left            =   2460
      TabIndex        =   1
      Top             =   780
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Currency Name"
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frmCurValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCName_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from CurValues where CName='" & cboCName.Text & "'", CON

If Not rs.EOF Then
txtCvalue.Text = rs!CValue
End If



End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
CON.Execute "delete from CurValues where CName='" & cboCName.Text & "'"
CON.Execute "insert into CurValues(CName,CValue) values('" & cboCName.Text & "','" & txtCvalue & "')"
cboCName.Text = ""
txtCvalue = ""
MsgBox "data updated", vbInformation
End Sub

Private Sub Form_Load()


If rs.State = 1 Then rs.Close
rs.Open "select * from CurValues where CName='" & cboCName.Text & "'", CON

If Not rs.EOF Then
txtCvalue.Text = rs!CValue
End If

End Sub

