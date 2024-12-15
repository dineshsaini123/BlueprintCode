VERSION 5.00
Begin VB.Form frmIssue_ReceiceMaster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Issue/Receve Master"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   Icon            =   "frmIssue_ReceiceMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtid 
      Height          =   330
      Left            =   5355
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txtCateogry 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   135
      MaxLength       =   30
      TabIndex        =   8
      Top             =   1275
      Width           =   5190
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   645
      Left            =   90
      TabIndex        =   5
      Top             =   225
      Width           =   5865
      Begin VB.OptionButton OptionReceive 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Receive Item Category"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   420
         Left            =   2895
         TabIndex        =   7
         Top             =   150
         Width           =   2895
      End
      Begin VB.OptionButton OptionIssue 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Issue Item Category"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   420
         Left            =   105
         TabIndex        =   6
         Top             =   150
         Value           =   -1  'True
         Width           =   2805
      End
   End
   Begin VB.Frame panel1 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   135
      TabIndex        =   0
      Top             =   1860
      Width           =   6270
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   675
         Left            =   30
         Picture         =   "frmIssue_ReceiceMaster.frx":57EE
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   75
         Width           =   1440
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   675
         Left            =   4740
         Picture         =   "frmIssue_ReceiceMaster.frx":63D2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   75
         Width           =   1440
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   675
         Left            =   3170
         Picture         =   "frmIssue_ReceiceMaster.frx":6FB6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   75
         Width           =   1440
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   675
         Left            =   1600
         Picture         =   "frmIssue_ReceiceMaster.frx":7B9A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   75
         Width           =   1440
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   90
      Top             =   1800
      Width           =   6360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cateogry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   9
      Top             =   960
      Width           =   1860
   End
End
Attribute VB_Name = "frmIssue_ReceiceMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_Click()

If MsgBox(" Want To Delete ... ?", vbQuestion + vbYesNo) = vbYes Then
   CON.Execute "delete from Issue_ReceiveMaster where " & stringyear & " and id=" & txtId.Text & ""
   txtCateogry.Text = ""
   txtCateogry.SetFocus
End If

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdRef_Click()
txtCateogry.Text = ""
txtCateogry.SetFocus
End Sub

Private Sub CmdSave_Click()
If MsgBox(" Want To Save ... ?", vbQuestion + vbYesNo) = vbYes Then
On Error GoTo aa1
   CON.Execute "delete from Issue_ReceiveMaster where name='" & txtCateogry.Text & "'"
   CON.Execute "insert into Issue_ReceiveMaster(Name,Category) values('" & txtCateogry.Text & "','" & IIf(OptionIssue.value = True, "Issue", "Receive") & "')"
   txtCateogry.Text = ""
   txtCateogry.SetFocus

Exit Sub
aa1:
MsgBox "" & Err.DESCRIPTION
End If
End Sub

Private Sub Form_Load()
BackColorFrom Me, panel1

Me.Top = 1500
Me.Left = 2000

'formButtonValidation cmdDelete
   
End Sub

Private Sub txtCateogry_GotFocus()
If PopUpValue1 <> "" Then
   txtId.Text = PopUpValue2
   txtCateogry.Text = PopUpValue1
''   If PopUpValue2 = "Issue" Then
''      OptionIssue.value = True
''   Else
''      OptionReceive.value = True
''   End If
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   
End If
End Sub
Private Sub txtCateogry_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
If OptionIssue.value = True Then
   popuplistModel10 "select Name,Id from Issue_ReceiveMaster where Category='Issue' order by Name", CON
  Else
   popuplistModel10 "select Name,Id from Issue_ReceiveMaster where Category='Receive' order by Name", CON
End If
End If
End Sub
