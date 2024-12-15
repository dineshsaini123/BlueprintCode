VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmUpDateSchool 
   Caption         =   "Modify School ...."
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   7125
   Icon            =   "frmUpDate.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2490
   ScaleWidth      =   7125
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Modify"
      Height          =   675
      Left            =   675
      Picture         =   "frmUpDate.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton close 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   675
      Left            =   1845
      Picture         =   "frmUpDate.frx":044E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin MSMask.MaskEdBox txtschool 
      Height          =   285
      Left            =   690
      TabIndex        =   0
      Top             =   765
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtScid 
      Height          =   285
      Left            =   5985
      TabIndex        =   2
      Top             =   765
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "School : "
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   810
      Width           =   780
   End
End
Attribute VB_Name = "frmUpDateSchool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub close_Click()
Unload Me
End Sub

Private Sub Command1_Click()

If MsgBox("Want to Edit ?", vbQuestion + vbYesNo) = vbYes Then
   updateSchooll txtScid.Text, txtschool.Text
End If

End Sub

Private Sub Form_Load()
Me.Top = 100
Me.Left = 100

Me.Height = 3150
Me.Width = 7200

End Sub

Private Sub txtschool_GotFocus()
If RS.State = 1 Then RS.close
If PopUpValue1 <> "" Then
   
txtScid = PopUpValue1
txtschool.Text = PopUpValue2 & ", " & PopUpValue3
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
popupvalue4 = ""

End If
End Sub
Sub updateSchooll(scid As String, scname As String)

con.Execute "update INVOICEA set ScName='" & scname & "'  where ScID ='" & scid & "'"
con.Execute "update CreditA set ScName='" & scname & "'  where ScID ='" & scid & "'"
con.Execute "update INVOICEA_sp set ScName='" & scname & "' where ScID ='" & scid & "'"
con.Execute "update ORDERA set ScName='" & scname & "'  where ScID ='" & scid & "'"
con.Execute "update AppForm  set School_PartyName='" & scname & "'  where (ID ='" & scid & "' and School_Party='School')"
con.Execute "update AppForm set PName='" & scname & "'  where (code ='" & scid & "' and School_Party='party')"


End Sub
Private Sub txtschool_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   
   Screen.MousePointer = vbHourglass
   tblNo = 9
   frmSearchItem.Show
   Screen.MousePointer = vbDefault
   
End If
End Sub
