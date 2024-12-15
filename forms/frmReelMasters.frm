VERSION 5.00
Begin VB.Form frmReelMasters 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Height          =   390
      Left            =   150
      TabIndex        =   5
      Top             =   75
      Width           =   3990
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   150
      TabIndex        =   1
      Top             =   3600
      Width           =   4035
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   75
         Picture         =   "frmReelMasters.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   1230
      End
      Begin VB.CommandButton cmdDelete_3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   1350
         Picture         =   "frmReelMasters.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   1230
      End
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2625
         Picture         =   "frmReelMasters.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Width           =   1230
      End
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   150
      TabIndex        =   0
      Top             =   450
      Width           =   3990
   End
End
Attribute VB_Name = "frmReelMasters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDelete_3_Click()
  CON.Execute "delete from Reel_ReamMaster where Name ='" & txtName & "' and Category='" & Paper_Master & "'"
  txtName = ""
  Add
End Sub
Private Sub cmdExit_12_Click()
Unload Me
End Sub
Private Sub cmdSave_2_Click()
   CON.Execute "insert into Reel_ReamMaster(Name,Category,fyear,setupid) values('" & txtName & "','" & Paper_Master & "','" & main.session & "','" & main.setupid & "')"
   Add
   txtName = ""
End Sub
Sub Add()
   List1.Clear
   If rs.State = 1 Then rs.Close
   rs.Open "select Name from Reel_ReamMaster where category='" & Paper_Master & "' order by Name", CON, adOpenKeyset, adLockReadOnly
   While rs.EOF = False
        List1.AddItem rs(0)
        rs.MoveNext
   Wend
End Sub
Private Sub Form_Load()
Add

Me.TOP = 2200
Me.Left = 6000

End Sub
Private Sub Form_Unload(Cancel As Integer)
frmReals.AddItem
End Sub

Private Sub List1_Click()
txtName = List1.Text
End Sub
