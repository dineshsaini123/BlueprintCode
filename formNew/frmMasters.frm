VERSION 5.00
Begin VB.Form frmMasters 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   4380
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      MaxLength       =   50
      TabIndex        =   5
      Top             =   195
      Width           =   4155
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   150
      TabIndex        =   1
      Top             =   3780
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
         Left            =   60
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
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Width           =   1230
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4155
   End
End
Attribute VB_Name = "frmMasters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDelete_3_Click()
  con.Execute "delete from MasterTbl where Name ='" & txtName & "' and Category='" & HeadTbl & "'"
  txtName = ""
  Add
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdSave_2_Click()
On Error GoTo aaa1:
   
If RS.State = 1 Then RS.close
RS.Open "select * from MasterTbl where name='" & txtName & "'", con
If RS.EOF = True Then
   con.Execute "insert into MasterTbl(Name,Category) values('" & txtName & "','" & HeadTbl & "')"
End If

addData HeadTbl

Add
txtName = ""
   
Exit Sub
aaa1:
MsgBox "" & err.DESCRIPTION
   
End Sub
Sub Add()
   
   List1.Clear
   If RS.State = 1 Then RS.close
   RS.Open "select Name from MasterTbl where category='" & HeadTbl & "' order by Name", con, adOpenKeyset, adLockReadOnly
   While RS.EOF = False
        List1.AddItem RS(0)
        RS.MoveNext
   Wend
   
   
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Add

Me.Top = 1000
Me.Left = 500

BackColorFrom Me, 1

End Sub
Private Sub Form_Unload(cancel As Integer)
If frmNo = "boobmaster" Then
 frmbook.AddItem
Else
 papermaker.AddItem
End If

End Sub

Private Sub List1_Click()
txtName = List1.Text
End Sub
Private Sub txtName_GotFocus()
  If PopUpValue1 <> "" Then
     txtName.Text = PopUpValue1
  End If
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    If HeadTbl = "SerName" Then
       popuplistModel10 "SELECT distinct SerName FROM BOOKS where SerName<>''  order  by SerName", con
       'popuplistModel10 value, con
    End If
End If

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     cmdSave_2_Click
  End If
  
End Sub
