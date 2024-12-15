VERSION 5.00
Begin VB.Form frmArea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Area"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   7275
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   135
      TabIndex        =   1
      Top             =   1440
      Width           =   6945
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   45
         Picture         =   "frmArea.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   975
      End
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
         Height          =   720
         Left            =   1035
         Picture         =   "frmArea.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
         Width           =   975
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
         Height          =   720
         Left            =   1995
         Picture         =   "frmArea.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2970
         Picture         =   "frmArea.frx":23AC
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   975
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
         Height          =   720
         Left            =   5895
         Picture         =   "frmArea.frx":27B9
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   975
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   4920
         Picture         =   "frmArea.frx":339D
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   3945
         Picture         =   "frmArea.frx":3F81
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   975
      End
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   855
      TabIndex        =   0
      Top             =   450
      Width           =   4275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name :"
      Height          =   195
      Left            =   270
      TabIndex        =   9
      Top             =   495
      Width           =   510
   End
End
Attribute VB_Name = "frmArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit1 As Boolean
Private Sub cmdAdd_1_Click()



txtName = ""

cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False

txtName.SetFocus

   
   
End Sub

Private Sub cmdDelete_3_Click()

If MsgBox("want To delete ?", vbInformation + vbYesNo) = vbYes Then

 CON.BeginTrans
 CON.Execute "delete from  DISTRICTS where Transportname='" & txtName & "' and " & stringyear
 CON.CommitTrans
 
End If

cmdAdd_1_Click
End Sub

Private Sub cmdEdit_4_Click()

cmdEdit_4.Enabled = False
cmdSave_2.Enabled = True
cmdDelete_3.Enabled = True
cmdSave_2.SetFocus

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub



Private Sub cmdSave_2_Click()


If txtName = "" Then
   MsgBox "Enter DISTRICT Name. ...", vbInformation
   txtName.SetFocus
   Exit Sub
End If
   



CON.BeginTrans

If edit1 = True Then
   CON.Execute "delete from  [DISTRICTS] where DISTRICTNAME='" & txtName & "'"
 
End If

   

CON.Execute "INSERT INTO  [DISTRICTS]" & _
           "([DISTRICTNAME]" & _
           ",[Fyear]" & _
           ",[setupid]" & _
     ") Values" & _
           "('" & UCase(txtName) & "'" & _
           ",'" & main.session & "'" & _
           ",'" & main.setupid & "')"
        
        
CON.CommitTrans

'MsgBox "Date Saved ....", vbInformation
cmdSave_2.Enabled = False

Call cmdAdd_1_Click
   


End Sub

Private Sub cmdSearch_Click()

   popuplist10 "select DISTRICTNAME,AGENTNAME from [DISTRICTS] where  " & stringyear & " order by DISTRICTNAME", CON
 
   cmdSave_2.Enabled = False
   cmdEdit_4.Enabled = True

End Sub
Private Sub cmdSearch_GotFocus()
  
  
  If PopUpValue1 <> "" Then
  
  txtName = PopUpValue1
  lblAgn.Caption = PopUpValue2
  
  
  End If
   
  PopUpValue1 = ""
  PopUpValue2 = ""
  
  cmdEdit_4.Enabled = True
  

End Sub


Private Sub Form_Activate()
cmdAdd_1_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub





Private Sub Form_Load()
 Me.TOP = 1800
 Me.Left = 1500
End Sub






