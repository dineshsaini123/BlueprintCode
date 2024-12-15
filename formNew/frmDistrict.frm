VERSION 5.00
Begin VB.Form frmDistrict 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "District"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   810
      TabIndex        =   1
      Top             =   360
      Width           =   4275
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   45
      TabIndex        =   0
      Top             =   1305
      Width           =   7260
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
         Picture         =   "frmDistrict.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1170
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
         Left            =   4797
         Picture         =   "frmDistrict.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   45
         Width           =   1200
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
         Left            =   5985
         Picture         =   "frmDistrict.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   45
         Width           =   1200
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
         Left            =   3609
         Picture         =   "frmDistrict.frx":23AC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   45
         Width           =   1200
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
         Left            =   2421
         Picture         =   "frmDistrict.frx":27B9
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   45
         Width           =   1200
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
         Left            =   1233
         Picture         =   "frmDistrict.frx":339D
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   45
         Width           =   1200
      End
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
         Picture         =   "frmDistrict.frx":3F81
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   45
         Width           =   1200
      End
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   0
      Top             =   1260
      Width           =   7350
   End
   Begin VB.Label lblAgn 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   855
      TabIndex        =   10
      Top             =   720
      Width           =   4245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   405
      Width           =   510
   End
End
Attribute VB_Name = "frmDistrict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit1 As Boolean
Private Sub cmdAdd_1_Click()

   
'txtRId = ""

txtname = ""
lblAgn.Caption = ""

cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False

txtname.SetFocus

   
   
End Sub

Private Sub cmdDelete_3_Click()

If MsgBox("want To delete ?", vbInformation + vbYesNo) = vbYes Then

 CON.BeginTrans
 CON.Execute "delete from  DISTRICTS where Transportname='" & txtname & "' and " & stringyear
 CON.CommitTrans
 
End If

cmdAdd_1_Click
End Sub

Private Sub cmdEdit_4_Click()

cmdEdit_4.Enabled = False
cmdSave_2.Enabled = True
cmdDelete_3.Enabled = True
cmdSave_2.SetFocus

edit1 = True

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub



Private Sub cmdSave_2_Click()


If txtname = "" Then
   MsgBox "Enter DISTRICT Name. ...", vbInformation
   txtname.SetFocus
   Exit Sub
End If
   



CON.BeginTrans

If edit1 = True Then
   CON.Execute "delete from  [DISTRICTS] where " & stringyear & " and DISTRICTNAME='" & txtname & "'"
 
End If

   

CON.Execute "INSERT INTO  [DISTRICTS]" & _
           "([DISTRICTNAME]" & _
           ",[Fyear]" & _
           ",[setupid]" & _
     ") Values" & _
           "('" & UCase(txtname) & "'" & _
           ",'" & main.session & "'" & _
           ",'" & main.setupid & "')"
        
        
CON.CommitTrans

'MsgBox "Date Saved ....", vbInformation
cmdSave_2.Enabled = False

Call cmdAdd_1_Click
   


End Sub

Private Sub cmdSearch_Click()
   
   popuplistModel10 "select DISTRICTNAME,AGENTNAME from [DISTRICTS] where  " & stringyear & " order by DISTRICTNAME", CON
 
   cmdSave_2.Enabled = False
   cmdEdit_4.Enabled = True

End Sub
Private Sub cmdSearch_GotFocus()
  
  
  If PopUpValue1 <> "" Then
  
  txtname = PopUpValue1
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
 Me.Top = 1800
 Me.Left = 1500
 
 BackColorFrom Me
 
End Sub




