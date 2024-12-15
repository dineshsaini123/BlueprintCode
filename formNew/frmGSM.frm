VERSION 5.00
Begin VB.Form frmGSM 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GSM  Master"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox cid 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   3150
      MaxLength       =   10
      TabIndex        =   9
      Top             =   990
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox na 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   900
      MaxLength       =   100
      TabIndex        =   0
      Top             =   990
      Width           =   2250
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   855
      ScaleHeight     =   870
      ScaleWidth      =   5085
      TabIndex        =   3
      Top             =   2250
      Width           =   5085
      Begin VB.CommandButton Help 
         Caption         =   "&Help"
         Height          =   450
         Left            =   0
         TabIndex        =   8
         Top             =   30
         Width           =   15
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   720
         Left            =   1305
         Picture         =   "frmGSM.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton Del 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   720
         Left            =   2535
         Picture         =   "frmGSM.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton Abandon 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   720
         Left            =   45
         Picture         =   "frmGSM.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton REPORTCD 
         Caption         =   "&Print"
         Height          =   585
         Left            =   4860
         TabIndex        =   5
         Top             =   1020
         Width           =   1245
      End
      Begin VB.CommandButton close 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   720
         Left            =   3780
         Picture         =   "frmGSM.frx":23AC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   75
         Width           =   1245
      End
   End
   Begin VB.TextBox txtsizeinfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   885
      MaxLength       =   100
      TabIndex        =   1
      Top             =   1395
      Width           =   4005
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   810
      Top             =   2205
      Width           =   5145
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 For Search "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   900
      TabIndex        =   12
      Top             =   630
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "GSM :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   135
      TabIndex        =   11
      Top             =   1035
      Width           =   795
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   " Remarks :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   45
      TabIndex        =   10
      Top             =   1440
      Width           =   1740
   End
End
Attribute VB_Name = "frmGSM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ref As ADODB.Recordset
Dim flag As Boolean
Dim value As String
Dim RS As New ADODB.Recordset

Sub COMPINI()

na.Text = ""
add1.Text = ""

'maxNo
End Sub

Private Sub ABANDON_Click()

na.Text = ""
'maxNo

na.SetFocus
End Sub






Private Sub close_Click()
Unload Me
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If
End Sub



Private Sub Del_Click()

X = MsgBox("Are you sure you wish to delete the selected item ", 4, "Confirmation")
If X = 6 Then
   
   CON.Execute "Delete from GSMMaster where GSM= '" & cid.Text & "' and " & stringyear & ""
   na.Text = ""
   txtsizeinfo = ""
   
   na.SetFocus
End If

End Sub

Private Sub Form_Load()

Me.Top = 2000
Me.Left = 2000

'maxNo
BackColorFrom Me, 1

'formButtonValidation Del
   

End Sub

Private Sub Na_GotFocus()

If PopUpValue1 <> "" Then

na.Text = PopUpValue1
cid = PopUpValue1


PopUpValue1 = ""
PopUpValue2 = ""

End If


End Sub

Private Sub na_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   value = "Select GSM as [Paper Size],gsm_info as Remarks from GSMMaster order by GSM"
   popuplistModel10 value, CON
End If

End Sub

Private Sub save_Click()

Dim voucher As Boolean

Set RS = New ADODB.Recordset
If na = "" Then
   MsgBox "Enter Size ...", vbCritical
   Exit Sub
End If

On Error GoTo xx1


If RS.State = 1 Then RS.close
RS.Open "select * from GSMMaster where GSM='" & Trim(cid.Text) & "' and " & stringyear & "", CON, adOpenDynamic, adLockOptimistic

If RS.EOF = True Then
        
   RS.AddNew
   RS!GSM = (na.Text)
   RS!gsm_info = Trim(txtsizeinfo.Text)
   RS!fyear = session
   RS!setupid = setupid
   RS.update
    
   MsgBox " Record Saved "
   
   na.Text = ""
   txtsizeinfo.Text = ""
   
   na.SetFocus
Else
   RS!GSM = (na.Text)
   RS!gsm_info = Trim(txtsizeinfo.Text)
   
   na.Text = ""
   txtsizeinfo.Text = ""

   RS.update
End If

Exit Sub

xx1:

MsgBox Err.DESCRIPTION
na.Text = ""
txtsizeinfo.Text = ""
na.SetFocus

End Sub




