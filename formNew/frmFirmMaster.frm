VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmFirmMaster 
   Caption         =   "Firm Mster"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   Icon            =   "frmFirmMaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   6735
   Begin VB.TextBox txtadd2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   840
      MaxLength       =   100
      TabIndex        =   2
      Top             =   1440
      Width           =   4005
   End
   Begin VB.ComboBox desc 
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   10815
      TabIndex        =   14
      Top             =   6765
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox ob 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   10905
      TabIndex        =   13
      Top             =   6765
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox cid 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   4860
      MaxLength       =   10
      TabIndex        =   12
      Top             =   600
      Width           =   405
   End
   Begin VB.TextBox txtFirmName 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   840
      MaxLength       =   100
      TabIndex        =   0
      Top             =   600
      Width           =   3990
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   810
      ScaleHeight     =   870
      ScaleWidth      =   5085
      TabIndex        =   9
      Top             =   2340
      Width           =   5085
      Begin VB.CommandButton Help 
         Caption         =   "&Help"
         Height          =   450
         Left            =   240
         TabIndex        =   11
         Top             =   150
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   720
         Left            =   1305
         Picture         =   "frmFirmMaster.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton Del 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   720
         Left            =   2535
         Picture         =   "frmFirmMaster.frx":0BF0
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton Abandon 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   720
         Left            =   60
         Picture         =   "frmFirmMaster.frx":17D4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton REPORTCD 
         Caption         =   "&Print"
         Height          =   585
         Left            =   4860
         TabIndex        =   10
         Top             =   1020
         Width           =   1245
      End
      Begin VB.CommandButton close 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   720
         Left            =   3780
         Picture         =   "frmFirmMaster.frx":23B8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   75
         Width           =   1245
      End
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   10890
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   6780
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox add2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   11190
      MaxLength       =   100
      TabIndex        =   7
      Top             =   6780
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.TextBox txtadd1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   840
      MaxLength       =   100
      TabIndex        =   1
      Top             =   1005
      Width           =   4005
   End
   Begin Crystal.CrystalReport Cr1 
      Left            =   6180
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   " Add2 :"
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
      Left            =   0
      TabIndex        =   21
      Top             =   1485
      Width           =   1740
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   765
      Top             =   2295
      Width           =   5145
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10650
      TabIndex        =   20
      Top             =   6735
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ledger Description"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10665
      TabIndex        =   19
      Top             =   6735
      Visible         =   0   'False
      Width           =   600
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
      Left            =   855
      TabIndex        =   18
      Top             =   240
      Width           =   2190
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Firm :"
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
      Left            =   90
      TabIndex        =   17
      Top             =   645
      Width           =   795
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Address2 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   11115
      TabIndex        =   16
      Top             =   6555
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   " Add1 :"
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
      Left            =   0
      TabIndex        =   15
      Top             =   1050
      Width           =   1740
   End
End
Attribute VB_Name = "frmFirmMaster"
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

txtFirmName.Text = ""
txtadd1 = ""
txtadd2 = ""
max_
txtFirmName.SetFocus
End Sub
Private Sub close_Click()
Unload Me
End Sub
Sub max_()

If RS.State = 1 Then RS.close
RS.Open "select max(Id) from FirmMaster", con
If IsNull(RS(0)) Then
   cid = 1
Else
   cid = RS(0) + 1
End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If
End Sub



Private Sub Del_Click()

X = MsgBox("Are you sure you wish to delete the selected item ", 4, "Confirmation")
If X = 6 Then
   
   con.Execute "Delete from firmMaster where id= '" & cid.Text & "'"
   txtFirmName.Text = ""
   txtadd1.Text = ""
   txtadd2.Text = ""
   
   txtFirmName.SetFocus
   
End If

End Sub

Private Sub Form_Load()

Me.Top = 1500
Me.Left = 1500

Me.Width = 6500
Me.Height = 5000



'maxNo
max_

BackColorFrom Me

End Sub



Private Sub save_Click()
 
Dim voucher As Boolean

Set RS = New ADODB.Recordset
If txtFirmName = "" Then
   MsgBox "Enter Firm Name ...", vbCritical
   txtFirmName.SetFocus
   Exit Sub
End If

On Error GoTo xx1


If RS.State = 1 Then RS.close
RS.Open "select * from FirmMaster where id=" & cid & "", con, adOpenDynamic, adLockOptimistic

If RS.EOF = True Then
        
   RS.AddNew
   RS!firmname = UCase(txtFirmName.Text)
   RS!add1 = Trim(txtadd1.Text)
   RS!add2 = Trim(txtadd2.Text)
   RS.update
    
   MsgBox " Record Saved ", vbInformation
   
   txtFirmName.Text = ""
   txtadd1.Text = ""
   txtadd2.Text = ""
   
   txtFirmName.SetFocus
Else
   RS!firmname = UCase(txtFirmName.Text)
   RS!add1 = Trim(txtadd1.Text)
   RS!add2 = Trim(txtadd2.Text)
   RS.update

   MsgBox " Record Modify ", vbInformation
   
   txtFirmName.Text = ""
   txtadd1.Text = ""
   txtadd2.Text = ""
   
   txtFirmName.SetFocus

End If

Exit Sub

xx1:

MsgBox err.DESCRIPTION

   txtFirmName.Text = ""
   txtadd1.Text = ""
   txtadd2.Text = ""

End Sub
Private Sub txtFirmName_GotFocus()

If PopUpValue1 <> "" Then

 txtFirmName.Text = PopUpValue1
 txtadd1.Text = PopUpValue2
 txtadd2.Text = PopUpValue3
 
 cid = popupvalue4
 
 PopUpValue1 = ""
 PopUpValue2 = ""
 PopUpValue3 = ""
 popupvalue4 = ""

End If

End Sub

Private Sub txtFirmName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   value = "Select FirmName,Add1,Add2,Id from FirmMaster order by firmname"
   popuplistModel10 value, con
End If
End Sub
