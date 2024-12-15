VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmPapersize 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtsizeinfo 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   1020
      MaxLength       =   100
      TabIndex        =   19
      Top             =   990
      Width           =   4005
   End
   Begin VB.TextBox add2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   11370
      MaxLength       =   100
      TabIndex        =   12
      Top             =   6765
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   11070
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   6765
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   990
      ScaleHeight     =   870
      ScaleWidth      =   5085
      TabIndex        =   4
      Top             =   1845
      Width           =   5085
      Begin VB.CommandButton close 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   720
         Left            =   3780
         Picture         =   "frmPapersize.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton REPORTCD 
         Caption         =   "&Print"
         Height          =   585
         Left            =   4860
         TabIndex        =   9
         Top             =   1020
         Width           =   1245
      End
      Begin VB.CommandButton Abandon 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   720
         Left            =   45
         Picture         =   "frmPapersize.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton Del 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   720
         Left            =   2535
         Picture         =   "frmPapersize.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   720
         Left            =   1305
         Picture         =   "frmPapersize.frx":23AC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton Help 
         Caption         =   "&Help"
         Height          =   450
         Left            =   240
         TabIndex        =   5
         Top             =   150
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.TextBox na 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   1035
      MaxLength       =   100
      TabIndex        =   3
      Top             =   585
      Width           =   2250
   End
   Begin VB.TextBox cid 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   3285
      MaxLength       =   10
      TabIndex        =   2
      Top             =   585
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox ob 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   11085
      TabIndex        =   1
      Top             =   6750
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox desc 
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   10995
      TabIndex        =   0
      Top             =   6750
      Visible         =   0   'False
      Width           =   390
   End
   Begin Crystal.CrystalReport Cr1 
      Left            =   6120
      Top             =   3510
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
      Left            =   180
      TabIndex        =   20
      Top             =   1035
      Width           =   1740
   End
   Begin VB.Label LABELUSERNAME 
      BackStyle       =   0  'Transparent
      Height          =   330
      Left            =   2475
      TabIndex        =   18
      Top             =   6840
      Width           =   6885
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
      Left            =   11295
      TabIndex        =   17
      Top             =   6540
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Size :"
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
      Left            =   270
      TabIndex        =   16
      Top             =   630
      Width           =   795
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
      Left            =   1035
      TabIndex        =   15
      Top             =   225
      Width           =   2190
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ledger Description"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10845
      TabIndex        =   14
      Top             =   6720
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10830
      TabIndex        =   13
      Top             =   6720
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   945
      Top             =   1800
      Width           =   5145
   End
End
Attribute VB_Name = "frmPapersize"
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
   
   con.Execute "Delete from SizeMaster where size1= '" & cid.Text & "' and " & stringyear & ""
   na.Text = ""
   txtsizeinfo = ""
   
   na.SetFocus
End If

End Sub

Private Sub Form_Load()

Me.Top = 2000
Me.Left = 2000

'maxNo

BackColorFrom Me

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
   value = "Select size1 as [Paper Size],size_info as Remarks from SizeMaster where " & stringyear & " order by size1"
   popuplistModel10 value, con
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
RS.Open "select * from SizeMaster where size1='" & Trim(cid.Text) & "' and " & stringyear & "", con, adOpenDynamic, adLockOptimistic

If RS.EOF = True Then
        
   RS.AddNew
   RS!size1 = (na.Text)
   RS!size_info = Trim(txtsizeinfo.Text)
   RS!fyear = session
   RS!setupid = setupid
   RS.update
    
   MsgBox " Record Saved "
   
   na.Text = ""
   txtsizeinfo.Text = ""
   
   na.SetFocus
Else
   RS!size1 = (na.Text)
   RS!size_info = Trim(txtsizeinfo.Text)
   RS.update
End If

Exit Sub

xx1:

MsgBox err.DESCRIPTION
na.Text = ""
txtsizeinfo.Text = ""
na.SetFocus

End Sub



