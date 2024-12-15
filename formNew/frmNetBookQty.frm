VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNetBookQty 
   ClientHeight    =   3228
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   7512
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3228
   ScaleWidth      =   7512
   Begin VB.TextBox txtscid 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6240
      TabIndex        =   6
      Top             =   1020
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtSchoolName 
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Top             =   1020
      Width           =   5235
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   945
      ScaleHeight     =   876
      ScaleWidth      =   3768
      TabIndex        =   0
      Top             =   1845
      Width           =   3765
      Begin VB.CommandButton save 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View"
         Height          =   720
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   1065
      End
      Begin VB.CommandButton REPORTCD 
         Caption         =   "&Print"
         Height          =   585
         Left            =   4860
         TabIndex        =   3
         Top             =   1020
         Width           =   1245
      End
      Begin VB.CommandButton close 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   720
         Left            =   2460
         Picture         =   "frmNetBookQty.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         Width           =   1065
      End
      Begin VB.CommandButton cmdPrint_7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   720
         Left            =   1260
         Picture         =   "frmNetBookQty.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
         Width           =   1065
      End
   End
   Begin Crystal.CrystalReport cr 
      Left            =   5100
      Top             =   2040
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker txtFrom 
      Height          =   315
      Left            =   960
      TabIndex        =   9
      Top             =   300
      Width           =   1335
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
      Format          =   142606337
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txtto 
      Height          =   315
      Left            =   2760
      TabIndex        =   10
      Top             =   300
      Width           =   1275
      _ExtentX        =   2244
      _ExtentY        =   550
      _Version        =   393216
      Format          =   142606337
      CurrentDate     =   42409
   End
   Begin VB.Label lblAson 
      Caption         =   "To"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   300
      Width           =   315
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "State :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   8
      Top             =   1020
      Width           =   495
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 For Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   960
      TabIndex        =   7
      Top             =   720
      Width           =   2715
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   900
      Top             =   1800
      Width           =   3885
   End
End
Attribute VB_Name = "frmNetBookQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dt_from As Date
Dim dt_to As Date
Dim dt_str As String
Dim bb_2 As Boolean
Dim bb1 As Boolean
Dim Edit As Boolean
Dim Add As Boolean
Dim CON_next As ADODB.Connection

Private Sub close_Click()
Unload Me
End Sub

Private Sub cmbAgentName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cboPayment.SetFocus
End If
End Sub
Sub max_sp()
If RS.State = 1 Then RS.close
RS.Open "select max(DNo) from DonnationMain", con
If Not IsNull(RS(0)) Then
   txtSponsorshipNo = RS(0) + 1
Else
   txtSponsorshipNo = 1
End If

End Sub
Private Sub cmdPrint_7_Click()

DoEvents
DoEvents
DoEvents
DSNNew

MainMenu.cr1.Reset
MainMenu.cr1.ReportFileName = rptPath & "/SchoolWiseSale_new.rpt"
MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
MainMenu.cr1.ReplaceSelectionFormula "{tmpSchoolWiseBkWiseNSale.states}='" & txtSchoolName & "'"
MainMenu.cr1.WindowShowPrintSetupBtn = True
MainMenu.cr1.WindowShowExportBtn = True
MainMenu.cr1.WindowState = crptMaximized
MainMenu.cr1.WindowShowRefreshBtn = True
MainMenu.cr1.WindowShowSearchBtn = True
MainMenu.cr1.Action = 1


End Sub
Private Sub Form_Load()

If RS.State = 1 Then RS.close
RS.Open "select * from financialyear where fyear='" & session & "'", CCON
If RS.EOF = False Then
   txtFrom = RS!FromDate
   txtto = RS!toDate
End If


Me.Width = 8000
Me.Height = 4000

BackColorFrom Me
End Sub

''Private Sub List1_sc_DblClick()
''txtSponsorshipNo = Trim(Mid(List1_sc.Text, 1, InStr(List1_sc.Text, "=") - 1))
''SearchData
''End Sub

Private Sub save_Click()
  Screen.MousePointer = vbHourglass
    
    If txtSchoolName <> "" Then
    con.Execute "exec SchoolWise_BookWiseSalenet '" & txtSchoolName & "','" & txtFrom & "','" & txtto & "'"
    End If
  
  cmdPrint_7.Enabled = True
  Screen.MousePointer = vbDefault
    


End Sub
Private Sub txtSchoolName_GotFocus()
If PopUpValue1 <> "" Then
   txtScId = PopUpValue2
   txtSchoolName = PopUpValue1
   
   PopUpValue1 = ""
   PopUpValue2 = ""
End If
End Sub

Private Sub txtSchoolName_KeyDown(KeyCode As Integer, Shift As Integer)
    
If KeyCode = 113 Then

searchType = "party"
value = "select States from SLEDGER where gledger='SUNDRY DEBTORS' and " & stringyear & " group by States"
popuplist_client value, con
set_focus = True

End If

End Sub
