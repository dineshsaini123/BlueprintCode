VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmTypeProduct 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Type of Creation"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14940
   ClipControls    =   0   'False
   Icon            =   "frmTypeProduct.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9855
   ScaleWidth      =   14940
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   675
      Left            =   8640
      TabIndex        =   14
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   1140
      TabIndex        =   1
      Top             =   1140
      Width           =   4155
   End
   Begin VB.TextBox txtCode 
      Height          =   345
      Left            =   1140
      TabIndex        =   0
      Top             =   660
      Width           =   1575
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   1080
      TabIndex        =   2
      Top             =   1980
      Width           =   7575
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
         Height          =   840
         Left            =   45
         Picture         =   "frmTypeProduct.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1230
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
         Height          =   840
         Left            =   1290
         Picture         =   "frmTypeProduct.frx":0BF0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
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
         Left            =   2520
         Picture         =   "frmTypeProduct.frx":17D4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   1230
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
         Height          =   840
         Left            =   3750
         Picture         =   "frmTypeProduct.frx":23B8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
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
         Left            =   6240
         Picture         =   "frmTypeProduct.frx":27C5
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   1230
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
         Height          =   840
         Left            =   4995
         Picture         =   "frmTypeProduct.frx":33A9
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   1230
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   5310
      Left            =   360
      TabIndex        =   11
      Top             =   3660
      Width           =   8370
      _cx             =   14764
      _cy             =   9366
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   15787206
      ForeColor       =   -2147483640
      BackColorFixed  =   8454143
      ForeColorFixed  =   0
      BackColorSel    =   15787206
      ForeColorSel    =   -2147483640
      BackColorBkg    =   15787206
      BackColorAlternate=   15787206
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   12
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label header 
      BackColor       =   &H8000000D&
      Caption         =   "     Type Of Product"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   0
      Width           =   13335
   End
   Begin VB.Label Label5 
      Caption         =   "* Required fields"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   420
      TabIndex        =   9
      Top             =   9060
      Width           =   2955
   End
End
Attribute VB_Name = "frmTypeProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim editval As Boolean

Private Sub cmdAdd_1_Click()
   txtCode = ""
   txtName = ""
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = False
   cmdSave_2.Enabled = True
   
   Dim o As Object
   For Each o In Me
   If TypeOf o Is textbox Then
   o.Text = ""
   End If
   Next
   
   txtCode.SetFocus
End Sub

Private Sub cmdDelete_3_Click()
On Error GoTo save:


If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
   CON.BeginTrans
   CON.Execute "delete from groups where groupcode='" & txtCode & "'"
   CON.CommitTrans
   fillGrid
   Call cmdAdd_1_Click
End If
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False
cmdSave_2.Enabled = True
Exit Sub

save:
CON.RollbackTrans
MsgBox "" & err.DESCRIPTION

End Sub

Private Sub cmdEdit_4_Click()

cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
editval = True

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub
Sub fillGrid()
    
   Dim f As New ADODB.Recordset
   If f.State = 1 Then f.Close
   f.Open "select groupcode as Code,groupname as TypeOfProduct from groups order by groupcode", CON
   Set vs.DataSource = f
End Sub

Private Sub cmdSave_2_Click()


If txtCode.Text = "" Then
   MsgBox "Plz. Enter Code ...", vbCritical
   txtCode.SetFocus
   Exit Sub
End If

If txtName.Text = "" Then
   MsgBox "Plz. Enter Type Of Name ...", vbCritical
   txtName.SetFocus
   Exit Sub
End If

On Error GoTo save:

If editval = False Then


CON.BeginTrans
CON.Execute "exec insertData_Groups '" & txtCode & "','" & txtName & "','" & main.username & "'," & _
"'" & main.username & "','" & main.session & "'," & main.setupid & ""
CON.CommitTrans

'MsgBox ("Data Saved")


Else

CON.BeginTrans
CON.Execute "exec UpdateData_Groups '" & txtCode & "','" & txtName & "','" & main.username & "'," & _
"'" & main.session & "'," & main.setupid & ""
CON.CommitTrans

editval = False

'MsgBox ("Data Modified")


End If


fillGrid
cmdEdit_4.Enabled = True
cmdDelete_3.Enabled = True
Call cmdAdd_1_Click
Exit Sub

save:

CON.RollbackTrans
If err.Number = "-2147217900" Then
   MsgBox "Duplicate Data ...", vbCritical
   txtCode.SetFocus
End If



End Sub

Private Sub Command1_Click()
Dim rs1 As New ADODB.Recordset

J = 1
ss = ""

If rs.State = 1 Then rs.Close
rs.Open "select count(bookno),bookno from copymaster2 group by bookno  having count(bookno)>1", CON
While rs.EOF = False


If Len(ss) > 0 Then
   CON.Execute "delete from copymaster2 where auto in(" & ss & ")"
End If


ss = ""
J = 1


If rs1.State = 1 Then rs1.Close
rs1.Open "select bookno,productQuality,Auto from copymaster2 where bookno='" & rs(1) & "' order by productQuality desc", CON
While rs1.EOF = False

If J > 1 Then

If ss = "" Then
ss = rs1(2)
Else
ss = ss & "," & rs1(2)
End If

End If


J = J + 1
rs1.MoveNext

Wend
    
rs.MoveNext
Wend

End Sub

Private Sub Form_Activate()
txtCode.SetFocus
fillGrid
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub Form_Load()
header(0).TOP = MainMenu.TOP + 60
header(0).Left = MainMenu.Left
header(0).Width = MainMenu.Width
cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = False


End Sub
Private Sub vs_DblClick()
   txtCode = vs.TextMatrix(vs.RowSel, 0)
   txtName = vs.TextMatrix(vs.RowSel, 1)
   cmdEdit_4.Enabled = True
   cmdDelete_3.Enabled = True
   cmdSave_2.Enabled = False
   
End Sub
