VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmCityMaster 
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   9885
   Begin VB.Frame panel 
      Caption         =   "City Master"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   7485
      Left            =   120
      TabIndex        =   0
      Top             =   135
      Width           =   9645
      Begin VB.TextBox txtCityName 
         Height          =   285
         Left            =   2415
         MaxLength       =   25
         TabIndex        =   10
         Top             =   720
         Width           =   3315
      End
      Begin VB.ComboBox cboDisName 
         Height          =   315
         Left            =   2415
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   1065
         Width           =   3390
      End
      Begin VB.Frame buttonFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   135
         TabIndex        =   1
         Top             =   2205
         Width           =   6765
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
            Picture         =   "frmCityMaster.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   45
            Width           =   1245
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
            Left            =   1320
            Picture         =   "frmCityMaster.frx":0BE4
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   45
            Width           =   1290
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
            Left            =   2670
            Picture         =   "frmCityMaster.frx":17C8
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   45
            Width           =   1290
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
            Left            =   4035
            Picture         =   "frmCityMaster.frx":23AC
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   45
            Width           =   1290
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
            Left            =   5400
            Picture         =   "frmCityMaster.frx":27B9
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   45
            Width           =   1290
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
            Picture         =   "frmCityMaster.frx":339D
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1260
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
            Picture         =   "frmCityMaster.frx":3F81
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1260
            Width           =   975
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   4020
         Left            =   90
         TabIndex        =   11
         Top             =   3330
         Width           =   9420
         _cx             =   16616
         _cy             =   7091
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   11162880
         BackColorFixed  =   7917545
         ForeColorFixed  =   8388608
         BackColorSel    =   16777153
         ForeColorSel    =   -2147483647
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         ExplorerBar     =   0
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
      Begin VB.Label cityId 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5760
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblId 
         Height          =   330
         Left            =   4500
         TabIndex        =   14
         Top             =   630
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Image imgFile 
         Height          =   240
         Left            =   7245
         Picture         =   "frmCityMaster.frx":4B65
         Top             =   2610
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "City Name :"
         Height          =   255
         Left            =   165
         TabIndex        =   13
         Top             =   720
         Width           =   2385
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "District Name"
         Height          =   255
         Left            =   165
         TabIndex        =   12
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   960
         Left            =   90
         Top             =   2160
         Width           =   6855
      End
   End
End
Attribute VB_Name = "frmCityMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit1 As Boolean
Private Sub cboGpCode_Click()
  
   If cboGpCode = "" Then Exit Sub
   
   If RS.State = 1 Then RS.close
   RS.Open "select groupname from groups where " & stringyear & " and groupcode='" & Trim(cboGpCode) & "'", con
   If RS.EOF = False Then
      cboGPName.Text = RS(0)
   End If
   
End Sub
Private Sub cboGpCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cboGPName_Click()
   If cboGPName = "" Then Exit Sub
   
   If RS.State = 1 Then RS.close
   RS.Open "select groupcode from groups where " & stringyear & " and groupname='" & Trim(cboGPName) & "'", con
   If RS.EOF = False Then
      cboGpCode.Text = RS(0)
   End If

End Sub
Private Sub cboGPName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub


Private Sub cboDisName_Click()
fillGrid
End Sub

Private Sub cmdAdd_1_Click()




cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False

txtCityName.Text = ""
cboDisName.Text = ""
cityId.Caption = ""



edit1 = False
'fillGrid
txtCityName.SetFocus


   
End Sub
Sub fillGrid()

Screen.MousePointer = vbHourglass

Dim f_grid As New ADODB.Recordset
If cboDisName <> "" Then

SQL = "SELECT City.CityId,City,District " & _
" FROM City where (City.fyear='" & session & "' and City.setupid=" & setupid & " and District='" & cboDisName & "') order by City"

Else

SQL = "SELECT City.CityId,City,District " & _
" FROM City where (City.fyear='" & session & "' and City.setupid=" & setupid & ") order by City"



End If

If f_grid.State = 1 Then f_grid.close
f_grid.Open SQL, con, adOpenDynamic, adLockOptimistic

Set vs.DataSource = f_grid


Screen.MousePointer = vbDefault


End Sub
Private Sub cmdDelete_3_Click()

If MsgBox("want To delete ?", vbInformation + vbYesNo) = vbYes Then

 con.BeginTrans
 con.Execute "delete from City where CityID=" & cityId & " and " & stringyear
 con.CommitTrans
 
 vs.RemoveItem (vs.RowSel)
 
 txtCityName = ""
 cboDisName = ""
 
End If

End Sub

Private Sub cmdEdit_4_Click()

cmdEdit_4.Enabled = False
cmdSave_2.Enabled = True
cmdDelete_3.Enabled = True
edit1 = True
cmdSave_2.SetFocus

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub



Private Sub cmdSave_2_Click()



   
If txtCityName = "" Then
   MsgBox "Enter City Name. ...", vbInformation
   txtCityName.SetFocus
   Exit Sub
End If
   
If cboDisName = "" Then
   MsgBox "Select District Name. ...", vbInformation
   cboDisName.SetFocus
   Exit Sub
End If
   
   
   




If edit1 = True Then
   con.Execute "update City set City='" & UCase(txtCityName) & "',District='" & cboDisName & "' where " & stringyear & " and CityID=" & cityId & ""
   MsgBox "Data Updated....", vbInformation
Else


con.BeginTrans
con.Execute "INSERT INTO  [City]" & _
           "([City]" & _
           ",[District]" & _
           "" & _
           ",[Fyear]" & _
           ",[setupid]" & _
     ") Values" & _
           "('" & UCase(txtCityName) & "'" & _
           ",'" & cboDisName & "'" & _
           "" & _
           "" & _
           ",'" & main.session & "'" & _
           ",'" & main.setupid & "')"
con.CommitTrans

MsgBox "Data Saved....", vbInformation
End If


txtCityName = ""
cboDisName.Text = ""


'cmdSave_2.Enabled = False

'Call cmdAdd_1_Click
   


End Sub

Private Sub cmdSearch_Click()

   popuplist10 "select DISTRICTNAME,AGENTNAME from [DISTRICTS] where  " & stringyear & " order by DISTRICTNAME", con
 
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
'cmdAdd_1_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub
Private Sub Form_Load()
'BackColorFrom Me
Me.Left = 100
Me.Top = 100
Me.Width = 10000
Me.Height = 8190

Set RS = con.Execute("exec DistQry '" & session & "'," & main.setupid & "")
While RS.EOF = False
cboDisName.AddItem RS(0)
RS.MoveNext
Wend
'fillcombo cboDisName, "DISTRICTNAME", "DISTRICTS", CON
If cboDisName.ListIndex > 0 Then
   cboDisName.ListIndex = 1
End If
fillGrid
'Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.Top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub txtDisCode_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub
Private Sub vs_Click()
  
  cmdEdit_4.Enabled = True
  cmdSave_2.Enabled = False
  
  cityId.Caption = vs.TextMatrix(vs.RowSel, 1)
  txtCityName.Text = vs.TextMatrix(vs.RowSel, 2)
  cboDisName.Text = vs.TextMatrix(vs.RowSel, 3)
  

End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then vs_Click
End Sub


