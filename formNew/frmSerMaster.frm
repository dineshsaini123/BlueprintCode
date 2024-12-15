VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmSerMaster 
   Caption         =   "Series Master"
   ClientHeight    =   7632
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   7680
   Icon            =   "frmSerMaster.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7632
   ScaleWidth      =   7680
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1350
      MaxLength       =   50
      TabIndex        =   9
      Top             =   225
      Width           =   5055
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   840
      Left            =   270
      TabIndex        =   2
      Top             =   6390
      Width           =   6585
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "S&earch"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   4395
         Picture         =   "frmSerMaster.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   30
         Width           =   1065
      End
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   5490
         Picture         =   "frmSerMaster.frx":0BF0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   45
         Width           =   1065
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   3315
         Picture         =   "frmSerMaster.frx":17D4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   30
         Width           =   1065
      End
      Begin VB.CommandButton cmdDelete_3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2220
         Picture         =   "frmSerMaster.frx":1BE1
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   45
         Width           =   1065
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1140
         Picture         =   "frmSerMaster.frx":27C5
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   30
         Width           =   1065
      End
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   45
         Picture         =   "frmSerMaster.frx":33A9
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   45
         Width           =   1065
      End
   End
   Begin VB.ComboBox cboYes_No 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      ItemData        =   "frmSerMaster.frx":3F8D
      Left            =   1350
      List            =   "frmSerMaster.frx":3F97
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   630
      Width           =   2490
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   4935
      Left            =   180
      TabIndex        =   10
      Top             =   1080
      Width           =   7170
      _cx             =   12647
      _cy             =   8705
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
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
      ForeColorSel    =   4210752
      BackColorBkg    =   16777215
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
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
      ExplorerBar     =   1
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Yes/No :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   12
      Top             =   630
      Width           =   1515
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   225
      Top             =   6345
      Width           =   6720
   End
   Begin VB.Label lblName1 
      Height          =   285
      Left            =   6480
      TabIndex        =   11
      Top             =   270
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Series Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   300
      Width           =   1515
   End
End
Attribute VB_Name = "frmSerMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit1 As Boolean
Private Sub cmdAdd_1_Click()



txtName = ""
lblName1 = ""
cboYes_No.ListIndex = -1

vs.Clear

cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False
edit1 = False


fillVs
txtName.SetFocus

   
   
End Sub
Sub fillVs()

vs.Cols = 3

k1 = 0
vs.Clear
vs.rows = 2

Set rs1 = New ADODB.Recordset
rs1.Open "select Id,Series,Yes_No from SeriesMaster order by Series", con
While rs1.EOF = False

vs.rows = vs.rows + 1

vs.TextMatrix(k1, 0) = rs1(0)
vs.TextMatrix(k1, 1) = rs1(1)
vs.TextMatrix(k1, 2) = rs1(2)


k1 = k1 + 1
rs1.MoveNext
Wend


vs.FormatString = "Id|Series Name|Yes/No"
vs.ColWidth(0) = 800
vs.ColWidth(1) = 4000
vs.ColWidth(2) = 1200


End Sub
Private Sub cmdDelete_3_Click()

'If MsgBox("want To delete ?", vbInformation + vbYesNo) = vbYes Then
'
' con.Execute "delete from  SeriesMaster where id='" & lblName1.Caption & "'"
'
'
'End If
'
'cmdAdd_1_Click

End Sub

Private Sub cmdEdit_4_Click()

edit1 = True
cmdEdit_4.Enabled = False
cmdSave_2.Enabled = True
cmdDelete_3.Enabled = True
cmdSave_2.SetFocus

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub



Private Sub cmdPrint_Click()

End Sub

Private Sub cmdSave_2_Click()

On Error GoTo aa:


If txtName = "" Then
   MsgBox "Enter Series Name. ...", vbInformation
   txtName.SetFocus
   Exit Sub
End If

If MsgBox("Want to Save ?", vbQuestion + vbYesNo) = vbNo Then
   Exit Sub
End If
   




If edit1 = True Then
   
   con.Execute "update [SeriesMaster] set Series='" & UCase(txtName) & "',yes_No='" & cboYes_No.text & "' where id='" & lblName1.Caption & "'"

Else

con.Execute "INSERT INTO  [SeriesMaster]" & _
           "(Series,[yes_No]" & _
     ") Values" & _
           "('" & UCase(txtName) & "'" & _
           ",'" & cboYes_No.text & "')"
        
        

End If

Screen.MousePointer = vbDefault


cmdSave_2.Enabled = False
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False

fillVs

txtName.SetFocus

   
Exit Sub

aa:

Screen.MousePointer = vbDefault
MsgBox "" & err.DESCRIPTION

End Sub

Private Sub cmdSearch_Click()

'   value = "select Series,Id from [SeriesMaster] order by Series"
'   popuplistModel10 value, con
'
'   cmdSave_2.Enabled = False
'   cmdEdit_4.Enabled = True

End Sub
Private Sub cmdSearch_GotFocus()
  

If PopUpValue1 <> "" Then
   
   
   
   
   txtName.text = PopUpValue1
   lblName1.Caption = PopUpValue2
    
   
End If
 
PopUpValue1 = ""

cmdEdit_4.Enabled = True
  

End Sub


Private Sub Form_Activate()
cmdAdd_1_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then sendkeys "{tab}"
End Sub

Private Sub Form_Load()
 Me.top = 500
 Me.Left = 500
 
 BackColorFrom Me
 
 fillVs
 
End Sub
Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      cmdSave_2.SetFocus
   End If
End Sub
Private Sub txtName_LostFocus()
   txtName = UCase(txtName)
End Sub
Private Sub vs_DblClick()

If Len(vs.TextMatrix(vs.RowSel, 2)) > 0 Then

lblName1.Caption = vs.TextMatrix(vs.RowSel, 0)
txtName.text = vs.TextMatrix(vs.RowSel, 1)
cboYes_No.text = vs.TextMatrix(vs.RowSel, 2)

End If

cmdEdit_4.Enabled = True
cmdDelete_3.Enabled = False
cmdSave_2.Enabled = False

End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = 115 Then
'     If MsgBox("Want to Delete ?", vbQuestion + vbYesNo) = vbYes Then
'        con.Execute "delete from SeriesMaster  where id=" & vs.TextMatrix(vs.RowSel, 0) & ""
'     End If
'  End If
End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      If vs.Col = 2 Then
         sendkeys "{right}"
      ElseIf vs.Col = 3 Then
         sendkeys "{down}"
         vs.Col = 2
      End If
   End If
End Sub


