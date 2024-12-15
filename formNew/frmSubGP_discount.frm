VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmSubGP_discount 
   Caption         =   "Sub Group Discount"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9975
   Icon            =   "frmSubGP_discount.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8610
   ScaleWidth      =   9975
   Begin VB.Frame panel 
      Caption         =   "Sub Group Discount Category"
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
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   9885
      Begin VB.TextBox txtDisCode 
         Height          =   285
         Left            =   2415
         MaxLength       =   7
         TabIndex        =   12
         Top             =   630
         Width           =   2010
      End
      Begin VB.TextBox txtDisRate 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   2415
         TabIndex        =   11
         Top             =   1710
         Width           =   3135
      End
      Begin VB.ComboBox cboGPName 
         Height          =   315
         Left            =   2415
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   1335
         Width           =   3165
      End
      Begin VB.ComboBox cboGpCode 
         Height          =   315
         Left            =   2415
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   975
         Width           =   3165
      End
      Begin VB.Frame buttonFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   135
         TabIndex        =   1
         Top             =   2205
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
            Picture         =   "frmSubGP_discount.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   90
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
            Left            =   1365
            Picture         =   "frmSubGP_discount.frx":0BF0
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   90
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
            Left            =   2760
            Picture         =   "frmSubGP_discount.frx":17D4
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   90
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
            Left            =   4170
            Picture         =   "frmSubGP_discount.frx":23B8
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   90
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
            Left            =   5580
            Picture         =   "frmSubGP_discount.frx":27C5
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   90
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
            Picture         =   "frmSubGP_discount.frx":33A9
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
            Picture         =   "frmSubGP_discount.frx":3F8D
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1260
            Width           =   975
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   3840
         Left            =   90
         TabIndex        =   13
         Top             =   3330
         Width           =   9420
         _cx             =   16616
         _cy             =   6773
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
         BackColorFixed  =   12648447
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
      Begin VB.Label lblId 
         Height          =   330
         Left            =   4500
         TabIndex        =   18
         Top             =   630
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Image imgFile 
         Height          =   240
         Left            =   7245
         Picture         =   "frmSubGP_discount.frx":4B71
         Top             =   2610
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Rate"
         Height          =   255
         Left            =   165
         TabIndex        =   17
         Top             =   1800
         Width           =   2625
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Discount category code"
         Height          =   255
         Left            =   165
         TabIndex        =   16
         Top             =   630
         Width           =   2385
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Group Code"
         Height          =   255
         Left            =   165
         TabIndex        =   15
         Top             =   1020
         Width           =   2295
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Group Name"
         Height          =   255
         Left            =   165
         TabIndex        =   14
         Top             =   1395
         Width           =   2295
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   1005
         Left            =   90
         Top             =   2160
         Width           =   7035
      End
   End
End
Attribute VB_Name = "frmSubGP_discount"
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

Private Sub cmdAdd_1_Click()




cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False

txtDisCode = ""
txtDisRate = ""
cboGpCode = ""
cboGPName = ""
lblId.Caption = ""

edit1 = False
fillGrid

txtDisCode.SetFocus


   
End Sub
Sub fillGrid()

Dim f_grid As New ADODB.Recordset

SQL = "SELECT DISCCATS_Sub.AutoId,DISCCATS_Sub.categorycode, DISCCATS_Sub.groupcode, GROUPS.groupname," & _
" DISCCATS_Sub.discountrate FROM GROUPS LEFT JOIN DISCCATS_Sub ON GROUPS.groupcode = DISCCATS_Sub.groupcode where (DISCCATS_Sub.fyear='" & session & "' and DISCCATS_Sub.setupid=" & setupid & ") order by Categorycode"

If f_grid.State = 1 Then f_grid.close
f_grid.Open SQL, con, adOpenForwardOnly, adLockReadOnly

Set vs.DataSource = f_grid

'vs.ColWidth(1) = 800

For i = 1 To vs.Rows - 1
vs.Cell(flexcpPicture, i, 1) = imgFile
Next


End Sub
Private Sub cmdDelete_3_Click()

If MsgBox("want To delete ?", vbInformation + vbYesNo) = vbYes Then

 con.BeginTrans
 con.Execute "delete from DISCCATS_Sub where AutoId=" & lblId & " and " & stringyear
 con.CommitTrans
 
End If

cmdAdd_1_Click
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


If txtDisCode = "" Then
   MsgBox "Enter Enter Code. ...", vbInformation
   txtDisCode.SetFocus
   Exit Sub
End If
   
   
If cboGpCode = "" Then
   MsgBox "Select Group Code. ...", vbInformation
   cboGpCode.SetFocus
   Exit Sub
End If
   
If cboGPName = "" Then
   MsgBox "Select Group Name. ...", vbInformation
   cboGPName.SetFocus
   Exit Sub
End If
   



If edit1 = True Then
   con.Execute "update DISCCATS_Sub set categorycode='" & UCase(txtDisCode) & "',groupcode='" & cboGpCode & "'," & _
   " discountrate=" & Val(txtDisRate) & " where " & stringyear & " and AutoId=" & lblId & ""
   MsgBox "Data Updated....", vbInformation
Else


con.BeginTrans
con.Execute "INSERT INTO  [DISCCATS_Sub]" & _
           "([categorycode]" & _
           ",[groupcode]" & _
           ",[discountrate]" & _
           ",[Fyear]" & _
           ",[setupid]" & _
     ") Values" & _
           "('" & UCase(txtDisCode) & "'" & _
           ",'" & cboGpCode & "'" & _
           "" & _
           "," & Val(txtDisRate) & "" & _
           ",'" & main.session & "'" & _
           ",'" & main.setupid & "')"
con.CommitTrans

MsgBox "Data Saved....", vbInformation
End If




cmdSave_2.Enabled = False

Call cmdAdd_1_Click
   


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
cmdAdd_1_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{tab}"
End Sub
Private Sub Form_Load()
 
 'Me.TOP = 1800
 'Me.Left = 1500
 
 Me.Width = 10095
 Me.Height = 9200
 
 BackColorFrom Me
 
 fillcombo cboGpCode, "groupcode", "groups", con
 fillcombo cboGPName, "groupname", "groups", con
 
 fillGrid
 
 'cboGpCode.ListIndex = -1
 'cboGPName.ListIndex = -1
 
 
 
End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.Top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub txtDisCode_GotFocus()

If PopUpValue1 <> "" Then
    txtDisCode = PopUpValue1
    cboGpCode.Text = PopUpValue2
    txtDisRate = PopUpValue3
    
   If RS.State = 1 Then RS.close
   RS.Open "select groupname from groups where " & stringyear & " and groupcode='" & Trim(cboGpCode.Text) & "'", con
   If RS.EOF = False Then
      Me.cboGPName.Text = RS(0)
   End If
   
    
    
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    
End If

End Sub

Private Sub txtDisCode_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
   'SELECT Categorycode,Groupcode,Discountrate FROM [chitraData].[dbo].[DISCCATS_Sub]
    'popuplist "select SELECT Categorycode,Groupcode,Discountrate FROM DISCCATS_Sub where " & stringyear & " order by Categorycode", CON
    value = "SELECT Categorycode,Groupcode,Discountrate FROM DISCCATS_Sub where " & stringyear & " order by Categorycode"
   popuplistModel10 value, con

   
End If

End Sub

Private Sub txtDisCode_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{tab}"
End Sub
Private Sub vs_Click()
  
  cmdEdit_4.Enabled = True
  cmdSave_2.Enabled = False
  
  lblId.Caption = vs.TextMatrix(vs.RowSel, 1)
  txtDisCode = vs.TextMatrix(vs.RowSel, 2)
  cboGpCode = vs.TextMatrix(vs.RowSel, 3)
  cboGPName = vs.TextMatrix(vs.RowSel, 4)
  txtDisRate = vs.TextMatrix(vs.RowSel, 5)


End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then vs_Click
End Sub


