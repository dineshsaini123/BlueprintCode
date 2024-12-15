VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmBookGpMaster 
   Caption         =   "Book Group "
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10725
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   10725
   WindowState     =   2  'Maximized
   Begin VB.Frame panel 
      Caption         =   "Book Group"
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
      Height          =   7530
      Left            =   270
      TabIndex        =   5
      Top             =   135
      Width           =   10320
      Begin VB.ComboBox cboCat 
         Height          =   315
         ItemData        =   "frmBookGpMaster.frx":0000
         Left            =   1800
         List            =   "frmBookGpMaster.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   495
         Width           =   1875
      End
      Begin VB.TextBox txtGPName 
         Height          =   285
         Left            =   1785
         MaxLength       =   49
         TabIndex        =   2
         Top             =   1215
         Width           =   3135
      End
      Begin VB.TextBox txtGpCode 
         Height          =   285
         Left            =   1785
         MaxLength       =   7
         TabIndex        =   1
         Top             =   855
         Width           =   3135
      End
      Begin VB.Frame buttonFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   135
         TabIndex        =   6
         Top             =   1710
         Width           =   7935
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
            Height          =   795
            Left            =   45
            Picture         =   "frmBookGpMaster.frx":001D
            Style           =   1  'Graphical
            TabIndex        =   10
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
            Height          =   795
            Left            =   1290
            Picture         =   "frmBookGpMaster.frx":0C01
            Style           =   1  'Graphical
            TabIndex        =   3
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
            Height          =   795
            Left            =   2520
            Picture         =   "frmBookGpMaster.frx":17E5
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   135
            Width           =   1275
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
            Height          =   795
            Left            =   3795
            Picture         =   "frmBookGpMaster.frx":23C9
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   135
            Width           =   1275
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
            Height          =   795
            Left            =   6465
            Picture         =   "frmBookGpMaster.frx":27D6
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   120
            Width           =   1410
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
            Height          =   795
            Left            =   5085
            Picture         =   "frmBookGpMaster.frx":33BA
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   135
            Width           =   1365
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   4335
         Left            =   90
         TabIndex        =   11
         Top             =   2940
         Width           =   10005
         _cx             =   17648
         _cy             =   7646
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
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   1140
         Left            =   135
         Top             =   1665
         Width           =   7980
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         Height          =   255
         Left            =   135
         TabIndex        =   15
         Top             =   540
         Width           =   1875
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Group Name"
         Height          =   285
         Left            =   135
         TabIndex        =   14
         Top             =   1215
         Width           =   1845
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Group Code"
         Height          =   255
         Left            =   135
         TabIndex        =   13
         Top             =   855
         Width           =   1830
      End
      Begin VB.Label lblId 
         Height          =   285
         Left            =   3735
         TabIndex        =   12
         Top             =   630
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Image imgFile 
         Height          =   240
         Left            =   8280
         Picture         =   "frmBookGpMaster.frx":3F9E
         Top             =   2385
         Visible         =   0   'False
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmBookGpMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit1 As Boolean

Private Sub cboCat_LostFocus()
fillGrid
End Sub

Private Sub cmdAdd_1_Click()
lblId.Caption = ""
'cboCat.ListIndex = -1
txtGpCode = ""
txtGPName = ""

edit1 = False

cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False

fillGrid

cboCat.SetFocus

'txtName.SetFocus

   
   
End Sub

Private Sub cmdDelete_3_Click()

If MsgBox("want To delete ?", vbInformation + vbYesNo) = vbYes Then

 con.BeginTrans
 con.Execute "delete from  GROUPS where AutoId=" & lblId & " and " & stringyear
 con.CommitTrans
 
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

On Error GoTo save_:


If txtGpCode = "" Then
   MsgBox "Enter Group Code. ...", vbInformation
   txtGpCode.SetFocus
   Exit Sub
End If
   
   
If txtGPName = "" Then
   MsgBox "Enter Group Name. ...", vbInformation
   txtGPName.SetFocus
   Exit Sub
End If
   
If cboCat = "" Then
   MsgBox "Select Category. ...", vbInformation
   cboCat.SetFocus
   Exit Sub
End If




If edit1 = True Then
   
   con.Execute "update  [GROUPS] set groupcode='" & txtGpCode & "',groupname='" & txtGPName & "',Category='" & Trim(cboCat) & "'" & _
   " where " & stringyear & " and AutoId=" & lblId.Caption & ""
 
 
Else


   
con.BeginTrans
con.Execute "INSERT INTO  [GROUPS]" & _
           "(groupcode" & _
           ",groupname" & _
           ",Category" & _
           ",fyear" & _
           ",[setupid]" & _
     ") Values" & _
           "('" & UCase(txtGpCode) & "'" & _
           ",'" & UCase(txtGPName) & "'" & _
           ",'" & UCase(cboCat) & "'" & _
           ",'" & main.session & "'" & _
           ",'" & main.setupid & "')"
con.CommitTrans

End If


MsgBox "Date Saved ....", vbInformation
cmdSave_2.Enabled = False

Call cmdAdd_1_Click
   

Exit Sub
save_:
MsgBox "" & err.DESCRIPTION

End Sub
Private Sub cmdSearch_GotFocus()
  
  
  If PopUpValue1 <> "" Then
  
  txtName = PopUpValue1
  
  
  End If
   
  PopUpValue1 = ""
  PopUpValue2 = ""
  
  cmdEdit_4.Enabled = True
  

End Sub
Sub fillGrid()

Dim f_grid As New ADODB.Recordset

If f_grid.State = 1 Then f_grid.close
If cboCat.Text = "" Then
f_grid.Open "select AutoId,Category,Groupcode,Groupname from GROUPS where " & stringyear & " order by Groupcode", con, adOpenDynamic, adLockOptimistic
Else
f_grid.Open "select AutoId,Category,Groupcode,Groupname from GROUPS where category ='" & cboCat & "' and  " & stringyear & " order by Groupcode", con, adOpenDynamic, adLockOptimistic

End If
Set vs.DataSource = f_grid

vs.ColWidth(1) = 1000

For I = 1 To vs.Rows - 1
vs.Cell(flexcpPicture, I, 1) = imgFile
Next


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
 
 fillGrid
 
 BackColorFrom Me
 
End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.Top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub vs_Click()
  cmdSave_2.Enabled = False
  cmdEdit_4.Enabled = True
  
  lblId.Caption = vs.TextMatrix(vs.RowSel, 1)
  cboCat.Text = vs.TextMatrix(vs.RowSel, 2)
  txtGpCode = vs.TextMatrix(vs.RowSel, 3)
  txtGPName = vs.TextMatrix(vs.RowSel, 4)
  
  
  

End Sub

