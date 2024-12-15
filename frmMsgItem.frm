VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form frmMsgItem 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7545
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   11775
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   675
      TabIndex        =   15
      Top             =   6075
      Width           =   5340
      Begin VB.CommandButton cmdexit 
         Caption         =   "&Exit"
         Height          =   465
         Left            =   3750
         TabIndex        =   17
         Top             =   225
         Width           =   1290
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete"
         Height          =   465
         Left            =   2475
         TabIndex        =   8
         Top             =   225
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   465
         Left            =   1200
         TabIndex        =   7
         Top             =   225
         Width           =   1215
      End
      Begin VB.CommandButton cmdref 
         Caption         =   "&Refresh"
         Height          =   465
         Left            =   75
         TabIndex        =   16
         Top             =   225
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   5775
      TabIndex        =   5
      Top             =   1950
      Width           =   690
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3465
      Left            =   525
      TabIndex        =   6
      Top             =   2400
      Width           =   5940
      _cx             =   10477
      _cy             =   6112
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmMsgItem.frx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
   Begin VB.TextBox txtqty 
      Height          =   315
      Left            =   4425
      TabIndex        =   4
      Top             =   1950
      Width           =   1290
   End
   Begin VB.TextBox txtFCode 
      Height          =   315
      Left            =   2550
      TabIndex        =   3
      Top             =   1950
      Width           =   1665
   End
   Begin VB.TextBox txtRawCode 
      Height          =   315
      Left            =   630
      TabIndex        =   2
      Top             =   1950
      Width           =   1740
   End
   Begin MSComCtl2.DTPicker pdates 
      Height          =   315
      Left            =   2100
      TabIndex        =   1
      Top             =   1050
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20119553
      CurrentDate     =   39578
   End
   Begin VB.TextBox txtPno 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   1050
      Width           =   1440
   End
   Begin VB.Label Label7 
      Caption         =   "For Delete Row From Grid Press Delete"
      Height          =   240
      Left            =   675
      TabIndex        =   18
      Top             =   7125
      Width           =   3765
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Quantity"
      Height          =   240
      Left            =   4500
      TabIndex        =   14
      Top             =   1650
      Width           =   1140
   End
   Begin VB.Label Label5 
      Caption         =   "Finish Code"
      Height          =   240
      Left            =   2625
      TabIndex        =   13
      Top             =   1650
      Width           =   1140
   End
   Begin VB.Label Label4 
      Caption         =   "Raw Code"
      Height          =   240
      Left            =   600
      TabIndex        =   12
      Top             =   1650
      Width           =   1140
   End
   Begin VB.Label Label3 
      Caption         =   "Date"
      Height          =   165
      Left            =   2175
      TabIndex        =   11
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Production No"
      Height          =   240
      Left            =   600
      TabIndex        =   10
      Top             =   825
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Manufacture Form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   525
      TabIndex        =   9
      Top             =   225
      Width           =   2490
   End
End
Attribute VB_Name = "frmMsgItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim K As Integer
Dim b As Boolean
Sub setwidth()
vs.FormatString = "Raw Code|Finish Code|Finish Qty"
vs.ColWidth(0) = 2000
vs.ColWidth(1) = 2000
vs.ColWidth(2) = 2000
vs.Rows = 2
End Sub
Private Sub cmdAdd_Click()

If Me.txtRawCode.Text = "" And Me.txtFCode.Text = "" And Me.txtqty.Text = "" Then
MsgBox "Plz. Enter All The Entry !!", vbInformation
Me.txtPno.SetFocus
Exit Sub
End If

If b = False Then
K = 1
Else
K = K + 1
vs.Rows = vs.Rows + 1
End If

vs.TextMatrix(K, 0) = Me.txtRawCode.Text
vs.TextMatrix(K, 1) = Me.txtFCode.Text
vs.TextMatrix(K, 2) = Me.txtqty.Text
b = True
Me.txtRawCode.Text = ""
Me.txtFCode.Text = ""
Me.txtqty.Text = ""
Me.txtRawCode.SetFocus


End Sub

Private Sub cmdDel_Click()
If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
   CON.Execute "delete from MfgTable where pno='" & Me.txtPno.Text & "' and " & stridnyear
   cmdref_Click
End If
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdref_Click()
vs.Clear
setwidth

txtPno.Text = ""
txtRawCode.Text = ""
txtFCode.Text = ""
txtqty.Text = ""
txtPno.SetFocus


End Sub
Private Sub CmdSave_Click()
If MsgBox("Want To save ?", vbQuestion + vbYesNo) = vbYes Then
 saveData
End If
End Sub
Sub saveData()
If rs.State = 1 Then rs.Close
rs.Open "select * from MfgTable where Pno='" & Me.txtPno.Text & "' and " & stridnyear, CON
If rs.EOF = False Then
CON.Execute "delete from MfgTable where Pno='" & Me.txtPno.Text & "' and " & stridnyear
End If

If rs.State = 1 Then rs.Close
rs.Open "select * from MfgTable where " & stridnyear, CON, adOpenDynamic, adLockOptimistic
For I = 1 To vs.Rows - 1
If vs.TextMatrix(I, 0) <> "" Then
rs.AddNew
rs!Pno = txtPno.Text
rs!Dates = Me.pdates.Value
rs!RCode = vs.TextMatrix(I, 0)
rs!FCode = vs.TextMatrix(I, 1)
rs!Qty = vs.TextMatrix(I, 2)
rs!createdby = main.username
rs!createdon = Now
rs!fyear = main.session: rs!setupid = main.setupid
rs.Update
End If
Next
vs.Clear
txtPno.Text = ""
Me.txtRawCode.Text = ""
Me.txtFCode.Text = ""
txtqty.Text = ""
txtPno.SetFocus


End Sub
Sub SearchData()

If rs.State = 1 Then rs.Close
rs.Open "select * from MfgTable where Pno='" & Me.txtPno.Text & "' and " & stridnyear, CON
If rs.EOF = False Then

txtPno.Text = rs!Pno
Me.pdates.Value = rs!Dates
For I = 1 To rs.RecordCount
vs.Rows = vs.Rows + 1
vs.TextMatrix(I, 0) = rs!RCode
vs.TextMatrix(I, 1) = rs!FCode
vs.TextMatrix(I, 2) = rs!Qty
rs.MoveNext
Next
End If


End Sub

Private Sub Form_Load()
setwidth
Me.Top = 50
Me.Left = 50
Me.pdates.Value = Date
End Sub
Private Sub pdates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Me.txtRawCode.SetFocus
End Sub

Private Sub txtFCode_GotFocus()
If PopUpValue1 <> "" Then
Me.txtFCode.Text = PopUpValue1
PopUpValue1 = ""
End If
End Sub
Private Sub txtFCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtqty.SetFocus
End If

If KeyCode = 113 Then
   popuplist10 "select BOOKCODE,BOOKNAME + ': ' + convert(varchar,size1) + ' ' + unit1 + ' ' + convert(varchar,size2) + ' ' + unit2 + ': ' + quality as Item from books where upper(ltrim(rtrim(GROUPCODE)))=upper(ltrim(rtrim('Yes'))) and " & stridnyear & " ", CON
End If
End Sub
Private Sub txtPno_GotFocus()

If PopUpValue1 <> "" Then
Me.txtPno.Text = PopUpValue1
SearchData
PopUpValue1 = ""
End If

End Sub

Private Sub txtPno_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.pdates.SetFocus

If txtPno.Text <> "" Then
SearchData
End If

End If


If KeyCode = 113 Then
   popuplist10 "select Pno,Dates from mfgtable where  " & stridnyear & "   group by Pno,Dates", CON
End If

End Sub
Private Sub txtqty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Call cmdAdd_Click
End If
End Sub

Private Sub txtRawCode_GotFocus()
If PopUpValue1 <> "" Then
Me.txtRawCode.Text = PopUpValue1
PopUpValue1 = ""
End If
End Sub
Private Sub txtRawCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.txtFCode.SetFocus
End If

If KeyCode = 113 Then
   popuplist10 "select BOOKCODE as Code,BOOKNAME + ': ' + convert(varchar,size1) + ' ' + unit1 + ' ' + convert(varchar,size2) + ' ' + unit2 + ': ' + quality as Item from books where  " & stridnyear & "  and upper(ltrim(rtrim(GROUPCODE)))='NO'", CON
End If
   
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
vs.RemoveItem (vs.RowSel)
End If
End Sub
