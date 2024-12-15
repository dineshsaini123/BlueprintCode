VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmPaperWastage 
   Caption         =   "Wastage"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   6045
   Icon            =   "frmPaperWastage.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   6045
   Begin VB.TextBox txtQty2 
      Height          =   330
      Left            =   3330
      MaxLength       =   6
      TabIndex        =   1
      Top             =   135
      Width           =   825
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Height          =   630
      Left            =   1215
      Picture         =   "frmPaperWastage.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1125
      Width           =   1035
   End
   Begin VB.TextBox txtQty1 
      Height          =   330
      Left            =   2205
      MaxLength       =   6
      TabIndex        =   0
      Top             =   135
      Width           =   825
   End
   Begin VB.CommandButton cmdExit_12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   630
      Left            =   2340
      Picture         =   "frmPaperWastage.frx":0BF0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1125
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave_2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   630
      Left            =   90
      Picture         =   "frmPaperWastage.frx":17D4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1125
      Width           =   1035
   End
   Begin VB.TextBox txtWastage 
      Height          =   330
      Left            =   2205
      MaxLength       =   5
      TabIndex        =   2
      Top             =   540
      Width           =   735
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3750
      Left            =   90
      TabIndex        =   6
      Top             =   1890
      Width           =   5730
      _cx             =   10107
      _cy             =   6615
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3060
      TabIndex        =   9
      Top             =   180
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Printing Quantity  Range:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   135
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Wastage percentage (%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   540
      Width           =   2175
   End
End
Attribute VB_Name = "frmPaperWastage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDel_Click()

If MsgBox("Want to Delete ?", vbQuestion + vbYesNo) = vbYes Then
   con.Execute "delete from WastageQty where (Wastage='" & txtWastage.Text & "' and Qty1='" & txtQty1.Text & "' and Qty2='" & txtQty2.Text & "')"
   txtWastage.Text = ""
   txtQty1.Text = ""
   txtQty2.Text = ""
   
   txtWastage.SetFocus
   fillVs
End If

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub
Private Sub cmdSave_2_Click()

If RS.State = 1 Then RS.close
RS.Open "select * from WastageQty where (Wastage='" & txtWastage.Text & "' and Qty1='" & txtQty1.Text & "' and Qty2='" & txtQty2.Text & "')", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
   con.Execute "insert into WastageQty(Qty1,Qty2,Wastage) values('" & txtQty1.Text & "','" & txtQty2.Text & "','" & txtWastage.Text & "')"
Else
   If RS.State = 1 Then RS.close
   RS.Open "select * from WastageQty where (Wastage='" & txtWastage.Text & "' and Qty1='" & txtQty1.Text & "' and Qty2='" & txtQty2.Text & "')", con, adOpenDynamic, adLockOptimistic
   If RS.EOF = False Then
      RS!Qty1 = txtQty1.Text
      RS!Qty2 = txtQty2.Text
      RS!Wastage = txtWastage.Text
      RS.update
   End If
End If

    fillVs
    
End Sub
Sub fillVs()

Set rs1 = New ADODB.Recordset
rs1.Open "select Qty1,Qty2,Wastage from WastageQty order by Qty1,Qty2", con, adOpenDynamic, adLockOptimistic
Set vs.DataSource = rs1

vs.FormatString = "<Qty From|<Qty. To|^Wastage"
vs.ColWidth(0) = 2000
vs.ColWidth(1) = 2000
vs.ColWidth(2) = 2000

    
    
End Sub
Private Sub Form_Load()

Me.Width = 6165
Me.Height = 6400


fillVs


End Sub
Private Sub vs_Click()
   txtQty1.Text = vs.TextMatrix(vs.RowSel, 0)
   txtQty2.Text = vs.TextMatrix(vs.RowSel, 1)
   txtWastage.Text = vs.TextMatrix(vs.RowSel, 2)
End Sub
