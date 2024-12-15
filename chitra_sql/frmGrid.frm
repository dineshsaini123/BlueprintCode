VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmGrid 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboGrid 
      Height          =   315
      Left            =   1485
      TabIndex        =   14
      Top             =   765
      Width           =   1635
   End
   Begin VB.ComboBox cboUser 
      Height          =   315
      Left            =   1485
      TabIndex        =   12
      Top             =   360
      Width           =   1635
   End
   Begin VB.TextBox txtTop 
      Height          =   330
      Left            =   1485
      TabIndex        =   10
      Top             =   2385
      Width           =   960
   End
   Begin VB.TextBox txtLeft 
      Height          =   330
      Left            =   1485
      TabIndex        =   8
      Top             =   1980
      Width           =   960
   End
   Begin VB.TextBox txtHieght 
      Height          =   330
      Left            =   1485
      TabIndex        =   6
      Top             =   1575
      Width           =   960
   End
   Begin VB.TextBox txtWidth 
      Height          =   330
      Left            =   1485
      TabIndex        =   4
      Top             =   1170
      Width           =   960
   End
   Begin VB.TextBox txtNoCols 
      Height          =   330
      Left            =   1485
      TabIndex        =   3
      Top             =   2880
      Width           =   960
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Setting"
      Height          =   600
      Left            =   540
      TabIndex        =   1
      Top             =   3870
      Width           =   1725
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   4065
      Left            =   4230
      TabIndex        =   0
      Top             =   360
      Width           =   6810
      _cx             =   12012
      _cy             =   7170
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
      BackColorBkg    =   -2147483636
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
      AllowUserResizing=   4
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Grid Name :"
      Height          =   285
      Left            =   495
      TabIndex        =   15
      Top             =   765
      Width           =   1005
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "User :"
      Height          =   285
      Left            =   495
      TabIndex        =   13
      Top             =   360
      Width           =   1005
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Top :"
      Height          =   285
      Left            =   495
      TabIndex        =   11
      Top             =   2430
      Width           =   1005
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Left :"
      Height          =   285
      Left            =   495
      TabIndex        =   9
      Top             =   2025
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Hieght :"
      Height          =   285
      Left            =   495
      TabIndex        =   7
      Top             =   1620
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Width :"
      Height          =   285
      Left            =   495
      TabIndex        =   5
      Top             =   1215
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "No Of Cols :"
      Height          =   285
      Left            =   495
      TabIndex        =   2
      Top             =   2925
      Width           =   1005
   End
End
Attribute VB_Name = "frmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboGrid_Click()


If RS.State = 1 Then RS.close
'User='" & cboUser.Text & "' and
RS.Open "select * from grid_ini where (GridName='" & cboGrid.Text & "')", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
   
   txtNoCols = RS!GridCols
   txtWidth = RS!gridwidth
   txtHieght = RS!gridHieght
   txtTop = RS!gridTop
   txtLeft = RS!gridLeft
   
   txtWidth = RS!gridwidth
   txtHieght = RS!gridHieght
 
   
End If



Call ini_grid(vs, cboUser.Text)
  
End Sub

Private Sub CmdSave_Click()


Set RS = New ADODB.Recordset
'If rs.State = 1 Then rs.Close
RS.Open "select * from grid_ini", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
  For I = 0 To vs.Cols - 1
   RS("Colwidth_" & I) = vs.ColWidth(I)
   RS.update
  Next
  
  RS!gridTop = txtTop
  RS!gridLeft = txtLeft
  RS!gridwidth = txtWidth
  RS!gridHieght = txtHieght
  
  RS!gridminraw = vs.RowHeight(0)
  RS.update
  
End If


End Sub

Private Sub Form_Load()

If RS.State = 1 Then RS.close
RS.Open "select username FROM UsrePermission group by username", coninfo
While RS.EOF = False
 cboUser.AddItem RS(0)
 RS.MoveNext
Wend


If RS.State = 1 Then RS.close
RS.Open "select GridName FROM grid_ini group by GridName", con
While RS.EOF = False
 cboGrid.AddItem RS(0)
 RS.MoveNext
Wend



End Sub
Function ini_grid(vs_grid As VSFlexGrid, user_name As String)
  
  'GridName='" & grid_name & "' and
  
If RS.State = 1 Then RS.close
RS.Open "select * from grid_ini", con
If RS.EOF = False Then
   
  
  
    vs.Top = RS!gridTop
    vs.Left = RS!gridLeft
   
   
   vs.Height = RS!gridHieght
   vs.Width = RS!gridwidth
   
   
   If RS!WordWrap = True Then
      vs.WordWrap = True
   Else
      vs.WordWrap = False
   End If
   
   If Not IsNull(RS!gridminraw) Then
      vs.RowHeightMin = RS!gridminraw
   End If
   

   vs_grid.Cols = RS!GridCols
   For I = 0 To vs_grid.Cols - 1
       vs_grid.ColWidth(I) = RS.Fields("Colwidth_" & I) & ""
   Next
   
End If






End Function

