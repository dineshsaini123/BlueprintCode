VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmColor 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   2565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   540
      Left            =   120
      TabIndex        =   1
      Top             =   3915
      Width           =   2295
   End
   Begin VSFlex7LCtl.VSFlexGrid vs 
      Height          =   3750
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2295
      _cx             =   4048
      _cy             =   6615
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
      GridColorFixed  =   16761087
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
      Rows            =   6
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   600
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmColor.frx":0000
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
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdApply_Click()

    con.Execute "update color_setting set apply='" & False & "' where module='" & module_ & "'"
    For k1 = 0 To vs.Rows - 1
       If vs.TextMatrix(k1, 1) = "-1" Then
        con.Execute "update color_setting set apply='" & True & "' where name='" & vs.TextMatrix(k1, 2) & "' and module='" & module_ & "'"
       Else
        con.Execute "update color_setting set apply='" & False & "' where name='" & vs.TextMatrix(k1, 2) & "' and module='" & module_ & "'"
       End If
    Next
    MsgBox "Apply ....", vbInformation


End Sub
Private Sub Form_Load()

    Me.Top = 1000
    Me.Left = 2000
    vs.ColWidth(2) = 0
    
    If RS.State = 1 Then RS.close
    RS.Open "select color_dark,name,apply from color_setting where module='" & module_ & "'", con
    For k1 = 0 To vs.Rows - 1
       vs.Cell(flexcpBackColor, k1, 0) = RS(0)
       vs.TextMatrix(k1, 1) = RS!Apply
       vs.TextMatrix(k1, 2) = RS!Name
       RS.MoveNext
    Next


End Sub
