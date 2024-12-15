VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmAddPaperOp 
   Caption         =   "Paper Opening"
   ClientHeight    =   7104
   ClientLeft      =   60
   ClientTop       =   396
   ClientWidth     =   6852
   Icon            =   "frmAddPaperOp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7104
   ScaleWidth      =   6852
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPlanNo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   45
      TabIndex        =   5
      Top             =   135
      Width           =   600
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
      Left            =   1170
      Picture         =   "frmAddPaperOp.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6255
      Width           =   1065
   End
   Begin VB.CommandButton Commandsave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sa&ve"
      Height          =   735
      Left            =   45
      Picture         =   "frmAddPaperOp.frx":0419
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6210
      Width           =   1080
   End
   Begin VB.TextBox txtPrinter 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   675
      TabIndex        =   2
      Top             =   135
      Width           =   6090
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   5520
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   6780
      _cx             =   11959
      _cy             =   9737
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   12582847
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483645
      BackColorAlternate=   -2147483643
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
      Cols            =   3
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
      Editable        =   2
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
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   14640
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   6360
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmAddPaperOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_f As New ADODB.Recordset
Private Sub cmdEdit_4_Click()
Unload Me
End Sub

Private Sub Commandsave_Click()

For k1 = 1 To vs.rows - 1

If vs.TextMatrix(k1, 0) <> "" Then
   con.Execute "update PrinterWisePaperOp set OP_Ream=" & vs.TextMatrix(k1, 1) & ",OP_sheet=" & vs.TextMatrix(k1, 2) & "  where Paper='" & vs.TextMatrix(k1, 0) & "' and PlanNo=" & txtPlanNo.text & " and Printer='" & txtPrinter & "'"
End If


Next

End Sub

Private Sub Form_Load()

Dim k1 As Integer

'txtPrinter.text = frmPaperPlan.binder_name.text
'txtPlanNo.text = frmPaperPlan.txtOrdNo.text

vs.rows = 2

k1 = 1

If rs_f.State = 1 Then rs_f.close
rs_f.Open "select Paper,OP_Ream,OP_Sheet from PrinterWisePaperOp where PlanNo=" & txtPlanNo.text & " and Printer='" & txtPrinter & "'", con
While rs_f.EOF = False

vs.TextMatrix(k1, 0) = rs_f(0)
vs.TextMatrix(k1, 1) = rs_f(1) & ""
vs.TextMatrix(k1, 2) = rs_f(2) & ""
k1 = k1 + 1
vs.rows = vs.rows + 1

rs_f.MoveNext

Wend


vs.FormatString = "Paper|OP Ream|OP Sheet"

vs.ColWidth(0) = 4000
vs.ColWidth(1) = 1200
vs.ColWidth(2) = 1200





End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   
   If vs.Col = 1 Then
      sendkeys "{right}"
   ElseIf vs.Col = 2 Then
      sendkeys "{home}"
      sendkeys "{down}"
   End If
   

End If



End Sub
