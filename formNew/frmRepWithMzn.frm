VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmRepWithMzn 
   Caption         =   "Rep. Add to Manager"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   12615
   Icon            =   "frmRepWithMzn.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   12615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   510
      Left            =   11475
      TabIndex        =   3
      Top             =   90
      Width           =   1005
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   510
      Left            =   9315
      TabIndex        =   2
      Top             =   90
      Width           =   1005
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   510
      Left            =   10395
      TabIndex        =   1
      Top             =   90
      Width           =   1005
   End
   Begin VSFlex7Ctl.VSFlexGrid vs1 
      Height          =   6555
      Left            =   45
      TabIndex        =   0
      Top             =   675
      Width           =   12435
      _cx             =   21934
      _cy             =   11562
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   11162880
      BackColorFixed  =   12640511
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   8404992
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
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRepWithMzn.frx":000C
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
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
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
   Begin VB.Label Label1_rows 
      Height          =   240
      Left            =   45
      TabIndex        =   4
      Top             =   7290
      Width           =   1320
   End
End
Attribute VB_Name = "frmRepWithMzn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str11 As String
Private Sub cmdAdd_Click()
  HeadTbl = "manager"
  frmMasters.Show 1
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdok_Click()

For I = 1 To vs1.Rows - 1

If (vs1.TextMatrix(I, 0) <> "" And vs1.TextMatrix(I, 2) <> "") Then
   CON_blue.Execute "update Rep set manager='" & vs1.TextMatrix(I, 2) & "' where headname='" & vs1.TextMatrix(I, 0) & "'"
End If

Next

MsgBox ("Data Updated....")

End Sub

Private Sub Form_Load()
  HeadTbl = "manager"
  
  AddItem
  
  fillrep
  
  Label1_rows.Caption = "Rows : " & vs1.Rows
  
End Sub
Sub AddItem()
 
str11 = ""
 
If RS.State = 1 Then RS.close
RS.Open "Select * from MasterTbl where category='" & HeadTbl & "' order by name", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   If str11 = "" Then
      str11 = RS!Name
   Else
      str11 = str11 & "|" & RS!Name
   End If
   
   RS.MoveNext
Wend

vs1.ColComboList(2) = str11
 
End Sub
Sub fillrep()
   
   Dim f As New ADODB.Recordset
   Dim k1 As Integer
   k1 = 1
   vs1.Rows = 1
   
   f.Open "select Headname,HeadEmail,Manager from [rep] where len(headname)>0 group by Headname,HeadEmail,Manager order by headname", CON_blue
   
   While f.EOF = False
   
   
   
   vs1.Rows = vs1.Rows + 1
   
   vs1.TextMatrix(k1, 0) = f!headName
   vs1.TextMatrix(k1, 1) = f!HeadEmail
    vs1.TextMatrix(k1, 2) = f!Manager & ""
   
   k1 = k1 + 1
   
   
   
   f.MoveNext
   Wend
   
   vs1.FormatString = "HeadName|HeadEmail|Manager"
   
   vs1.ColWidth(0) = 3500
   vs1.ColWidth(1) = 3500
   vs1.ColWidth(2) = 2500
   
   
   
End Sub
Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   CON_blue.Execute "update Rep set manager='" & vs1.TextMatrix(vs1.RowSel, 4) & "' where repid='" & vs1.TextMatrix(vs1.RowSel, 0) & "'"
   SendKeys "{down}"

End If

End Sub
Private Sub vs1_SelChange()

  '' CON_blue.Execute "update Rep set manager='" & vs1.TextMatrix(vs1.RowSel, 4) & "' where repid='" & vs1.TextMatrix(vs1.RowSel, 0) & "'"
  '' SendKeys "{down}"
   
End Sub
