VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReturnStatus 
   Caption         =   "Return Status"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   13650
   Icon            =   "frmReturnStatus.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   13650
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FAEFC9&
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8415
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   90
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   3960
      TabIndex        =   2
      Top             =   90
      Width           =   4335
      Begin VB.OptionButton Option_Clear 
         Caption         =   "Approved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1845
         TabIndex        =   5
         Top             =   180
         Width           =   1275
      End
      Begin VB.OptionButton Option_Pending 
         Caption         =   "Not Approved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         TabIndex        =   4
         Top             =   135
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.OptionButton Option_Both 
         Caption         =   "Both"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3285
         TabIndex        =   3
         Top             =   180
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FAEFC9&
      Caption         =   "&Update"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9675
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   90
      Width           =   1230
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FAEFC9&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10935
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   90
      Width           =   1230
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7455
      Left            =   45
      TabIndex        =   7
      Top             =   810
      Width           =   13530
      _cx             =   23865
      _cy             =   13150
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
      ForeColorSel    =   16711680
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
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   2
      RowHeightMin    =   580
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmReturnStatus.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
   Begin MSComCtl2.DTPicker txtFrom 
      Height          =   330
      Left            =   765
      TabIndex        =   8
      Top             =   180
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   582
      _Version        =   393216
      Format          =   79560705
      CurrentDate     =   39795
   End
   Begin MSComCtl2.DTPicker txtTo 
      Height          =   330
      Left            =   2475
      TabIndex        =   9
      Top             =   180
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   582
      _Version        =   393216
      Format          =   79560705
      CurrentDate     =   39795
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2160
      TabIndex        =   11
      Top             =   180
      Width           =   330
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   180
      TabIndex        =   10
      Top             =   180
      Width           =   690
   End
End
Attribute VB_Name = "frmReturnStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo search_

If MsgBox("Want to Update ?", vbQuestion + vbYesNo) = vbYes Then
Screen.MousePointer = vbHourglass
   For I = 1 To vs.Rows - 1
    If (vs.TextMatrix(I, 0) <> "" And vs.TextMatrix(I, 6) <> "") Then
       con.Execute "update return_auth set [Status]='" & vs.TextMatrix(I, 6) & "' where (Id=" & vs.TextMatrix(I, 0) & ")"
    End If
   Next
Screen.MousePointer = vbDefault
End If


Exit Sub

search_:

MsgBox "" & err.DESCRIPTION
Screen.MousePointer = vbDefault

End Sub
Private Sub cmdView_Click()

On Error GoTo search_

Screen.MousePointer = vbHourglass

Set RS = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Dim k1 As Integer
Dim rss_ As New ADODB.Recordset

k1 = 0
vs.Clear
vs.Rows = 2
vs.Cols = 8
Dim totFrt As Double
Dim dt_str1 As String
Dim s1 As String

totFrt = 0

dt_str1 = "(Recd_Date >= convert(smalldatetime,'" & txtFrom.value & "',103) and Recd_Date <= convert(smalldatetime,'" & txtTo.value & "',103))"

Dim rs_ As ADODB.Recordset
Set rs_ = New ADODB.Recordset
s1 = 0

Set rss_ = New ADODB.Recordset


If Option_Pending.value = True Then
   rss_.Open "SELECT ID,Recd_Date,Party,Bundles,Mode,Remarks,[Status],Dated FROM return_auth where [Status]='Not Approved' and " & dt_str1 & " order by id", con
ElseIf Option_Clear.value = True Then
   rss_.Open "SELECT ID,Recd_Date,Party,Bundles,Mode,Remarks,[Status],Dated FROM return_auth where [Status]='Approved' and " & dt_str1 & " order by id", con
ElseIf Option_Both.value = True Then
   rss_.Open "SELECT ID,Recd_Date,Party,Bundles,Mode,Remarks,[Status],Dated FROM return_auth where " & dt_str1 & " order by id", con
End If

For I = 1 To rss_.RecordCount

  
If rss_.EOF = False Then
  DoEvents
  DoEvents

  
  vs.TextMatrix(I, 0) = rss_!ID
  vs.TextMatrix(I, 1) = rss_!Recd_Date
  vs.TextMatrix(I, 2) = rss_!party
  vs.TextMatrix(I, 3) = rss_!bundles
  vs.TextMatrix(I, 4) = rss_!mode
  vs.TextMatrix(I, 5) = rss_!Remarks
  vs.TextMatrix(I, 6) = rss_!Status & ""
  vs.TextMatrix(I, 7) = rss_!Dated & ""
  
  If k1 = 0 Then
    vs.Cell(flexcpBackColor, I, 0) = vbWhite
    vs.Cell(flexcpBackColor, I, 1) = vbWhite
    vs.Cell(flexcpBackColor, I, 2) = vbWhite
    vs.Cell(flexcpBackColor, I, 3) = vbWhite
    vs.Cell(flexcpBackColor, I, 4) = vbWhite
    vs.Cell(flexcpBackColor, I, 5) = vbWhite
    vs.Cell(flexcpBackColor, I, 6) = vbWhite
   
   Else
    vs.Cell(flexcpBackColor, I, 0) = &HC0FFFF
    vs.Cell(flexcpBackColor, I, 1) = &HC0FFFF
    vs.Cell(flexcpBackColor, I, 2) = &HC0FFFF
    vs.Cell(flexcpBackColor, I, 3) = &HC0FFFF
    vs.Cell(flexcpBackColor, I, 4) = &HC0FFFF
    vs.Cell(flexcpBackColor, I, 5) = &HC0FFFF
    vs.Cell(flexcpBackColor, I, 6) = &HC0FFFF
    
  End If
  vs.Rows = vs.Rows + 1
  
  
  
  DoEvents
  DoEvents
  rss_.MoveNext
  
  
  
 
  
  
End If
  
Next


vs.FormatString = "Id|Recd.Date|Party|Bundles|Mode|Remarks|Status|Dated"
vs.ColWidth(0) = 800
vs.ColWidth(1) = 1000
vs.ColWidth(2) = 2500
vs.ColWidth(3) = 1000
vs.ColWidth(4) = 2000
vs.ColWidth(5) = 2500
vs.ColWidth(6) = 1800
vs.ColWidth(7) = 1000


vs.MergeCells = flexMergeFixedOnly
vs.MergeCol(0) = True
vs.MergeCol(1) = True
'vs.MergeCol(2) = True
'vs.MergeCol(7) = True

cmdUpdate.Enabled = True
Screen.MousePointer = vbDefault

Exit Sub

search_:
Screen.MousePointer = vbDefault
message ("" & err.DESCRIPTION)


End Sub


Private Sub Form_Load()
Me.Left = 50
Me.Top = 50

Me.Height = 9180
Me.Width = 13770

txtFrom.value = Format(Date, "dd/MM/yyyy")
txtTo.value = Format(Date, "dd/MM/yyyy")

End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'    If vs.Col = 8 Then
'       For h1 = vs.Row To vs.Rows - 1
'          If vs.TextMatrix(vs.RowSel, 0) = vs.TextMatrix(h1, 0) Then
'             vs.TextMatrix(h1, 8) = vs.TextMatrix(vs.RowSel, 8)
'             SendKeys "{down}"
'          End If
'       Next
'    End If
'End If
End Sub

Private Sub vs_SelChange()
   If vs.Col = 6 Then
      vs.Editable = flexEDKbdMouse
   Else
      vs.Editable = flexEDNone
   End If
End Sub

Private Sub VSFlexGrid1_Click()

End Sub

