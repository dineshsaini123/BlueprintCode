VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTCS 
   Caption         =   "TCS Report"
   ClientHeight    =   7656
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9468
   Icon            =   "frmTCS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7656
   ScaleWidth      =   9468
   Begin VB.TextBox txtAmt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2160
      MaxLength       =   39
      TabIndex        =   2
      Text            =   "5000000"
      Top             =   612
      Width           =   1392
   End
   Begin VB.CommandButton cmdExit_12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   615
      Left            =   7704
      Picture         =   "frmTCS.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   180
      Width           =   1428
   End
   Begin VB.CommandButton cmdRepQty 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export to Excel"
      Height          =   615
      Left            =   6228
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   180
      Width           =   1428
   End
   Begin VB.CommandButton cmdAdd_1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
      Height          =   615
      Left            =   4752
      Picture         =   "frmTCS.frx":0BF0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   180
      Width           =   1428
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   6228
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   9180
      _cx             =   16192
      _cy             =   10985
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
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
      BackColorBkg    =   16777215
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
      Rows            =   1000
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTCS.frx":17D4
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
      Editable        =   1
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
         Left            =   14580
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   4020
         Width           =   195
      End
   End
   Begin MSComCtl2.DTPicker txtFrom 
      Height          =   312
      Left            =   252
      TabIndex        =   0
      Top             =   216
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   174915585
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txtToDate 
      Height          =   312
      Left            =   2160
      TabIndex        =   1
      Top             =   216
      Width           =   1416
      _ExtentX        =   2498
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   174915585
      CurrentDate     =   42409
   End
   Begin VB.Label lblAson 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   1404
      TabIndex        =   9
      Top             =   648
      Width           =   780
   End
   Begin VB.Label lblAson 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   1680
      TabIndex        =   8
      Top             =   252
      Width           =   276
   End
End
Attribute VB_Name = "frmTCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_1_Click()

Screen.MousePointer = vbHourglass

vs.Cols = 2
vs.Rows = 1

Dim k1 As Integer
k1 = 1

Dim str_date As String
str_date = "(VoucherDate>=convert(smalldatetime,'" & txtFrom.value & "',103) and VoucherDate<=convert(smalldatetime,'" & txtToDate.value & "',103))"


If RS.State = 1 Then RS.close
RS.Open "SELECT SubLedger,sum(amount) as amount FROM VOUCHERS where GenLedger='SUNDRY DEBTORS'" & _
" and DebitorCredit='C' and " & str_date & " group by SubLedger", con
While RS.EOF = False

If RS(1) >= Val(txtAmt.Text) Then
    vs.Rows = vs.Rows + 1
    vs.TextMatrix(k1, 0) = RS(0)
    vs.TextMatrix(k1, 1) = RS(1)
    
    k1 = k1 + 1
    
End If


RS.MoveNext
Wend


vs.FormatString = "Party Name|>Total Receipt Amt"
vs.ColWidth(0) = 6400
vs.ColWidth(1) = 2000

Screen.MousePointer = vbDefault

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdRepQty_Click()
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim str_ As String




If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r As Long
Dim Q1, q2, J As Double



row_ = 1
col_ = 1

xl.Columns("A:H").ColumnWidth = 12
J = 2


For I = 0 To vs.Rows - 1
    For J = 0 To vs.Cols - 1
      
        xlSheet.Cells(row_, col_).value = vs.TextMatrix(I, J)
       
        col_ = col_ + 1
    Next
    row_ = row_ + 1
    col_ = 1
Next

MsgBox "Task Completed....", vbInformation
End Sub

Private Sub Form_Load()

Me.Width = 9500
Me.Height = 8000

txtFrom.value = Format(fromDate_setup, "dd/MM/yyyy")
txtToDate.value = Format(toDate_setup, "dd/MM/yyyy")

End Sub
