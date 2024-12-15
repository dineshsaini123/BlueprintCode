VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmYearReport 
   Caption         =   "Report"
   ClientHeight    =   9732
   ClientLeft      =   60
   ClientTop       =   408
   ClientWidth     =   17988
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmYearReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9732
   ScaleWidth      =   17988
   Begin VB.CommandButton Command1_excel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export To Excel"
      Height          =   645
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   300
      Width           =   1545
   End
   Begin VB.CommandButton cmdview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
      Height          =   645
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   300
      Width           =   1365
   End
   Begin VB.TextBox txtgp 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1425
      TabIndex        =   4
      Top             =   1350
      Width           =   1395
   End
   Begin VB.CommandButton cmdFilterData 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Filter Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1050
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton CommandReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      Height          =   645
      Left            =   9450
      Picture         =   "frmYearReport.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   300
      Width           =   1275
   End
   Begin MSComCtl2.DTPicker txtFromSale 
      Height          =   390
      Left            =   1440
      TabIndex        =   0
      Top             =   300
      Width           =   1530
      _ExtentX        =   2709
      _ExtentY        =   699
      _Version        =   393216
      Format          =   162922497
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txttoSale 
      Height          =   390
      Left            =   3450
      TabIndex        =   1
      Top             =   300
      Width           =   1635
      _ExtentX        =   2879
      _ExtentY        =   699
      _Version        =   393216
      Format          =   162922497
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txtFromSaleRet 
      Height          =   390
      Left            =   1440
      TabIndex        =   2
      Top             =   825
      Width           =   1530
      _ExtentX        =   2709
      _ExtentY        =   677
      _Version        =   393216
      Format          =   162922497
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txttoSaleRet 
      Height          =   390
      Left            =   3450
      TabIndex        =   3
      Top             =   825
      Width           =   1635
      _ExtentX        =   2879
      _ExtentY        =   677
      _Version        =   393216
      Format          =   130023425
      CurrentDate     =   42409
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7590
      Left            =   0
      TabIndex        =   11
      Top             =   2025
      Width           =   17925
      _cx             =   31618
      _cy             =   13388
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
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
      ForeColorSel    =   12582912
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmYearReport.frx":0BF0
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Group Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   12
      Top             =   1425
      Width           =   1350
   End
   Begin VB.Label lblAson 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3045
      TabIndex        =   10
      Top             =   450
      Width           =   315
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Range :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   9
      Top             =   375
      Width           =   1350
   End
   Begin VB.Label Label2 
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3045
      TabIndex        =   8
      Top             =   900
      Width           =   315
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale Ret.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   300
      TabIndex        =   7
      Top             =   825
      Width           =   1350
   End
End
Attribute VB_Name = "frmYearReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFilterData_Click()

Dim s_
Dim a_strResult
Dim k1 As Integer

Dim str_date As String

str_date = "(fdate>=convert(smalldatetime,'" & txtFromSale.value & "',103) " & _
" and tdate<=convert(smalldatetime,'" & txttoSale.value & "',103))"

str_date = str_date & " or (fdate>=convert(smalldatetime,'" & txtFromSaleRet.value & "',103) " & _
" and tdate<=convert(smalldatetime,'" & txttoSaleRet.value & "',103))"



con.Execute "delete from tmp_yearlyTBL"

k1 = 0
s_ = ""

If RS.State = 1 Then RS.close
RS.Open "select dbname from Year_DBName_And_DateRangeTbl " & _
" where (fdate >= convert(smalldatetime,'" & txtFromSale.value & "',103) " & _
" and fdate <= convert(smalldatetime,'" & txttoSale.value & "',103)) or (tdate >= convert(smalldatetime,'" & txtFromSale.value & "',103) and tdate <= convert(smalldatetime,'" & txttoSale.value & "',103))", con




While RS.EOF = False

If s_ = "" Then
  s_ = RS!dbname
Else
  s_ = s_ & "," & RS!dbname
End If


RS.MoveNext
k1 = k1 + 1

Wend

If (k1 > 1) Then
   a_strResult = Split(s_, ",")
   con.Execute "exec Sp_yearlyQry '" & a_strResult(0) & "','" & a_strResult(1) & "','repwise_sale'"
ElseIf k1 = 1 Then
   con.Execute "exec Sp_yearlyQry '" & s_ & "','','repwise_sale'"

End If


''=================================================================
s_ = ""
k1 = 0

If RS.State = 1 Then RS.close
RS.Open "select dbname from Year_DBName_And_DateRangeTbl " & _
" where (fdate >= convert(smalldatetime,'" & txtFromSaleRet.value & "',103) " & _
" and fdate <= convert(smalldatetime,'" & txttoSaleRet.value & "',103)) or (tdate >= convert(smalldatetime,'" & txtFromSaleRet.value & "',103) and tdate <= convert(smalldatetime,'" & txttoSaleRet.value & "',103))"



While RS.EOF = False

If s_ = "" Then
  s_ = RS!dbname
Else
  s_ = s_ & "," & RS!dbname
End If


RS.MoveNext
k1 = k1 + 1

Wend

If (k1 > 1) Then
   a_strResult = Split(s_, ",")
   con.Execute "exec Sp_yearlyQry '" & a_strResult(0) & "','" & a_strResult(1) & "','repwise_saleret'"
ElseIf k1 = 1 Then
   con.Execute "exec Sp_yearlyQry '" & s_ & "','" & "" & "','repwise_saleret'"

End If

''------------------------------------------------------------------
If txtgp.Text <> "" Then
   con.Execute "delete from tmp_yearlyTBL where groupcode<>'" & txtgp.Text & "'"
End If


''MsgBox "Data Filter ......."

End Sub
Private Sub cmdView_Click()

Screen.MousePointer = vbHourglass

vs.Clear

cmdFilterData_Click

Dim str_date As String

str_date = "(invoicedate>=convert(smalldatetime,'" & txtFromSale.value & "',103) " & _
" and invoicedate<=convert(smalldatetime,'" & txttoSale.value & "',103))"

str_date = str_date & " or (invoicedate>=convert(smalldatetime,'" & txtFromSaleRet.value & "',103) " & _
" and invoicedate<=convert(smalldatetime,'" & txttoSaleRet.value & "',103))"



Dim ff As ADODB.Recordset

Set ff = New ADODB.Recordset

If ff.State = 1 Then ff.close

ff.Open "select agentName,sum(QtySale) as QtySale,sum(SaleAmt) as SaleAmt,sum(QtySaleRet) as QtySaleRet " & _
",sum(SaleRetAmt) as SaleRetAmt,sum(QtySale)-sum(QtySaleRet)as NetQty, ROUND(sum(SaleAmt)-sum(SaleRetAmt),2) as NetSale " & _
" from  Yearly_RepwiseSaleReturnQry where len(agentname)>0 and " & str_date & " group by agentname ", con

Set vs.DataSource = ff

vs.FormatString = "AgentName|>QtySale|>SaleAmt|>QtySaleRet|>SaleRetAmt|>NetQty|>NetSale"

vs.ColWidth(0) = 3500
vs.ColWidth(1) = 2000
vs.ColWidth(2) = 2000
vs.ColWidth(3) = 2000
vs.ColWidth(4) = 2000
vs.ColWidth(5) = 2000
vs.ColWidth(6) = 2000

MsgBox "List Generated ......."

Screen.MousePointer = vbDefault

End Sub

Private Sub Command1_excel_Click()


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

Private Sub CommandReturn_Click()
Unload Me
End Sub

Private Sub Form_Load()

Me.Top = 10
Me.Left = 10

Me.Width = 18500
Me.Height = 10100


txtFromSale.value = fromDate_setup
txttoSale.value = toDate_setup


txtFromSaleRet.value = fromDate_setup
txttoSaleRet.value = toDate_setup


End Sub
