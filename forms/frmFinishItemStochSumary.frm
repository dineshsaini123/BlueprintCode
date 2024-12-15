VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmFinishItemStochSumary 
   Caption         =   "Stock Summary ..."
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11550
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar pb 
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   8580
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ComboBox CboPName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4695
      TabIndex        =   8
      Top             =   435
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   10845
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   195
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   105
      TabIndex        =   2
      Top             =   75
      Width           =   4455
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   300
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   19660803
         UpDown          =   -1  'True
         CurrentDate     =   37701
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   300
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   19660803
         UpDown          =   -1  'True
         CurrentDate     =   37701
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   285
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   9885
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   195
      Width           =   945
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8835
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   195
      Width           =   1020
   End
   Begin Crystal.CrystalReport CR 
      Left            =   13380
      Top             =   180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7695
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   14835
      _cx             =   26167
      _cy             =   13573
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      BackColorSel    =   13888387
      ForeColorSel    =   16711680
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
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
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
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
      ExplorerBar     =   7
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
   Begin VB.Label unit 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   9225
      TabIndex        =   11
      Top             =   465
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Left            =   4695
      TabIndex        =   10
      Top             =   165
      Visible         =   0   'False
      Width           =   870
   End
End
Attribute VB_Name = "frmFinishItemStochSumary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboItemName_Click()
  On Error Resume Next
  fillGrid
End Sub
Private Function LoadName()
CboPName.Clear
Set rs = New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "Select distinct(BOOK) from copymaster", CON, adOpenKeyset, adLockReadOnly
If rs.EOF = False Then
    Do While Not rs.EOF
        CboPName.AddItem rs.Fields(0)
        rs.MoveNext
    Loop
End If
End Function
Private Sub CboPName_Click()
    If rs.State = 1 Then rs.Close
    rs.Open "select unit from ItemMaster where ItemName='" & CboPName.Text & "'", CON
    If rs.EOF = False Then
       unit.Caption = rs(0)
    End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
   
   'Call cmdSearch_Click
   
   
   CR.Reset
   CR.Connect = constr
   CR.ReportFileName = App.Path & "\Reports\stocksummary.rpt"
   CR.Formulas(0) = "Fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "Todate='" & ToDate.Value & "'"
   ss5 = "Chitra Exports   75, Mohkampur, Ind. Area, Phase-II"
   CR.Formulas(2) = "rptheader='" & ss5 & "'"
   
   CR.WindowShowCloseBtn = True
   CR.WindowShowPrintBtn = True
   CR.WindowControlBox = True
   CR.WindowShowPrintSetupBtn = True
   CR.WindowShowProgressCtls = True
   CR.WindowState = crptMaximized
   CR.Action = 0
   
 

End Sub
Sub UpdateStock()
     
Dim save As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim search As New ADODB.Recordset
Dim rs_sale As New ADODB.Recordset
Dim rs_rec As New ADODB.Recordset
Dim rs_ret As New ADODB.Recordset
Dim rs_op As New ADODB.Recordset
Dim rs_sep As New ADODB.Recordset
Dim rs_other As New ADODB.Recordset
Dim op, op1, sep, ret, Sale, rec, OTHER, total1, total2 As Long
Dim bb1 As Boolean

bb1 = False

op = 0
op1 = 0
sep = 0
ret = 0
Sale = 0
rec = 0
OTHER = 0


Screen.MousePointer = vbHourglass
Set save = New ADODB.Recordset
If save.State = 1 Then save.Close
save.Open "select * from BookStockSummary", CON, adOpenDynamic, adLockOptimistic

If rs.State = 1 Then rs.Close
rs.Open "select BOOKNO,ProductQuality,TypeofProduct,rulling,rate,NoofPages from Copymaster where " & stringyear & " order by BOOKNO", CON, adOpenKeyset, adLockReadOnly
If rs.EOF = False Then pb.max = rs.RecordCount

While rs.EOF = False
  
  
  
  '- Code For Opening
  '--- code rec
  If rs_rec.State = 1 Then rs_rec.Close
  rs_rec.Open "SELECT SUM(QUANTITY) FROM ProductReceipt where (PCODE='" & rs.Fields("BOOKNO").Value & "'" & _
  " AND CONVERT(smalldatetime,RECDATE,103)<convert(smalldatetime,'" & FromDate.Value & "',103)) and " & stringyear, CON
  If Not IsNull(rs_rec(0)) Then
  rec = rs_rec(0)
  Else
  rec = 0
  End If
  
  
  
 
  
  '-------------------------------------------------------------------------
  '--- code ret
  If rs_ret.State = 1 Then rs_ret.Close
  rs_ret.Open "SELECT SUM(QUANTITY) FROM CREDITB where (bookcode='" & rs.Fields("BOOKNO").Value & "'" & _
  " AND convert(smalldatetime,INVOICEDATE,103)<convert(smalldatetime,'" & FromDate.Value & "',103)) and " & stringyear, CON, adOpenKeyset, adLockReadOnly
  If Not IsNull(rs_ret(0)) Then
  ret = rs_ret(0)
  Else
  ret = 0
  End If
  
  
  
  
  
  total1 = (rec + ret)
   
  '--- code Sale
  If rs_sale.State = 1 Then rs_sale.Close
  rs_sale.Open "SELECT SUM(QUANTITY) FROM INVOICEB where (bookcode='" & rs.Fields("BOOKNO").Value & "'" & _
  " AND convert(smalldatetime,INVOICEDATE,103)<convert(smalldatetime,'" & FromDate.Value & "',103)) and " & stringyear, CON, adOpenKeyset, adLockReadOnly
  If Not IsNull(rs_sale(0)) Then
     Sale = rs_sale(0)
  Else
     Sale = 0
  End If
  

  total2 = Sale
  op = (total1 - total2)
  
  
  
  '- Code For Transaction
  '--- code rec
  If rs_rec.State = 1 Then rs_rec.Close
  rs_rec.Open "SELECT SUM(QUANTITY) FROM ProductReceipt where (PCode='" & rs.Fields("BOOKNO").Value & "' AND " & _
  "convert(smalldatetime,RECDATE,103)>=convert(smalldatetime,'" & FromDate.Value & "',103) " & _
  "and convert(smalldatetime,RECDATE,103)<=convert(smalldatetime,'" & ToDate.Value & "',103)) and " & stringyear, CON
  If Not IsNull(rs_rec(0)) Then
  rec = rs_rec(0)
  Else
  rec = 0
  End If
  
  
  
  
  '--- code ret
  If rs_ret.State = 1 Then rs_ret.Close
  rs_ret.Open "SELECT SUM(QUANTITY) FROM CREDITB where BOOKCode= '" & rs.Fields("BOOKNO").Value & "'" & _
  " AND convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & FromDate.Value & "',103) " & _
  "and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & ToDate.Value & "',103) and " & stringyear, CON, adOpenKeyset, adLockReadOnly
  If Not IsNull(rs_ret(0)) Then
  ret = rs_ret(0)
  Else
  ret = 0
  End If
  
  total1 = (rec + ret)
   
   
  '--- code Sale
  If rs_sale.State = 1 Then rs_sale.Close
  rs_sale.Open "SELECT SUM(QUANTITY) FROM INVOICEB where BOOKCode= '" & rs.Fields("BOOKNO").Value & "'" & _
  " AND convert(smalldatetime,INVOICEDATE,103)>=convert(smalldatetime,'" & FromDate.Value & "',103) " & _
  "and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & ToDate.Value & "',103) and " & stringyear, CON, adOpenKeyset, adLockReadOnly
  If Not IsNull(rs_sale(0)) Then
  Sale = rs_sale(0)
  Else
  Sale = 0
  End If
  
  total2 = (Sale)
  
  '--------------------------------------------
  If (rec + ret + op1 + op) > 0 Then
     bb1 = True
  End If
  If (Sale + sep + OTHER) > 0 Then
    bb1 = True
  End If
  
  
 If bb1 = True Then
 
    save.addNew
    save!Code = rs.Fields(0).Value
    save!book = rs!TypeofProduct + " (" + rs!rulling + ")" + str(rs!NoOfPages) + " " + rs!ProductQuality
    save!op = op
    save!Receive = rec
    save!Return = ret
    save!OtherRec = op1
    save!Rec_Total = (rec + ret + op1 + op)
    save!sales = Sale
    save!Specimen = sep
    save!OTHER = OTHER
    save!Issue_Total = (Sale + sep + OTHER)
    save.Update
    
    bb1 = False
 End If
  
  rs.MoveNext
   
If pb.Value < pb.max Then
pb.Value = pb.Value + 1
End If

Wend


Screen.MousePointer = vbDefault

End Sub
Private Sub cmdSearch_Click()
    
    Screen.MousePointer = vbHourglass
    
    
    Dim opening As Long
    Dim search As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = CON
    cmd.CommandText = "delete from BookStockSummary where len(code)>0"
    cmd.Execute
    pb.Visible = True
    UpdateStock
    
    fillGrid
    
    pb.Visible = False
    
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
     Unload Me
  End If
End Sub
Sub VsWidth()
       vs.Cols = 8
       vs.TextMatrix(0, 0) = "Code"
       vs.TextMatrix(0, 1) = "Book Name"
       vs.TextMatrix(0, 2) = "Opening"
       vs.TextMatrix(0, 3) = "Receive"
       vs.TextMatrix(0, 4) = "Return"
       vs.TextMatrix(0, 5) = "Total"
       vs.TextMatrix(0, 6) = "Sale"
       vs.TextMatrix(0, 7) = "Closing Balance"
       
       vs.ColWidth(0) = 1000
       vs.ColWidth(1) = 5500
       vs.ColWidth(2) = 1200
       vs.ColWidth(3) = 1200
       vs.ColWidth(4) = 1200
       vs.ColWidth(5) = 1200
       vs.ColWidth(6) = 1200
       vs.ColWidth(7) = 1200
 
End Sub
Sub fillGrid()
    
    Dim fill As New ADODB.Recordset
    If fill.State = 1 Then fill.Close
    fill.Open "select * from BookStockSummary order by code", CON
    If fill.EOF = False Then
       vs.Rows = fill.RecordCount + 1
       For i = 1 To fill.RecordCount
         If fill.EOF = False Then
            vs.TextMatrix(i, 0) = fill.Fields(0).Value
            vs.TextMatrix(i, 1) = fill.Fields(1).Value
            vs.TextMatrix(i, 2) = fill.Fields(2).Value
            
            vs.TextMatrix(i, 3) = fill.Fields("Receive").Value
            vs.TextMatrix(i, 4) = fill.Fields("Return").Value
            
            vs.TextMatrix(i, 5) = fill.Fields("Rec_Total").Value
            vs.TextMatrix(i, 6) = fill.Fields("Sales").Value
            
            vs.TextMatrix(i, 7) = (fill.Fields("Rec_Total").Value - fill.Fields("Sales").Value)
            
         End If
         fill.MoveNext
       Next
    End If
    VsWidth

End Sub

Private Sub Form_Load()
ToDate.Value = Date
fillGrid
FromDate.Value = Date
LoadName

If rs.State = 1 Then rs.Close
rs.Open "setup", CON, adOpenStatic, adLockReadOnly, adCmdTable
CNSetup
FromDate.Value = rs!yarfrom
ToDate.Value = rs!yarto


End Sub
Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then ToDate.SetFocus
End Sub
