VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form frmFinishItemStochSumary 
   Caption         =   "Book Stock Summary"
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
      ItemData        =   "frmFinishItemStochSumary.frx":0000
      Left            =   4680
      List            =   "frmFinishItemStochSumary.frx":0002
      TabIndex        =   8
      Top             =   435
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
      Height          =   345
      Left            =   9525
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   375
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
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
         Format          =   20119555
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
         Format          =   20119555
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
      Height          =   345
      Left            =   8565
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   375
      Width           =   945
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Show"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7515
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   375
      Width           =   1020
   End
   Begin Crystal.CrystalReport CR 
      Left            =   11085
      Top             =   -75
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   6165
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   11805
      _cx             =   20823
      _cy             =   10874
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
      AllowUserResizing=   0
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
      Left            =   7965
      TabIndex        =   11
      Top             =   345
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      Width           =   750
   End
End
Attribute VB_Name = "frmFinishItemStochSumary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub cboItemName_Click()
  On Error Resume Next
  fillgrid
End Sub
Private Function LoadName()
'CboPName.Clear
'Set rs = New ADODB.Recordset
'If rs.State = 1 Then rs.Close
'rs.Open "Select distinct(BOOKNAME) from books ", CON, adOpenDynamic, adLockOptimistic
'If rs.EOF = False Then
'    Do While Not rs.EOF
'        CboPName.AddItem rs.Fields(0)
'        rs.MoveNext
'    Loop
'End If

CboPName.AddItem "Raw Item Stock"
CboPName.AddItem "Finish Item Stock"

End Function
Private Sub CboPName_Click()
    'If rs.State = 1 Then rs.Close
    'rs.Open "select unit from ItemMaster where ItemName='" & CboPName.Text & "'", CON
    'If rs.EOF = False Then
    '   unit.Caption = rs(0)
    'End If
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
   
   'Call cmdSearch_Click
   
   
   cr.Reset
   cr.Connect = "filedsn=ims;uid=sa;pwd=sidcdbserver"
   cr.ReportFileName = strrptpath & "\reports\stocksummary.rpt"
   cr.ReplaceSelectionFormula "{BookStockSummary.Issue_Total}<>0 and {BookStockSummary.fyear}='" & main.session & "' and {BookStockSummary.setupid}=" & main.setupid
   cr.Formulas(0) = "Fromdate='From: " & fromdate.Value & " To: " & todate.Value & "' "
   cr.WindowShowCloseBtn = True
   cr.WindowShowPrintBtn = True
   cr.WindowControlBox = True
   cr.WindowShowPrintSetupBtn = True
   cr.WindowShowProgressCtls = True
   cr.WindowState = crptMaximized
   cr.Action = 0
   
 


End Sub
Sub UpdateStock()
     
Dim Save As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim search As New ADODB.Recordset
Dim rs_sale As New ADODB.Recordset
Dim rs_rec As New ADODB.Recordset
Dim rs_ret As New ADODB.Recordset
Dim rs_op As New ADODB.Recordset
Dim rs_sep As New ADODB.Recordset
Dim rs_other As New ADODB.Recordset
Dim op, issue, sep, ret, Sale, rec, OTHER, total1, total2 As Long
op = 0
issue = 0
sep = 0
ret = 0
Sale = 0
rec = 0
OTHER = 0


Screen.MousePointer = vbHourglass
Set Save = New ADODB.Recordset
If Save.State = 1 Then Save.Close
Save.Open "select * from BookStockSummary where " & stridnyear & "", CON, adOpenDynamic, adLockOptimistic

If rs.State = 1 Then rs.Close
If CboPName.Text = "Raw Item Stock" Then
rs.Open "select * from  BOOKS where " & stridnyear & " and GROUPCODE='No' order by BOOKCODE", CON
Else
rs.Open "select * from  BOOKS where " & stridnyear & " and  GROUPCODE='Yes' order by BOOKCODE", CON
End If

While rs.EOF = False
  
  
  '- Code For Opening
  '--- code rec
  If rs_rec.State = 1 Then rs_rec.Close
  If CboPName.Text = "Raw Item Stock" Then
  rs_rec.Open "SELECT SUM(QUANTITY) FROM ReceiveRegister where ( " & stridnyear & " and groupcode='No' and BOOKCODE='" & rs.Fields("BOOKCODE").Value & "' AND convert(char(10),INVOICEDATE,103)<convert(char(10),'" & fromdate.Value & "',103))", CON
  Else
  rs_rec.Open "SELECT SUM(QUANTITY) FROM ReceiveRegister where ( " & stridnyear & " and groupcode='Yes' and BOOKCODE='" & rs.Fields("BOOKCODE").Value & "' AND convert(char(10),INVOICEDATE,103)<convert(char(10),'" & fromdate.Value & "',103))", CON
  End If
  
  If Not IsNull(rs_rec(0)) Then
  rec = rs_rec(0)
  Else
  rec = 0
  End If
  
  '-----Issue code
  
  If rs_rec.State = 1 Then rs_rec.Close
  If CboPName.Text = "Raw Item Stock" Then
  'rs_rec.Open "SELECT SUM(QUANTITY) FROM IssueRegister where (groupcode='No' and BOOKCODE='" & rs.Fields("BOOKCODE").Value & "' AND convert(char(10),INVOICEDATE,103)<convert(char(10),'" & FromDate.Value & "',103))", CON
  rs_rec.Open "SELECT SUM(QTY) FROM MfgTable where ( " & stridnyear & " and RCode='" & rs.Fields("BOOKCODE").Value & "' AND convert(char(10),Dates,103)<convert(char(10),'" & fromdate.Value & "',103))", CON
  Else
  rs_rec.Open "SELECT SUM(QUANTITY) FROM SaleRegister1 where ( " & stridnyear & " and groupcode='Yes' and BOOKCODE='" & rs.Fields("BOOKCODE").Value & "' AND convert(char(10),INVOICEDATE,103)<convert(char(10),'" & fromdate.Value & "',103))", CON
  End If
  
  If Not IsNull(rs_rec(0)) Then
  issue = rs_rec(0)
  Else
  issue = 0
  End If
  
  
   
  op = (rec - issue)
  
  
  '- Code For Transaction
  '--- code rec
  If rs_rec.State = 1 Then rs_rec.Close
  If CboPName.Text = "Raw Item Stock" Then
  rs_rec.Open "SELECT SUM(QUANTITY) FROM ReceiveRegister where ( " & stridnyear & " and groupcode='No' and BOOKCODE='" & rs.Fields("BOOKCODE").Value & "' AND (convert(char(10),INVOICEDATE,103)>=convert(char(10),'" & fromdate.Value & "',103) and convert(char(10),INVOICEDATE,103)<=convert(char(10),'" & todate.Value & "',103)))", CON
  Else
  rs_rec.Open "SELECT SUM(QUANTITY) FROM ReceiveRegister where ( " & stridnyear & " and groupcode='Yes' and BOOKCODE='" & rs.Fields("BOOKCODE").Value & "' AND (convert(char(10),INVOICEDATE,103)>=convert(char(10),'" & fromdate.Value & "',103) and convert(char(10),INVOICEDATE,103)<=convert(char(10),'" & todate.Value & "',103)))", CON
  End If
  
  If Not IsNull(rs_rec(0)) Then
  rec = rs_rec(0)
  Else
  rec = 0
  End If
  
  '-----Issue code
  
  If rs_rec.State = 1 Then rs_rec.Close
  If CboPName.Text = "Raw Item Stock" Then
  'rs_rec.Open "SELECT SUM(QUANTITY) FROM IssueRegister where (groupcode='No' and BOOKCODE='" & rs.Fields("BOOKCODE").Value & "' AND (convert(char(10),INVOICEDATE,103)>=convert(char(10),'" & FromDate.Value & "',103) and convert(char(10),INVOICEDATE,103)<=convert(char(10),'" & ToDate.Value & "',103)))", CON
  rs_rec.Open "SELECT SUM(QTY) FROM MfgTable where ( " & stridnyear & " and RCode='" & rs.Fields("BOOKCODE").Value & "' AND (convert(char(10),DATEs,103)>=convert(char(10),'" & fromdate.Value & "',103) and convert(char(10),DATEs,103)<=convert(char(10),'" & todate.Value & "',103)))", CON
  Else
  rs_rec.Open "SELECT SUM(QUANTITY) FROM SaleRegister1 where ( " & stridnyear & " and groupcode='Yes' and BOOKCODE='" & rs.Fields("BOOKCODE").Value & "' AND (convert(char(10),INVOICEDATE,103)>=convert(char(10),'" & fromdate.Value & "',103) and convert(char(10),INVOICEDATE,103)<=convert(char(10),'" & todate.Value & "',103)))", CON
  End If
  
  If Not IsNull(rs_rec(0)) Then
  issue = rs_rec(0)
  Else
  issue = 0
  End If
  
  
   
  
  '--------------------------------------------
  Save.AddNew
  Save!Code = rs.Fields(0).Value
  Save!book = rs(1) & ": " & rs!size1 & " " & rs!unit1 & " " & rs!size2 & " " & rs!unit2 & ": " & rs!quality
  Save!op = op
  Save!Receive = rec
  Save!sales = issue
  Save!Issue_Total = (op + (rec - issue))
  'save!quality = rs.Fields("QUALITY").Value
  Save!createdby = main.username
  Save!createdon = Now
  Save!fyear = main.session
  Save!setupid = main.setupid
  Save.Update
  
  
  rs.MoveNext
   
Wend


Screen.MousePointer = vbDefault

End Sub
Private Sub cmdSearch_Click()
    
    Screen.MousePointer = vbHourglass
    
    
    Dim Opening As Long
    Dim search As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = CON
    cmd.CommandText = "delete from BookStockSummary where  " & stridnyear & ""
    cmd.Execute
    
    UpdateStock
    
    fillgrid
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
     Unload Me
  End If
End Sub
Sub VsWidth()
       'vs.Cols = 12
       'vs.TextMatrix(0, 0) = "Code"
       'vs.TextMatrix(0, 1) = "Book Name"
       'vs.TextMatrix(0, 2) = "Opening"
       'vs.TextMatrix(0, 3) = "Receive/Purchase"
       'vs.TextMatrix(0, 4) = "Sale/Issue"
       'vs.TextMatrix(0, 5) = "Closing"
       
       'vs.TextMatrix(0, 4) = "Return"
       'vs.TextMatrix(0, 5) = "Other"
       'vs.TextMatrix(0, 6) = "Total"
       'vs.TextMatrix(0, 8) = "Specimen"
       'vs.TextMatrix(0, 9) = "Other"
       'vs.TextMatrix(0, 10) = "Total"
       
       
       vs.ColWidth(0) = 700
       vs.ColWidth(1) = 5500
       vs.ColWidth(2) = 1300
       vs.ColWidth(3) = 1300
       vs.ColWidth(4) = 1300
       vs.ColWidth(5) = 1300
       
       'vs.ColWidth(6) = 850
       'vs.ColWidth(7) = 850
       'vs.ColWidth(8) = 850
       'vs.ColWidth(9) = 850
       'vs.ColWidth(10) = 850
       'vs.ColWidth(11) = 850
 
End Sub
Sub fillgrid()
    vs.Clear
    vs.Cols = 6
     Dim fill As New ADODB.Recordset
    If fill.State = 1 Then fill.Close
    fill.Open "select * from BookStockSummary where  " & stridnyear & " and Issue_Total<>0 order by code", CON
    If fill.EOF = False Then
       If CboPName.Text = "Raw Item Stock" Then
       vs.FormatString = "Code|Book|Op|Purchase/Receive|Issue For Mfg|Balance"
       Else
       vs.FormatString = "Code|Book|Op|Purchase/Receive|Sales|Balance"
       End If
       vs.Rows = fill.RecordCount + 1
       For I = 1 To fill.RecordCount
         If fill.EOF = False Then
            vs.TextMatrix(I, 0) = fill.Fields("Code").Value
            vs.TextMatrix(I, 1) = fill.Fields("Book").Value & " " & fill.Fields("QUALITY").Value
            vs.TextMatrix(I, 2) = fill.Fields("Op").Value
            vs.TextMatrix(I, 3) = fill.Fields("Receive").Value
            vs.TextMatrix(I, 4) = fill.Fields("Sales").Value
            vs.TextMatrix(I, 5) = fill.Fields("Issue_Total").Value
         End If
         fill.MoveNext
       Next
    End If
    VsWidth

End Sub

Private Sub Form_Load()
Dim rs As New ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "Select * from setup where " & stridnyear & "", CON, adOpenStatic, adLockReadOnly, adCmdText
CNSetup
fromdate.Value = rs!yarfrom
todate.Value = rs!yarto
rs.Close
fillgrid
LoadName
CboPName.ListIndex = 0
End Sub
Private Sub fromdate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then todate.SetFocus
End Sub
