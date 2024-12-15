VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmConsumeItemSummary 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Consumable Item Stock Summary"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConsumeItemSummary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbogp 
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
      TabIndex        =   12
      Top             =   405
      Width           =   3255
   End
   Begin Crystal.CrystalReport CR 
      Left            =   11280
      Top             =   2295
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   8295
      TabIndex        =   2
      Top             =   405
      Width           =   3330
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
      Height          =   435
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1725
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
      Begin MSComCtl2.DTPicker FromDate 
         Height          =   300
         Left            =   840
         TabIndex        =   0
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
         Format          =   63832067
         UpDown          =   -1  'True
         CurrentDate     =   37701
      End
      Begin MSComCtl2.DTPicker ToDate 
         Height          =   300
         Left            =   2880
         TabIndex        =   1
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
         Format          =   63832067
         UpDown          =   -1  'True
         CurrentDate     =   37701
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         Height          =   375
         Left            =   120
         TabIndex        =   6
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
      Height          =   375
      Left            =   10185
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1440
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   915
      Width           =   1440
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7020
      Left            =   120
      TabIndex        =   10
      Top             =   900
      Width           =   10050
      _cx             =   17727
      _cy             =   12382
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmConsumeItemSummary.frx":0442
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name"
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
      Left            =   4725
      TabIndex        =   13
      Top             =   135
      Width           =   1020
   End
   Begin VB.Label unit 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7965
      TabIndex        =   11
      Top             =   450
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
      Left            =   8340
      TabIndex        =   9
      Top             =   165
      Width           =   1185
   End
End
Attribute VB_Name = "frmConsumeItemSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bb5 As Boolean
Dim boo As Boolean
Dim ss10 As String
Dim rs As New ADODB.Recordset
Dim Y As String
Private Sub cboItem_Click()
  On Error Resume Next
  fillgrid
End Sub
Private Function LoadName()
CboPName.Clear
If cbogp.ListIndex = 0 Then
  Y = "Yes"
  Else
  Y = "No"
End If

Set rs = New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "Select distinct(BOOKNAME) from Books where GROUPCODE='" & Y & "'", CON, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
    Do While Not rs.EOF
        CboPName.AddItem rs.Fields(0)
        rs.MoveNext
    Loop
End If
End Function
Sub addgp()
cbogp.AddItem "Ready Made"
cbogp.AddItem "Raw Item"
End Sub
Private Sub cboGp_Click()
 LoadName
End Sub

Private Sub CboPName_Click()
'    If rs.State = 1 Then rs.Close
'    rs.Open "select unit from Books where ItemName='" & CboPName.Text & "'", CON
'    If rs.EOF = False Then
'       unit.Caption = rs(0)
'    End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdprint_Click()

   'Call cmdSearch_Click
  
   CR.Reset
   CR.ReportFileName = App.Path & "\" & "reports\consumvablestocksummary.rpt"
   'CR.ReplaceSelectionFormula "{ConsumableitemstockSummaryRegister.dates}>=datevalue('" & Format(FromDate.Value, "MM/dd/yyyy") & "') and {ConsumableitemstockSummaryRegister.dates}<=datevalue('" & Format(ToDate.Value, "MM/dd/yyyy") & "')"
   CR.Formulas(0) = "Fromdate='" & FromDate.Value & "'"
   CR.Formulas(1) = "Todate='" & ToDate.Value & "'"
   CR.WindowShowCloseBtn = True
   CR.WindowShowPrintBtn = True
   CR.WindowControlBox = True
   CR.WindowShowPrintSetupBtn = True
   CR.WindowShowProgressCtls = True
   CR.WindowState = crptMaximized
   CR.Action = 1


End Sub
Sub UpdateStock()
Dim Save As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Search As New ADODB.Recordset
Dim search1 As New ADODB.Recordset
Dim rsS As New ADODB.Recordset
Dim oprs As New ADODB.Recordset
Dim Receive
Dim Issue
Dim purchaseret
Dim opening
Dim opDate As Date
Dim pr_rec, pr_issue
Dim opValue, v_Purchase, v_Issue
opValue = 0
v_Purchase = 0
v_Issue = 0
Dim quality As String
Screen.MousePointer = vbHourglass
Dim rs_S As New ADODB.Recordset
ss10 = ""
'===================
If Search.State = 1 Then Search.Close
Search.Open "select distinct(Dates) from [dates] where Dates>=datevalue('" & FromDate.Value & "') and Dates<=datevalue('" & ToDate.Value & "')  order by Dates", CON
If Search.EOF = False Then
While Search.EOF = False
    If rs.State = 1 Then rs.Close
    If CboPName.Text <> "" And cbogp.Text <> "" Then
      rs.Open "select ItemName,GROUPCODE from BooksTmp where GROUPCODE='" & Y & "' and ItemName='" & CboPName.Text & "' order by ItemName", CON
    ElseIf cbogp.Text <> "" Then
      rs.Open "select BOOKNAME,GROUPCODE,BOOKCODE from BooksTmp where GROUPCODE='" & Y & "' order by BOOKNAME", CON
    Else
    '  rs.Open "select BOOKNAME,GROUPCODE,BOOKCODE from BooksTmp order by ItemName", CON
    Exit Sub
    End If
   If rs.EOF = False Then
        quality = rs.Fields("GROUPCODE").Value
        While rs.EOF = False
              '------------Calculate Opening--------
              quality = rs.Fields("GROUPCODE").Value
              op = 0
              If oprs.State = 1 Then oprs.Close
              oprs.Open "select OpeningStock from BooksTmp where GROUPCODE='" & quality & "' and BOOKNAME='" & rs.Fields(0).Value & "'", CON
              If oprs.EOF = False Then
                    If rsS.State = 1 Then rsS.Close
                    rsS.Open "select * from ConsumeItemStockSummary where Name='" & rs.Fields(0).Value & "'", CON
                    If rsS.EOF = True Then
                       op = oprs.Fields(0).Value
                    End If
                Else
                    op = 0
              End If
             If search1.State = 1 Then search1.Close
             search1.Open "select sum(QUANTITY) from Purchaseb where INVOICEDATE<datevalue('" & FromDate.Value & "') and BOOKCODE='" & rs.Fields(2).Value & "'", CON
             If Not IsNull(search1.Fields(0)) Then
                pr_rec = search1.Fields(0).Value
             Else
                pr_rec = 0
             End If
             
             If search1.State = 1 Then search1.Close
             search1.Open "select sum(QUANTITY) from Issueb where INVOICEDATE<datevalue('" & FromDate.Value & "') and BOOKCODE='" & rs.Fields(0).Value & "'", CON
             If Not IsNull(search1.Fields(0)) Then
                pr_issue = search1.Fields(0).Value
                Else
                pr_issue = 0
             End If
''                     '--------------------------
''                     If search1.State = 1 Then search1.Close
''                     search1.Open "select sum(amt) from PurchaseValue where RecDate<datevalue('" & FromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "'", CON
''                     If Not IsNull(search1.Fields(0)) Then
''                        v_Purchase = search1.Fields(0).Value
''                        Else
''                        v_Purchase = 0
''                     End If
''                     If search1.State = 1 Then search1.Close
''                     search1.Open "select sum(amt) from IssueValue where IssueDate<datevalue('" & FromDate.Value & "') and Item='" & rs.Fields(0).Value & "'", CON
''                     If Not IsNull(search1.Fields(0)) Then
''                        v_Issue = search1.Fields(0).Value
''                        Else
''                        v_Issue = 0
''                     End If
''                     '-------------------------
             opValue = 0
             opValue = (v_Purchase - v_Issue)
             op = (op + (pr_rec - pr_issue))
             '-----------------end Code-------------------
             If search1.State = 1 Then search1.Close
             search1.Open "select sum(QUANTITY) from Purchaseb where INVOICEDATE=datevalue('" & Search(0) & "') and BOOKCODE='" & rs.Fields(2).Value & "'", CON
             If Not IsNull(search1.Fields(0)) Then
                Receive = search1.Fields(0).Value
                Else
                Receive = 0
             End If
             If search1.State = 1 Then search1.Close
             search1.Open "select sum(QUANTITY) from Issueb where INVOICEDATE=datevalue('" & Search(0) & "') and BOOKCODE='" & rs.Fields(2).Value & "'", CON
             If Not IsNull(search1.Fields(0)) Then
                Issue = search1.Fields(0).Value
                Else
                Issue = 0
             End If
'----------------------------------------------------------------------------
''                    If search1.State = 1 Then search1.Close
''                    search1.Open "select sum(amt) from PurchaseValue where RecDate=datevalue('" & Search(0) & "') and ItemName='" & rs.Fields(0).Value & "'", CON
''                    If Not IsNull(search1.Fields(0)) Then
''                    v_Purchase = search1.Fields(0).Value
''                    Else
''                    v_Purchase = 0
''                    End If
''                    If search1.State = 1 Then search1.Close
''                    search1.Open "select sum(amt) from IssueValue where IssueDate=datevalue('" & Search(0) & "') and Item='" & rs.Fields(0).Value & "'", CON
''                    If Not IsNull(search1.Fields(0)) Then
''                    v_Issue = search1.Fields(0).Value
''                    Else
''                    pr_issue = 0
''                    End If
'--------------------- Data Fiter and Now  Save Coding========================
opening = 0
If Issue = 0 And Receive = 0 And op = 0 Then
   GoTo aaa:
End If
boo = False
If search1.State = 1 Then search1.Close
search1.Open "select * from ConsumeItemStockSummary where name='" & rs.Fields(0).Value & "'", CON
If search1.EOF = False Then
  bb5 = False
Else
  bb5 = True
End If
If bb5 = True Then
    Set Save = New ADODB.Recordset
    If Save.State = 1 Then Save.Close
    Save.Open "select * from ConsumeItemStockSummary", CON, adOpenDynamic, adLockOptimistic
    Save.AddNew
    Save!Name = rs.Fields(0).Value
    Save!gpname = quality
    ss = rs.Fields(0).Value
    Save!dates = Search.Fields(0).Value
    Save!OpenStock = op
    Save!Issue = Issue
    Save!ReceiveStock = Receive
    Save!ClosingStock = ((op + Receive) - (Issue))
    Save!heatno = heatno
    Save!opValue = opValue
    Save!Purchase = v_Purchase
    Save!Receive = v_Issue
    Save.Update
    boo = True
Else
If Issue > 0 Or Receive > 0 Then
    Set Save = New ADODB.Recordset
    If Save.State = 1 Then Save.Close
    Save.Open "select * from ConsumeItemStockSummary", CON, adOpenDynamic, adLockOptimistic
    Save.AddNew
    Save!Name = rs.Fields(0).Value
    Save!gpname = quality
    ss = rs.Fields(0).Value
    Save!dates = Search.Fields(0).Value
    Save!OpenStock = op
    Save!Issue = Issue
    Save!ReceiveStock = Receive
    Save!ClosingStock = ((op + Receive) - (Issue))
    Save!heatno = heatno
    Save!opValue = opValue
    Save!Purchase = v_Purchase
    Save!Receive = v_Issue
    Save.Update
    boo = True
End If
End If
'==================================
 If rsS.State = 1 Then rsS.Close
  rsS.Open "select ClosingStock,Aouto  from ConsumeItemStockSummary where Name='" & ss & "'  order by Dates,aouto", CON
  If rsS.EOF = False Then
      If rsS.RecordCount = 1 Then
        opening = rsS.Fields(0).Value
      ElseIf rsS.RecordCount = 2 Then
        rsS.MoveLast
        rsS.MovePrevious
        opening = rsS.Fields(0).Value
        rsS.MoveNext
        nn = rsS.Fields(1).Value
      ElseIf rsS.RecordCount > 2 Then
        rsS.MoveLast
        rsS.MovePrevious
        opening = rsS.Fields(0).Value
        rsS.MoveNext
        nn = rsS.Fields(1).Value
      End If
      If nn <> "" Then
      Set Save = New ADODB.Recordset
      If Save.State = 1 Then Save.Close
      Save.Open "select * from ConsumeItemStockSummary where Aouto=" & nn & "", CON, adOpenDynamic, adLockOptimistic
      If Save.EOF = False Then
         Save!OpenStock = opening
         Save!ClosingStock = ((opening + Receive) - (Issue))
         If boo = True Then
           Save.Update
         End If
         opening = 0
         nn = ""
      End If
     End If
     Else
  End If
aaa:
ss10 = quality
rs.MoveNext
Wend
End If
Search.MoveNext
Wend

Else

      '------------------- Opening Show No Transiction
If rs.State = 1 Then rs.Close
rs.Open "select distinct(Name) from ConsumeItemStockSummary where GName='" & quality & "'", CON
If rs.EOF = False Then
While rs.EOF = False
'------------Calculate Opening--------
op = 0
If oprs.State = 1 Then oprs.Close
oprs.Open "select OpeningStock from BooksTmp where GROUPCODE='" & quality & "' and ItemName='" & rs.Fields(0).Value & "'", CON
If oprs.EOF = False Then
If rsS.State = 1 Then rsS.Close
 rsS.Open "select * from ConsumeItemStockSummary where GpName='" & quality & "' and Name='" & rs.Fields(0).Value & "'", CON
 If rsS.EOF = True Then
     op = oprs.Fields(0).Value
 End If
 Else
op = 0
End If
If search1.State = 1 Then search1.Close
search1.Open "select sum(QUANTITY) from Purchaseb where INVOICEDATE<datevalue('" & Search(0) & "') and BOOKCODE='" & rs.Fields(2).Value & "'", CON
If Not IsNull(search1.Fields(0)) Then
Receive = search1.Fields(0).Value
Else
Receive = 0
End If
If search1.State = 1 Then search1.Close
search1.Open "select sum(QUANTITY) from Issueb where INVOICEDATE<datevalue('" & Search(0) & "') and BOOKCODE='" & rs.Fields(2).Value & "'", CON
If Not IsNull(search1.Fields(0)) Then
Issue = search1.Fields(0).Value
Else
Issue = 0
End If
'''----------------------------------------------
''If search1.State = 1 Then search1.Close
''search1.Open "select sum(amt) from PurchaseValue where RecDate<datevalue('" & FromDate.Value & "') and ItemName='" & rs.Fields(0).Value & "'", CON
''If Not IsNull(search1.Fields(0)) Then
''v_Purchase = search1.Fields(0).Value
''Else
''v_Purchase = 0
''End If
''If search1.State = 1 Then search1.Close
''search1.Open "select sum(amt) from IssueValue where IssueDate<datevalue('" & FromDate.Value & "') and Item='" & rs.Fields(0).Value & "'", CON
''If Not IsNull(search1.Fields(0)) Then
''v_Issue = search1.Fields(0).Value
''Else
''pr_issue = 0
''End If
'-------------------------
opValue = (v_Purchase - v_Issue)
op = ((op + pr_rec) - pr_issue)
If op = 0 Then GoTo ab1:
Set Save = New ADODB.Recordset
If Save.State = 1 Then Save.Close
Save.Open "select * from ConsumeItemStockSummary", CON, adOpenDynamic, adLockOptimistic
Save.AddNew
Save!GroupName = quality
Save!gpname = quality
Save!Name = rs.Fields(0).Value
ss = rs.Fields(0).Value
Save!dates = FromDate.Value
Save!OpenStock = op
Save!ReceiveStock = Receive
Save!Issue = Issue
Save!ClosingStock = ((Save!OpenStock + Receive) - (Issue))
Save!opValue = opValue
Save!Purchase = v_Purchase
Save!Receive = v_Issue
Save.Update
ab1:
rs.MoveNext
Wend
End If
End If
Screen.MousePointer = vbDefault
End Sub
Sub AddDate()
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim K As Integer
    Dim D As Date
    Set cmd.ActiveConnection = CON
    cmd.CommandText = "delete from dates"
    cmd.Execute
    K = DateDiff("d", FromDate.Value, ToDate.Value)
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dates", CON, adOpenDynamic, adLockOptimistic
     D = FromDate.Value
     For I = 0 To K
           rs.AddNew
           rs!dates = D
           rs.Update
           D = D + 1
     Next
    Set cmd.ActiveConnection = CON
    cmd.CommandText = "delete from bookstmp"
    cmd.Execute
    CON.Execute "insert into bookstmp select BOOKCODE,BOOKNAME,QUALITY,GROUPCODE,OpeningStock from books where OpeningStock>0"
    Dim r As New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "select * from ItemQry", CON, adOpenDynamic, adLockOptimistic
    While rs.EOF = False
        If r.State = 1 Then r.Close
        r.Open "select * from bookstmp where BOOKCODE='" & rs.Fields("BOOKCODE").Value & "'", CON, adOpenDynamic, adLockOptimistic
        If r.EOF = True Then
        r.AddNew
        r.Fields("BOOKCODE").Value = rs.Fields("BOOKCODE").Value
        r.Fields("BOOKNAME").Value = rs.Fields("BOOKNAME").Value
        r.Fields("QUALITY").Value = rs.Fields("QUALITY").Value
        r.Fields("GROUPCODE").Value = rs.Fields("GROUPCODE").Value
        r.Fields("OpeningStock").Value = 0
        r.Update
        End If
        rs.MoveNext
    Wend
End Sub
Private Sub cmdSearch_Click()
    
   ' If CboPName.Text = "" Then
   '    MsgBox "Please Select Item !!", vbCritical
   '    Exit Sub
   ' End If
    Screen.MousePointer = vbHourglass
    AddDate
    Dim opening As Long
    Dim Search As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = CON
    cmd.CommandText = "delete from ConsumeItemStockSummary"
    cmd.Execute
    UpdateStock
    fillgrid
    Screen.MousePointer = vbDefault
End Sub
Sub fillgrid()
     Dim fill As New ADODB.Recordset
    If fill.State = 1 Then fill.Close
    fill.Open "select Name,dates,OpenStock,ReceiveStock,Issue,ClosingStock from ConsumeItemStockSummary where Dates>=datevalue('" & FromDate.Value & "') and Dates<=datevalue('" & ToDate.Value & "') order by Dates,Aouto", CON
    If fill.EOF = False Then
       'Set vs.DataSource = fill
       vs.FormatString = "Name|^date|>OpenStock|>Purchase|>Issue|>ClosingStock"
       
       '-----------------------------------------
       
       vs.Rows = fill.RecordCount + 1
       For I = 1 To fill.RecordCount
         If fill.EOF = False Then
         
            vs.TextMatrix(I, 0) = fill.Fields(0).Value
            vs.TextMatrix(I, 1) = fill.Fields(1).Value
            vs.TextMatrix(I, 2) = fill.Fields(2).Value ' Format(fill.Fields(2).Value, "#,##.00")
            vs.TextMatrix(I, 3) = fill.Fields(3).Value 'Format(fill.Fields(3).Value, "#,##.00")
            vs.TextMatrix(I, 4) = fill.Fields(4).Value 'Format(fill.Fields(4).Value, "#,##.00")
            vs.TextMatrix(I, 5) = fill.Fields(5).Value 'Format(fill.Fields(5).Value, "#,##.00")
            
         
         End If
         fill.MoveNext
       Next
       
       '-----------------------------------------
     
     Else
       Set vs.DataSource = fill
       vs.FormatString = "Name|^date|>OpenStock|>Purchase|>Issue|>ClosingStock"
     End If
     
     VsWidth

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
     Unload Me
  End If
End Sub
Sub VsWidth()
       
       vs.ColWidth(0) = 2500
       vs.ColWidth(1) = 1100
       vs.ColWidth(2) = 1600
       vs.ColWidth(3) = 1600
       vs.ColWidth(4) = 1700
       vs.ColWidth(5) = 1700
       
 
End Sub


Private Sub Form_Load()
ToDate.Value = Date
fillgrid
FromDate.Value = Date

addgp
End Sub
Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then ToDate.SetFocus
End Sub
