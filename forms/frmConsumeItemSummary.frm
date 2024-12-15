VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmConsumeItemSummary 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Item Stock Summary"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15060
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10545
   ScaleWidth      =   15060
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdre 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Re Order"
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
      Left            =   12360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.ComboBox cboCategory 
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
      Left            =   7830
      TabIndex        =   12
      Top             =   405
      Visible         =   0   'False
      Width           =   3090
   End
   Begin Crystal.CrystalReport CR 
      Left            =   11175
      Top             =   315
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
      Left            =   4680
      TabIndex        =   2
      Top             =   420
      Width           =   3090
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
      Height          =   420
      Left            =   10860
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   1440
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
         Format          =   59768835
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
         Format          =   59768835
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
      Height          =   420
      Left            =   10860
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1365
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
      Height          =   420
      Left            =   10860
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   915
      Width           =   1440
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   8340
      Left            =   120
      TabIndex        =   10
      Top             =   900
      Width           =   10695
      _cx             =   18865
      _cy             =   14711
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
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
      TabIndex        =   13
      Top             =   150
      Width           =   510
   End
   Begin VB.Label unit 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   7905
      TabIndex        =   11
      Top             =   450
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
      Left            =   7830
      TabIndex        =   9
      Top             =   90
      Visible         =   0   'False
      Width           =   870
   End
End
Attribute VB_Name = "frmConsumeItemSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim save As New ADODB.Recordset
     Dim rs As New ADODB.Recordset
     Dim search As New ADODB.Recordset
     Dim search1 As New ADODB.Recordset
     Dim rss As New ADODB.Recordset
     Dim oprs As New ADODB.Recordset
     Dim Receive
     Dim Issue
     Dim purchaseret
     Dim opening
     Dim opDate As Date
     Dim quality As String
     Dim pr_rec, pr_issue
     Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

  
Private Sub cboItem_Click()
  On Error Resume Next
  fillGrid
End Sub
Private Function LoadName()
CboPName.Clear
Set rs = New ADODB.Recordset
If rs.State = 1 Then rs.Close
rs.Open "Select distinct(CourseName) from ItemCreation", CON, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
    Do While Not rs.EOF
        CboPName.AddItem rs.Fields(0)
        rs.MoveNext
    Loop
End If
End Function
Private Sub cboCategory_Click()
'LoadName
End Sub

Private Sub CboPName_Click()
    cboCategory.Clear
    If rs.State = 1 Then rs.Close
    rs.Open "select ItemName from itemcreation where CourseName='" & CboPName.Text & "'", CON
    While rs.EOF = False
       cboCategory.AddItem rs(0)
       rs.MoveNext
    Wend
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdPrint_Click()

   'Call cmdsearch_Click
  
  
   CR.Reset
   CR.Connect = constr
   
  If CboPName.Text <> "" And cboCategory.Text <> "" Then
   CR.ReportFileName = App.Path & "\Reports\consumvablestocksummary.rpt"
  Else
   CR.ReportFileName = App.Path & "\Reports\consumvablestocksummary.rpt"
  End If
   
   CR.Formulas(0) = "Fromdate='" & fromdate.Value & "'"
   CR.Formulas(1) = "Todate='" & todate.Value & "'"
   CR.Formulas(2) = "gp='" & CboPName & "'"
   CR.WindowShowCloseBtn = True
   CR.WindowShowPrintBtn = True
   CR.WindowControlBox = True
   CR.WindowShowPrintSetupBtn = True
   CR.WindowShowProgressCtls = True
   CR.WindowState = crptMaximized
   CR.Action = 1


End Sub
Sub UpdateStock()
     
     
     
     
     Screen.MousePointer = vbHourglass
      
         
         
         
    Dim rs_S As New ADODB.Recordset

    
    quality = "Consumable"
    If search.State = 1 Then search.Close
    search.Open "select distinct(Dates) from [dates] where Dates>=datevalue('" & fromdate.Value & "') and Dates<=datevalue('" & todate.Value & "')  order by Dates", CON
    If search.EOF = False Then
        
        While search.EOF = False
           
            If rs.State = 1 Then rs.Close
            If CboPName.Text <> "" Then
               rs.Open "select distinct(ItemName) from itemcreation where CourseName='Consumable' and ItemName='" & CboPName.Text & "'", CON
               Else
               rs.Open "select distinct(ItemName) from itemcreation where CourseName='Consumable'", CON
            End If
           
           If rs.EOF = False Then
                While rs.EOF = False
                      '------------Calculate Opening--------
                      op = 0
                      If oprs.State = 1 Then oprs.Close
                      oprs.Open "select Opening from itemcreation where CourseName='" & quality & "' and Itemname='" & rs.Fields(0).Value & "'", CON
                      If oprs.EOF = False Then
                            If rss.State = 1 Then rss.Close
                            rss.Open "select * from ConsumeItemStockSummary where Name='" & rs.Fields(0).Value & "'", CON
                            If rss.EOF = True Then
                               op = oprs.Fields(0).Value
                            End If
                        Else
                            op = 0
                      End If
                     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from FinishPurchase where dates<datevalue('" & fromdate.Value & "') and ItemName='" & rs.Fields(0).Value & "'", CON
                     If Not IsNull(search1.Fields(0)) Then
                        pr_rec = search1.Fields(0).Value
                        Else
                        pr_rec = 0
                     End If
                     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssueDeppt where dates<datevalue('" & fromdate.Value & "') and Itemname='" & rs.Fields(0).Value & "'", CON
                     If Not IsNull(search1.Fields(0)) Then
                        pr_issue = search1.Fields(0).Value
                        Else
                        pr_issue = 0
                     End If
                     
                     'If search1.State = 1 Then search1.Close
                     'search1.Open "select sum(Qty) from RawPurchaseReturn where invdate<datevalue('" & FromDate.Value & "') and [Group]='" & quality & "' and [Name]='" & rs.Fields(0).Value & "'", con
                     'If Not IsNull(search1.Fields(0)) Then
                     '   purchaseret = search1.Fields(0).Value
                     '   Else
                     '   purchaseret = 0
                     'End If
                     
                     
                     
                     op = (op + (pr_rec - pr_issue))
                     
                     
                     
                     '-----------------end Code-------------------

                     
                     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from FinishPurchase where dates=datevalue('" & search(0) & "') and ItemName='" & rs.Fields(0).Value & "'", CON
                     If Not IsNull(search1.Fields(0)) Then
                        Receive = search1.Fields(0).Value
                        Else
                        Receive = 0
                     End If
                     
                     
                     If search1.State = 1 Then search1.Close
                     search1.Open "select sum(Qty) from IssueDeppt where dates=datevalue('" & search(0) & "') and Itemname='" & rs.Fields(0).Value & "'", CON
                     If Not IsNull(search1.Fields(0)) Then
                        Issue = search1.Fields(0).Value
                        Else
                        Issue = 0
                     End If
               
                     
                     
'                     If search1.State = 1 Then search1.Close
'                     search1.Open "select sum(Qty) from RawFinishurchaseReturn where invdate=datevalue('" & Search(0) & "') and [Group]='" & quality & "' and [Name]='" & rs.Fields(0).Value & "'", con
'                     If Not IsNull(search1.Fields(0)) Then
'                        Finishurchaseret = search1.Fields(0).Value
'                        Else
'                        Finishurchaseret = 0
'                     End If
                     
                     
                     
        
        '--------------------- Data Fiter and Now  Save Coding========================
         opening = 0
         If Issue = 0 And Receive = 0 And op = 0 Then
            GoTo aaa:
         End If
          
          Set save = New ADODB.Recordset
          If save.State = 1 Then save.Close
          save.Open "select * from ConsumeItemStockSummary", CON, adOpenDynamic, adLockOptimistic
          save.addNew
          save!Name = rs.Fields(0).Value
          save!gpname = quality
          ss = rs.Fields(0).Value
          save!dates = search.Fields(0).Value
          save!OpenStock = op
          save!Issue = Issue
          save!ReceiveStock = Receive
          'save!PurhaseReturn = purchaseret
          save!ClosingStock = ((op + Receive) - (Issue))
          save.Update
          
          
          
          If rss.State = 1 Then rss.Close
          rss.Open "select ClosingStock,Aouto  from ConsumeItemStockSummary where Name='" & ss & "'  order by Dates,aouto", CON
          If rss.EOF = False Then
                
              If rss.RecordCount = 1 Then
                opening = rss.Fields(0).Value
              ElseIf rss.RecordCount = 2 Then
                rss.MoveLast
                rss.MovePrevious
                opening = rss.Fields(0).Value
                rss.MoveNext
                nn = rss.Fields(1).Value
              ElseIf rss.RecordCount > 2 Then
                rss.MoveLast
                rss.MovePrevious
                opening = rss.Fields(0).Value
                rss.MoveNext
                nn = rss.Fields(1).Value
              
              End If
              
              
             If nn <> "" Then
              Set save = New ADODB.Recordset
              If save.State = 1 Then save.Close
              save.Open "select * from ConsumeItemStockSummary where Aouto=" & nn & "", CON, adOpenDynamic, adLockOptimistic
              If save.EOF = False Then
                 save!OpenStock = opening
                 save!ClosingStock = ((opening + Receive) - (Issue))
                 save.Update
                 opening = 0
                 nn = ""
              End If
             End If
             
             Else
              
              
          End If
aaa:
                                       
                                       
                                   rs.MoveNext
                                  Wend
                                  End If
                                   
                                   
                                   
                                   search.MoveNext
                                   Wend
    
    
    Else
    
    
    
    
              '------------------- Opening Show No Transiction

                                
                                
                                If rs.State = 1 Then rs.Close
                                rs.Open "select distinct(Name) from ConsumeItemStockSummary where GpName='" & quality & "'", CON
                                If rs.EOF = False Then
                                While rs.EOF = False
                                
                                                '------------Calculate Opening--------
                                op = 0
                                If oprs.State = 1 Then oprs.Close
                                    oprs.Open "select OpeningStock,Openingdate from itemcreation where CourseName='" & quality & "' and ItemName='" & rs.Fields(0).Value & "'", CON
                                    If oprs.EOF = False Then
                                        If rss.State = 1 Then rss.Close
                                            rss.Open "select * from ConsumeItemStockSummary where GpName='" & quality & "' and Name='" & rs.Fields(0).Value & "'", CON
                                            If rss.EOF = True Then
                                                op = oprs.Fields(0).Value
                                            End If
                                            Else
                                        op = 0
                                    End If
                                
                            
                            If search1.State = 1 Then search1.Close
                            search1.Open "select sum(Qty) from Finishurchase where dates<datevalue('" & search(0) & "') and ItemName='" & rs.Fields(0).Value & "'", CON
                            If Not IsNull(search1.Fields(0)) Then
                            Receive = search1.Fields(0).Value
                            Else
                            Receive = 0
                            End If
                            
                            
                            If search1.State = 1 Then search1.Close
                            search1.Open "select sum(Qty) from IssueDeppt where dates<datevalue('" & search(0) & "') and Item='" & rs.Fields(0).Value & "'", CON
                            If Not IsNull(search1.Fields(0)) Then
                            Issue = search1.Fields(0).Value
                            Else
                            Issue = 0
                            End If
           
                                                           
'
'                                    If search1.State = 1 Then search1.Close
'                                    search1.Open "select sum(Qty) from RawPurchaseReturn where invdate<datevalue('" & FromDate.Value & "') and [Group]='" & quality & "' and [Name]='" & rs.Fields(0).Value & "'", con
'                                    If Not IsNull(search1.Fields(0)) Then
'                                    purchaseret = search1.Fields(0).Value
'                                    Else
'                                    purchaseret = 0
'                                    End If
                                                           
                                                           
                                                           
                                                                 
                                         op = ((op + pr_rec) - pr_issue)

                                         If op = 0 Then GoTo ab1:
                                         
                                         Set save = New ADODB.Recordset
                                         If save.State = 1 Then save.Close
                                         save.Open "select * from ConsumeItemStockSummary", CON, adOpenDynamic, adLockOptimistic
                                         save.addNew
                                         save!GroupName = quality
                                         save!gpname = quality
                                         save!Name = rs.Fields(0).Value
                                         'save!unit = unit
                                         ss = rs.Fields(0).Value
                                         save!dates = fromdate.Value
                                         save!OpenStock = op
                                         save!ReceiveStock = Receive
                                         save!Issue = Issue
                                         'save!PurhaseReturn = purchaseret
                                         save!ClosingStock = ((save!OpenStock + Receive) - (Issue))
                                         save.Update
                                         
                                         
ab1:

                                       rs.MoveNext
                                   Wend
                                   End If
  
 '==============

    
      
    
End If
            

  
   
Screen.MousePointer = vbDefault

End Sub
Sub AddDate()
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim K As Integer
    Dim d As Date
    Set cmd.ActiveConnection = CON
    
    cmd.CommandText = "delete from dates"
    cmd.Execute
    K = DateDiff("d", fromdate.Value, todate.Value)
    If rs.State = 1 Then rs.Close
    rs.Open "select * from dates", CON, adOpenDynamic, adLockOptimistic
     d = fromdate.Value
     For i = 0 To K
           rs.addNew
           rs!dates = d
           rs.Update
           d = d + 1
       Next
    
End Sub

Private Sub cmdre_Click()
   
''   CR.Reset
''   CR.Connect = "filedsn=stockdsn;pwd=java;"
''   CR.ReportFileName = App.Path & "\consumvableReorder.rpt"
''   CR.ReplaceSelectionFormula "{reorder.gpname}='" & cboCategory.Text & "' and {reorder.reorder}>0"
''   CR.Formulas(0) = "Fromdate='" & FromDate.Value & "'"
''   CR.Formulas(1) = "Todate='" & ToDate.Value & "'"
''   CR.WindowShowCloseBtn = True
''   CR.WindowShowPrintBtn = True
''   CR.WindowControlBox = True
''   CR.WindowShowPrintSetupBtn = True
''   CR.WindowShowProgressCtls = True
''   CR.WindowState = crptMaximized
''   CR.Action = 1

End Sub

Sub showitem()
    
    Screen.MousePointer = vbHourglass
    Dim opening As Long
    Dim search As New ADODB.Recordset
    Dim save As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Set cmd.ActiveConnection = CON
    cmd.CommandText = "delete from ConsumeItemStockSummary"
    cmd.Execute
    
    '============================
    CON.Execute "delete from temp"
     
    '=================
   
    
    If Me.cboCategory.Text <> "" Then
        
        CON.Execute "insert into temp(gp,itemname,dates,head) select coursename,itemname,dates,head from trans where (dates>=datevalue('" & fromdate.Value & "') and dates<=datevalue('" & todate.Value & "')) and itemname='" & Me.cboCategory.Text & "'"
        CON.Execute "insert into temp(gp,itemname,dates,head) select coursename,itemname,dates,head from trans1 where (dates>=datevalue('" & fromdate.Value & "') and dates<=datevalue('" & todate.Value & "')) and itemname='" & Me.cboCategory.Text & "'"
        'con.Execute "insert into temp(gp,itemname,dates,head) select coursename,itemname,dates,head from trans2 where (dates>=datevalue('" & FromDate.Value & "') and dates<=datevalue('" & ToDate.Value & "')) and (head= 'Non Consume') and itemname='" & Me.cboCategory.Text & "'"
        'con.Execute "insert into temp(gp,itemname,dates,head) select coursename,itemname,dates,head from trans3 where (dates>=datevalue('" & FromDate.Value & "') and dates<=datevalue('" & ToDate.Value & "')) and (head= 'Non Consume') and itemname='" & Me.cboCategory.Text & "'"
    
    End If
      
    If rs.State = 1 Then rs.Close
    rs.Open "select ItemName,coursename from itemcreation where  itemname='" & cboCategory.Text & "'", CON
    If rs.EOF = False Then
        quality = rs.Fields(1).Value
        op = 0
        If oprs.State = 1 Then oprs.Close
        oprs.Open "select Opening from itemcreation where Itemname='" & rs.Fields(0).Value & "'", CON
        If oprs.EOF = False Then
           op = oprs.Fields(0).Value
        Else
          op = 0
        End If
        If search1.State = 1 Then search1.Close
        search1.Open "select sum(Qty) from FinishPurchase where dates<datevalue('" & fromdate.Value & "') and (ItemName)='" & (rs.Fields(0).Value) & "'", CON
        If Not IsNull(search1.Fields(0)) Then
           pr_rec = search1.Fields(0).Value
           Else
           pr_rec = 0
        End If
        If search1.State = 1 Then search1.Close
        search1.Open "select sum(Qty) from IssueDeppt where dates<datevalue('" & fromdate.Value & "') and Itemname='" & rs.Fields(0).Value & "'", CON
        If Not IsNull(search1.Fields(0)) Then
           pr_issue = search1.Fields(0).Value
           Else
           pr_issue = 0
        End If
        
         
        'If search1.State = 1 Then search1.Close
        'search1.Open "select sum(Qty) from ReturnConsumable where dates<datevalue('" & FromDate.Value & "') and Itemname='" & rs.Fields(0).Value & "'", con
       ' If Not IsNull(search1.Fields(0)) Then
        '   ret = search1.Fields(0).Value
         '  Else
         '  ret = 0
        'End If

        
        
        
        'If search1.State = 1 Then search1.Close
        'search1.Open "select sum(Qty) from DamageEntry where dates<datevalue('" & FromDate.Value & "') and Itemname='" & rs.Fields(0).Value & "'", con
       ' If Not IsNull(search1.Fields(0)) Then
        '   dem = search1.Fields(0).Value
        '   Else
        '   dem = 0
        'End If
        
        
        
        op = ((op + pr_rec) - (pr_issue))
                     
       End If
      
      
    Dim b1 As Boolean
    Dim cl
    cl = 0
    
    b1 = False
   
      
    If rs.State = 1 Then rs.Close
    rs.Open "select distinct(dates) from temp where itemname='" & cboCategory.Text & "' order by dates", CON
    If rs.EOF = False Then
    
    If rs.EOF = True Then
       MsgBox "No Transaction Done !", vbInformation
       Screen.MousePointer = vbDefault
       Exit Sub
    End If
  
    
    While rs.EOF = False
        If search1.State = 1 Then search1.Close
        search1.Open "select sum(Qty) from FinishPurchase where dates=datevalue('" & rs.Fields(0).Value & "') and ItemName='" & cboCategory.Text & "'", CON
        If Not IsNull(search1.Fields(0)) Then
           pr_rec = search1.Fields(0).Value
         Else
           pr_rec = 0
        End If
        If search1.State = 1 Then search1.Close
        search1.Open "select sum(Qty) from IssueDeppt where dates=datevalue('" & rs.Fields(0).Value & "') and Itemname='" & cboCategory.Text & "'", CON
        If Not IsNull(search1.Fields(0)) Then
           pr_issue = search1.Fields(0).Value
           Else
           pr_issue = 0
        End If
        
        
        
        'f search1.State = 1 Then search1.Close
       ' search1.Open "select sum(Qty) from ReturnConsumable where dates=datevalue('" & rs.Fields(0).Value & "') and Itemname='" & cboCategory.Text & "'", con
       ' If Not IsNull(search1.Fields(0)) Then
       '   ret = search1.Fields(0).Value
        '   Else
        '   ret = 0
       ' End If

        
        
        
       ' If search1.State = 1 Then search1.Close
       ' search1.Open "select sum(Qty) from DamageEntry where dates=datevalue('" & rs.Fields(0).Value & "') and Itemname='" & cboCategory.Text & "'", con
       ' If Not IsNull(search1.Fields(0)) Then
       '    dem = search1.Fields(0).Value
       '    Else
       '    dem = 0
       ' End If
        
        
        
        '--------------------- Data Fiter and Now  Save Coding========================
        
        If save.State = 1 Then save.Close
        save.Open "select * from  ConsumeItemStockSummary", CON, adOpenDynamic, adLockOptimistic
        save.addNew
        save!gpname = quality
        save!Name = cboCategory.Text
        save!dates = rs.Fields(0).Value
        save!ReceiveStock = pr_rec
        save!Issue = pr_issue
        'save!ret = ret
        'save!demage = dem
        If b1 = False Then
           save!OpenStock = op
           save!ClosingStock = ((op + pr_rec) - (pr_issue))
           cl = ((op + pr_rec) - (pr_issue))
        Else
           save!OpenStock = cl
           save!ClosingStock = ((cl + pr_rec) - (pr_issue))
           cl = ((cl + pr_rec) - (pr_issue))
        End If
        save.Update
        b1 = True
    rs.MoveNext
    Wend
    
    
    
    
    Else
       
        If op = 0 Then
        Screen.MousePointer = vbDefault
          Exit Sub
          
        End If
        
        If save.State = 1 Then save.Close
        save.Open "select * from ConsumeItemStockSummary", CON, adOpenDynamic, adLockOptimistic
        save.addNew
        save!gpname = quality
        save!Name = cboCategory.Text
        save!dates = fromdate.Value
        save!ReceiveStock = 0
        save!Issue = 0
        'save!ret = 0
        'save!demage = 0

        save!OpenStock = op
        save!ClosingStock = op
        save.Update
        
    End If
     
    'fillgrid
    '=============================
    
    
    
    
                vs.Cols = 6
               vs.FormatString = "Name|^date|>OpenStock|>Purchase|>Issue|>ClosingStock"
                
                Dim fill As New ADODB.Recordset
                If fill.State = 1 Then fill.Close
                fill.Open "select Name,dates,OpenStock,ReceiveStock,Issue,ClosingStock  from ConsumeItemStockSummary order by Dates,Aouto", CON
                If fill.EOF = False Then
                   vs.Rows = fill.RecordCount + 1
                   For i = 1 To fill.RecordCount
                     If fill.EOF = False Then
                         vs.TextMatrix(i, 0) = fill.Fields(0).Value
                        vs.TextMatrix(i, 1) = fill.Fields(1).Value
                        vs.TextMatrix(i, 2) = fill.Fields(2).Value
                        vs.TextMatrix(i, 3) = fill.Fields(3).Value
                        vs.TextMatrix(i, 4) = fill.Fields(4).Value
                        vs.TextMatrix(i, 5) = fill.Fields(5).Value
                      End If
                     fill.MoveNext
                   Next
                  
              End If
                 
       vs.ColWidth(0) = 2500
       vs.ColWidth(1) = 1100
       vs.ColWidth(2) = 1600
       vs.ColWidth(3) = 1600
       vs.ColWidth(4) = 1700
     

    
    
    
    
    
    
    
    '===================================

     Screen.MousePointer = vbDefault
End Sub
Private Sub cmdSearch_Click()

Dim unit As String

If DateValue(fromdate.Value) > DateValue(todate.Value) Then
   MsgBox "Invalid Month Selection..", vbCritical
   Exit Sub
End If


If CboPName.Text <> "" And cboCategory.Text <> "" Then
 showitem
 Exit Sub
End If

Screen.MousePointer = vbHourglass
Dim opening As Long
Dim search As New ADODB.Recordset
Dim save As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
opening = 0

CON.Execute "delete from ConsumeItemStockSummary"

If rs.State = 1 Then rs.Close
If CboPName.Text <> "" Then
rs.Open "select * from ItemCreation where CourseName='" & CboPName.Text & "'", CON
Else
rs.Open "select * from ItemCreation", CON
End If



If save.State = 1 Then save.Close
save.Open "select * from ConsumeItemStockSummary", CON, adOpenDynamic, adLockOptimistic


If rs.EOF = False Then
vs.Rows = rs.RecordCount + 1

For i = 1 To rs.RecordCount

unit = rs!unit

Receive = 0
opening = 0
Issue = 0

If rs1.State = 1 Then rs1.Close
rs1.Open "select Opening from itemcreation where itemname='" & rs!itemname & "'", CON
If rs1.EOF = False Then
opening = rs1(0)
End If

If rs1.State = 1 Then rs1.Close
rs1.Open "select sum(Qty) from FinishPurchase where convert(smalldatetime,Dates,103)<convert(smalldatetime,'" & fromdate.Value & "',103) and itemname='" & rs!itemname & "'", CON, adOpenKeyset, adLockReadOnly
If Not IsNull(rs1(0)) Then
opening = opening + rs1(0)
End If

If rs1.State = 1 Then rs1.Close
rs1.Open "select sum(Qty) from IssueDeppt where convert(smalldatetime,Dates,103)<convert(smalldatetime,'" & fromdate.Value & "',103) and itemname='" & rs!itemname & "'", CON, adOpenKeyset, adLockReadOnly
If Not IsNull(rs1(0)) Then
opening = opening - rs1(0)
End If



If rs1.State = 1 Then rs1.Close
rs1.Open "select sum(Qty) from FinishPurchase where (convert(smalldatetime,Dates,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) and convert(smalldatetime,Dates,103)<=convert(smalldatetime,'" & todate.Value & "',103)) and itemname='" & rs!itemname & "'", CON, adOpenKeyset, adLockReadOnly
If Not IsNull(rs1(0)) Then
Receive = rs1(0)
End If

If rs1.State = 1 Then rs1.Close
rs1.Open "select sum(Qty) from IssueDeppt where (convert(smalldatetime,Dates,103)>=convert(smalldatetime,'" & fromdate.Value & "',103) and convert(smalldatetime,Dates,103)<=convert(smalldatetime,'" & todate.Value & "',103)) and itemname='" & rs!itemname & "'", CON, adOpenKeyset, adLockReadOnly
If Not IsNull(rs1(0)) Then
Issue = rs1(0)
End If



save.addNew
save!unit = unit
save!Name = rs!itemname
save!ReceiveStock = Receive
save!Issue = Issue
save!OpenStock = opening
save!ClosingStock = ((opening + Receive) - Issue)
save.Update
rs.MoveNext


opening = 0

Next




End If











fillGrid
Screen.MousePointer = vbDefault
End Sub
Sub MultipleItem()
Dim kk As Integer
kk = 0
If rs.State = 1 Then rs.Close
rs.Open "select distinct(ItemName),coursename from itemcreation where CourseName='" & cboCategory.Text & "'", CON
If rs.EOF = False Then kk = rs.RecordCount
While rs.EOF = False
quality = rs.Fields(1).Value
op = 0
If oprs.State = 1 Then oprs.Close
oprs.Open "select Opening from itemcreation where coursename='" & quality & "' and Itemname='" & rs.Fields(0).Value & "'", CON
If oprs.EOF = False Then op = oprs.Fields(0).Value Else op = 0
If search1.State = 1 Then search1.Close
search1.Open "select sum(Qty) from FinishPurchase where dates<datevalue('" & fromdate.Value & "') and (ItemName)='" & (rs.Fields(0).Value) & "'", CON
If Not IsNull(search1.Fields(0)) Then
pr_rec = search1.Fields(0).Value
Else
pr_rec = 0
End If
If search1.State = 1 Then search1.Close
search1.Open "select sum(Qty) from IssueDeppt where dates<datevalue('" & fromdate.Value & "') and Itemname='" & rs.Fields(0).Value & "'", CON
If Not IsNull(search1.Fields(0)) Then
pr_issue = search1.Fields(0).Value
Else
pr_issue = 0
End If
op = (op + (pr_rec - pr_issue))
'-------------------- Data Fiter and Now  Save Coding========================
If search1.State = 1 Then search1.Close
search1.Open "select sum(Qty) from FinishPurchase where (dates>=datevalue('" & fromdate.Value & "') and dates<=datevalue('" & todate.Value & "')) and ItemName='" & (rs.Fields(0).Value) & "'", CON
If Not IsNull(search1.Fields(0)) Then
Receive = search1.Fields(0).Value
Else
Receive = 0
End If
If search1.State = 1 Then search1.Close
search1.Open "select sum(Qty) from IssueDeppt where (dates>=datevalue('" & fromdate.Value & "') and dates<=datevalue('" & todate.Value & "')) and ItemName='" & (rs.Fields(0).Value) & "'", CON
If Not IsNull(search1.Fields(0)) Then
Issue = search1.Fields(0).Value
Else
Issue = 0
End If
If op = 0 And Receive = 0 And Issue = 0 Then GoTo a10
Set save = New ADODB.Recordset
If save.State = 1 Then save.Close
save.Open "select * from ConsumeItemStockSummary", CON, adOpenDynamic, adLockOptimistic
save.addNew
save!gpname = quality
save!Name = rs.Fields(0).Value
save!dates = fromdate.Value
save!OpenStock = op
save!ClosingStock = (op + (Receive - Issue))
save!ReceiveStock = Receive
save!Issue = Issue
save.Update
a10:
rs.MoveNext
Wend
fillGrid
Screen.MousePointer = vbDefault
End Sub
Sub fillGrid()
    vs.Cols = 5
    vs.FormatString = "Item Name|>Opening|>Receipt|>Issue|>Balance"
    
    Dim fill As New ADODB.Recordset
    If fill.State = 1 Then fill.Close
    'fill.Open "select Name,OpenStock,ReceiveStock,Issue,ClosingStock  from ConsumeItemStockSummary where Dates>=datevalue('" & FromDate.Value & "') and Dates<=datevalue('" & ToDate.Value & "') order by Dates,Aouto", con
    fill.Open "select Name,OpenStock,ReceiveStock,Issue,ClosingStock  from ConsumeItemStockSummary order by Dates,Aouto", CON
    If fill.EOF = False Then
       
       
       '-----------------------------------------
       
       vs.Rows = fill.RecordCount + 1
       For i = 1 To fill.RecordCount
         If fill.EOF = False Then
         
            vs.TextMatrix(i, 0) = fill.Fields(0).Value
            vs.TextMatrix(i, 1) = fill.Fields(1).Value
            vs.TextMatrix(i, 2) = fill.Fields(2).Value
            vs.TextMatrix(i, 3) = fill.Fields(3).Value
            vs.TextMatrix(i, 4) = fill.Fields(4).Value
            
         
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
       
       vs.ColWidth(0) = 3000
       vs.ColWidth(1) = 1000
       vs.ColWidth(2) = 1400
       vs.ColWidth(3) = 1400
       vs.ColWidth(4) = 1700
     
       
 
End Sub
Sub AddCategory()
    If rs.State = 1 Then rs.Close
    rs.Open "select distinct(Name) from ProductMaster", CON
    cboCategory.Clear
    If rs.EOF = False Then
       While rs.EOF = False
          cboCategory.AddItem rs.Fields(0).Value
          rs.MoveNext
       Wend
    End If
    LoadName
    
End Sub


Private Sub Form_Load()
todate.Value = Date
'FillGrid
fromdate.Value = Date

AddCategory
End Sub
Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then todate.SetFocus
End Sub
