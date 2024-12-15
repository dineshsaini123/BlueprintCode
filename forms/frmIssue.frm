VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmIssue 
   Caption         =   "Book Issue "
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbogodown 
      Height          =   315
      Left            =   1755
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   945
      Width           =   960
   End
   Begin VB.TextBox txtLoose 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      TabIndex        =   31
      Top             =   6540
      Width           =   1215
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3420
      TabIndex        =   30
      Top             =   6480
      Width           =   1455
   End
   Begin Crystal.CrystalReport CR 
      Left            =   7860
      Top             =   7740
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   300
      TabIndex        =   21
      Top             =   6960
      Width           =   7290
      Begin VB.CommandButton cmdExit_12 
         Caption         =   "E&xit"
         Height          =   480
         Left            =   6120
         TabIndex        =   28
         Top             =   255
         Width           =   1005
      End
      Begin VB.CommandButton cmdPrint_7 
         Caption         =   "&Print"
         Height          =   480
         Left            =   5100
         TabIndex        =   27
         Top             =   255
         Width           =   1005
      End
      Begin VB.CommandButton cmdUndo_5 
         Caption         =   "&Undo"
         Height          =   480
         Left            =   4095
         TabIndex        =   26
         Top             =   255
         Width           =   1005
      End
      Begin VB.CommandButton cmdEdit_4 
         Caption         =   "&Edit"
         Height          =   480
         Left            =   3090
         TabIndex        =   25
         Top             =   255
         Width           =   1005
      End
      Begin VB.CommandButton cmdDelete_3 
         Caption         =   "&Delete"
         Height          =   480
         Left            =   2085
         TabIndex        =   24
         Top             =   255
         Width           =   1005
      End
      Begin VB.CommandButton cmdSave_2 
         Caption         =   "&Save"
         Height          =   480
         Left            =   1080
         TabIndex        =   23
         Top             =   255
         Width           =   1005
      End
      Begin VB.CommandButton cmdAdd_1 
         Caption         =   "&Add"
         Height          =   480
         Left            =   75
         TabIndex        =   22
         Top             =   255
         Width           =   1005
      End
   End
   Begin MSComCtl2.DTPicker Dates 
      Height          =   315
      Left            =   1740
      TabIndex        =   20
      Top             =   645
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   19791873
      CurrentDate     =   39500
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7320
      TabIndex        =   17
      Top             =   945
      Width           =   3690
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7320
      TabIndex        =   16
      Top             =   645
      Width           =   3690
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   -135
      Top             =   8910
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "filedsn=saru;"
      OLEDBString     =   "filedsn=saru;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ItemMaster"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Output Weight"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   75
      TabIndex        =   11
      Top             =   8715
      Visible         =   0   'False
      Width           =   465
      Begin VB.TextBox txtRawAndCasting 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3690
         TabIndex        =   12
         Text            =   "0"
         Top             =   1035
         Width           =   1320
      End
      Begin VB.Shape Shape2 
         Height          =   585
         Left            =   75
         Top             =   1365
         Width           =   3150
      End
      Begin VB.Label Label9 
         Caption         =   "Receiving Semi Finish &&  Finish"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   15
         Top             =   1515
         Width           =   3060
      End
      Begin VB.Shape Shape1 
         Height          =   615
         Left            =   0
         Top             =   1590
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Receiving  from casting"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   14
         Top             =   885
         Width           =   2325
      End
      Begin VB.Label Label1 
         Caption         =   "Raw Issue Weight"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   135
         TabIndex        =   13
         Top             =   570
         Width           =   1635
      End
   End
   Begin VB.TextBox txtTotal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7650
      TabIndex        =   4
      Text            =   "0"
      Top             =   6540
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1740
      TabIndex        =   2
      Top             =   1320
      Width           =   4065
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7320
      TabIndex        =   1
      Top             =   315
      Width           =   3690
   End
   Begin VB.TextBox txtHeating 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1740
      TabIndex        =   0
      Top             =   330
      Width           =   1740
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   4725
      Left            =   180
      TabIndex        =   3
      Top             =   1740
      Width           =   11685
      _cx             =   20611
      _cy             =   8334
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
      GridColor       =   8388608
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
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmIssue.frx":0000
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
      Begin VB.Frame VsFrame 
         Height          =   2370
         Left            =   0
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   4155
         Begin MSDataListLib.DataCombo cboItem 
            Height          =   2310
            Left            =   0
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   0
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   4075
            _Version        =   393216
            Appearance      =   0
            Style           =   1
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Godown Name"
      Height          =   270
      Index           =   7
      Left            =   225
      TabIndex        =   33
      Top             =   990
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Search Challan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   29
      Top             =   45
      Width           =   2505
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Search Binder"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7320
      TabIndex        =   10
      Top             =   45
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Total :"
      Height          =   195
      Index           =   6
      Left            =   2340
      TabIndex        =   9
      Top             =   6540
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "Remarks "
      Height          =   300
      Index           =   4
      Left            =   225
      TabIndex        =   8
      Top             =   1275
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Date "
      Height          =   270
      Index           =   1
      Left            =   195
      TabIndex        =   7
      Top             =   660
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Binder Name "
      Height          =   300
      Index           =   2
      Left            =   6120
      TabIndex        =   6
      Top             =   345
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Challan No :"
      Height          =   270
      Index           =   0
      Left            =   210
      TabIndex        =   5
      Top             =   345
      Width           =   1530
   End
End
Attribute VB_Name = "frmIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itemgp As String
Dim rates As Double
Dim i As Integer
Dim Status As String
Dim Item_Name As String
Dim unit As String
Dim qty As Integer
Dim iitem1 As String
Dim StockFlag As String

Private Sub cmdMain_Click()
Unload Me
End Sub
Sub cellposi()
 'VsFrame.Width = 3165
 VsFrame.Top = vs.Top + ((vs.CellTop)) - 1400
 VsFrame.Left = (vs.Left) - 200
End Sub
Sub total()
txtTotal.Text = 0
txtLoose.Text = 0

For j = 1 To vs.Rows - 1
If vs.TextMatrix(j, 0) <> "" Then
txtTotal.Text = (Val(txtTotal.Text) + Val(vs.TextMatrix(j, 1)))
txtLoose.Text = (Val(txtLoose.Text) + Val(vs.TextMatrix(j, 3)))
End If
Next

End Sub

Sub cellposiVs()
 Vs1Frame.Width = 2500
 Vs1Frame.Top = vs1.Top + ((vs1.CellTop))
 Vs1Frame.Left = (vs1.Left) + 550
End Sub
Sub AddItemInGrid1()
'
'    Dim rs_4 As New ADODB.Recordset
'
'    rs_4.Open "select * from bm order by BKDESC", con, adOpenDynamic, adLockOptimistic
'
'    Set cboitemvs1.RowSource = rs_4
'    cboitemvs1.ListField = "BKDESC"
'    cboitemvs1.BoundColumn = "BKCODE"
'    cboitemvs1.ReFill
    
End Sub
Sub AddItemInGrid3()
    'Adodc1.ConnectionString = "filedsn=Saru"
    'Adodc1.CommandType = adCmdText
'
'    Dim rs_3 As New ADODB.Recordset
'
'    'rs_3.Open "select * from ItemMaster where (ItemGp='Finish Item' or ItemGp='Scrap' or ItemGp='Losses' or (ItemGp='Chunnus (Pointing)' or ItemGp='Chunnus (Sonicut)') or  Itemgp='Rejuction') order by ItemName", con, adOpenDynamic, adLockOptimistic
'    rs_3.Open "select * from ItemMaster order by ItemName", con, adOpenDynamic, adLockOptimistic
'
'    'Adodc1.Refresh
'    Set cboItemVs3.RowSource = rs_3
'    cboItemVs3.ListField = "ItemName"
'    cboItemVs3.BoundColumn = "ItemName"
'    cboItemVs3.ReFill
    
End Sub
Sub AddItemInGrid2()
'    'Adodc1.ConnectionString = "filedsn=Saru"
'    'Adodc1.CommandType = adCmdText
'    Dim rs_2 As New ADODB.Recordset
'
'    'rs_2.Open "select * from ItemMaster where (ItemGp='Semi Finish (R/D)' or ItemGp= 'Semi Finish (Store)' or Itemgp='Rejuction') order by ItemName", con, adOpenDynamic, adLockOptimistic
'    rs_2.Open "select * from ItemMaster order by ItemName", con, adOpenDynamic, adLockOptimistic
'
'    Set cboItemVs2.RowSource = rs_2
'    cboItemVs2.ListField = "ItemName"
'    cboItemVs2.BoundColumn = "ItemName"
'    cboItemVs2.ReFill
    
End Sub
Sub AddItemInGrid()
    'Adodc1.ConnectionString = "filedsn=Saru"
    'Adodc1.CommandType = adCmdText
    Dim rs_1 As New ADODB.Recordset
    
    'rs_1.Open "select * from ItemMaster where ItemGp='Raw Item' or ItemGp='Scrap' or (ItemGp='Chunnus (Pointing)' or ItemGp='Chunnus (Sonicut)') order by ItemName", con, adOpenDynamic, adLockOptimistic
    rs_1.Open "select * from bm order by BKDESC", con, adOpenDynamic, adLockOptimistic
    
    Set cboItem.RowSource = rs_1
    cboItem.ListField = "BKDESC"
    cboItem.BoundColumn = "BKCODE"
    cboItem.ReFill
    
End Sub
Private Sub cboFinishItem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtQty.SetFocus
End Sub


Private Sub cbogodown_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then txtRemarks.SetFocus
End Sub

Private Sub cbogodown_LostFocus()
If cbogodown = "" Then
   MsgBox "Select Godown Name ..", vbCritical
   cbogodown.SetFocus
   Exit Sub
End If
End Sub

Private Sub cboItem_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        cellposi
        If cboItem.Text = "" Then
        VsFrame.Visible = False
        cmdSave_2.SetFocus
        Exit Sub
        End If
        vs.TextMatrix(vs.RowSel, 0) = cboItem.Text
        vs.TextMatrix(vs.RowSel, 6) = cboItem.BoundText
        
        
        vs.SetFocus
        
     ElseIf KeyCode = 27 Then
       
          VsFrame.Visible = False
        
     End If
End Sub
Sub saveInMaster()
         On Error Resume Next
      
         If RS.State = 1 Then RS.Close
         RS.Open "select * from ItemMaster where ItemName='" & iitem1 & "'", con, adOpenDynamic, adLockOptimistic
         If RS.EOF = True Then
            RS.AddNew
            RS.Fields("ItemGp").Value = frmAddMaster.cboGp.Text
            RS.Fields("ItemName").Value = iitem1
            RS.Fields("Unit").Value = "Kg"
            RS.Update
         Else
            MsgBox "This Item Already Exist !!", vbCritical
            Exit Sub
         End If
         frmAddMaster.Visible = False
  
End Sub
Private Sub cboitemvs1_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        cellposiVs
        vs1.TextMatrix(vs1.RowSel, 0) = cboitemvs1.Text
        
        If RS.State = 1 Then RS.Close
        RS.Open "select * from ItemMaster where ItemName='" & cboitemvs1.Text & "'", con
        If RS.EOF = True Then
            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
                frmAddMaster.Show 1
                iitem1 = cboitemvs1.Text
                saveInMaster
                cboitemvs1.Text = ""
                Vs1Frame.Visible = False
                vs1.SetFocus
             End If
        End If
        vs1.SetFocus
     ElseIf KeyCode = 27 Then
        Vs1Frame.Visible = False
     End If
End Sub


Private Sub cboItemVs2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        
        
        'cellposiVs2
        vs3.TextMatrix(vs3.RowSel, 0) = cboItemVs2.Text
        Set RS = New ADODB.Recordset
        If RS.State = 1 Then RS.Close
        RS.Open "select * from ItemMaster where ItemName='" & cboItemVs2.Text & "'", con
        If RS.EOF = True Then
            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
                frmAddMaster.Show 1
                iitem1 = cboItemVs2.Text
                saveInMaster
                
                cboItemVs2.Text = ""
             End If
        End If
        vs3.SetFocus
        
ElseIf KeyCode = 27 Then
         FrameVs2.Visible = False
End If
End Sub

Private Sub cboItemVs3_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        'cellposiVs3
        
        vs2.TextMatrix(vs2.RowSel, 0) = cboItemVs3.Text
        Set RS = New ADODB.Recordset
        If RS.State = 1 Then RS.Close
        RS.Open "select * from ItemMaster where ItemName='" & cboItemVs3.Text & "'", con
        If RS.EOF = True Then
            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
                Vs3Frame.Visible = False
                iitem1 = cboItemVs3.Text
                frmAddMaster.Show 1
                saveInMaster
                cboItemVs3.Text = ""
                vs2.SetFocus
             End If
        End If
        Vs3Frame.Visible = False
        'cboItemVs3.Visible = False
        vs2.SetFocus
     ElseIf KeyCode = 27 Then
        Vs3Frame.Visible = False
     End If

End Sub

Private Sub cmdadd_Click()
 If RS.State = 1 Then RS.Close
 RS.Open "select HeatingNo from IssueMaster where HeatingDate >=datevalue('" & FromDate.Value & "') and HeatingDate <=datevalue('" & ToDate.Value & "') order by HeatingNo", con
 ListHeatingNo.Clear
 If RS.EOF = False Then
    While RS.EOF = False
       ListHeatingNo.AddItem RS(0)
       RS.MoveNext
    Wend
 End If
End Sub
Private Sub cmdDelete_Click()
   
   If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
      
      DeleteRecord txtHeating.Text, "HeatingNo", "IssueMaster"
      DeleteRecord txtHeating.Text, "HeatingNo", "IssueRawMetrial"
      Call cmdRef_Click
      
   End If
End Sub
Sub DeleteStock()
    
'''Dim rr As New ADODB.Recordset
'''Dim rs_u As New ADODB.Recordset
'''Dim openning As Double
'''
'''
'''
'''
'''
''''================ Issue For Casting
'''
'''
''' If StockFlag = "1" Then
'''
'''    If rs_u.State = 1 Then rs_u.Close
'''    rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
'''    If rs_u.EOF = False Then
'''        rs_u!qty = rs_u!qty + qty
'''        rs_u.Update
'''    End If
'''
''' End If
'''
'''
'''
''' '================ Receive For Casting
'''
''' If StockFlag = "2" Then
'''
'''   If itemgp <> "Losses" Then
'''
'''        If rs_u.State = 1 Then rs_u.Close
'''        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
'''        If rs_u.EOF = False Then
'''            rs_u!qty = rs_u!qty - qty
'''            rs_u.Update
'''        End If
'''
'''    End If
'''
''' End If
'''
'''
'''
''' '================ Receive For Finish
'''
''' If StockFlag = "3" Then
'''
'''   If itemgp <> "Losses" Then
'''
'''        If rs_u.State = 1 Then rs_u.Close
'''        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
'''        If rs_u.EOF = False Then
'''             rs_u!qty = rs_u!qty - qty
'''            rs_u.Update
'''        End If
'''
'''    End If
'''
''' End If
'''
'''
'''
''' '================ Issue For Finish
'''
''' If StockFlag = "4" Then
'''
'''   If itemgp <> "Losses" Then
'''
'''        If rs_u.State = 1 Then rs_u.Close
'''        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
'''        If rs_u.EOF = False Then
'''             rs_u!qty = rs_u!qty + qty
'''            rs_u.Update
'''        End If
'''
'''    End If
'''
''' End If
'''
''' '====================================
'''
'''
'''
'''
'''
    
   
   
    
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFatch_Click()
AddSemifinish
'Total4

End Sub

Private Sub cmdFind_Click()
 Frame1.Visible = True
 FromDate.SetFocus
End Sub

Private Sub cmdModify_Click()
   If MsgBox("Want To Modify ?", vbQuestion + vbYesNo) = vbYes Then
      
        
      DeleteRecord txtHeating.Text, "HeatingNo", "IssueMaster"
      DeleteRecord txtHeating.Text, "HeatingNo", "IssueRawMetrial"
      
      
      
      'SaveData
      
      UpdateIssue
      
      Call cmdRef_Click
   End If
End Sub
Sub UpdateIssue()

Dim rss As New ADODB.Recordset
Dim Search As New ADODB.Recordset
    
If Search.State = 1 Then Search.Close
Search.Open "select ItemName,qty from Invoice where HeatNo='" & txtHeating.Text & "'", con
If Search.EOF = False Then
While Search.EOF = False

    If rss.State = 1 Then rss.Close
    rss.Open "select * from IssueRawMetrial where HeatingNo=" & txtHeating.Text & " and ItemName='" & Search.Fields(0).Value & "'", con, adOpenDynamic, adLockOptimistic
    If rss.EOF = False Then
       rss.Fields("Issue").Value = (CDbl(rss.Fields("Issue").Value) + CDbl(Search.Fields("qty").Value))
       rss.Update
    End If
    
    Search.MoveNext
    
Wend
  
End If
  
End Sub

Private Sub cmdRef_Click()
      txtHeating.Text = ""
      txtParty.Text = ""
      
      txtRemarks.Text = ""
      
      
      txtTotal1.Text = 0
      txtTotal2.Text = 0
      txtTotal3.Text = 0
      txtTotal4.Text = 0
      
      txtSize.Text = ""
      txtGrade.Text = ""
      txtRawAndCasting.Text = 0
      
      vs.Clear
      vs1.Clear
      vs2.Clear
      vs3.Clear
      
      SetWidth
      txtHeating.SetFocus
      cmdDelete.Enabled = False
      cmdModify.Enabled = False
      cmdSave.Enabled = True
      
      Record = ""
      
End Sub


Private Sub Command4_Click()
   Unload Me
End Sub
Private Sub cmdSave_Click()
    
    
    
    
    If RS.State = 1 Then RS.Close
    RS.Open "select * from IssueMaster where HeatingNo=" & txtHeating.Text & "", con
    If RS.EOF = False Then
       MsgBox "Heating No. Already Exist !!", vbInformation
       Exit Sub
    End If
    
    If txtHeating.Text = "" Then
       MsgBox "Please Enter Heating No !!", vbCritical
       txtHeating.SetFocus
       Exit Sub
    End If
    
    
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.Open "select * from IssueMaster where HeatingNo=" & txtHeating.Text & "", con
    If RS.EOF = True Then
       If MsgBox("Want To Save ?", vbQuestion + vbYesNo) = vbYes Then
        '  SaveData
       End If
    Else
          MsgBox "Dublicate Heating No !!", vbCritical
    End If
End Sub
Sub ItemGpSearch(Str As String)
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select ItemGp,Rate from ItemMaster where ItemName='" & Str & "'", con
    If rs1.EOF = False Then
       itemgp = rs1.Fields(0).Value
       rates = rs1.Fields(1).Value
    End If
    
End Sub
Sub UpdateStock()
    Dim rr As New ADODB.Recordset
    Dim rs_u As New ADODB.Recordset
    Dim openning As Double
    
 
    
    
 '================ Issue For Casting
 
 
 If StockFlag = "1" Then
    
    If rs_u.State = 1 Then rs_u.Close
    rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
    If rs_u.EOF = True Then
        rs_u.AddNew
        rs_u!itemname = Item_Name
        ItemGpSearch Item_Name
        rs_u!itemgp = itemgp
        rs_u!unit = unit
        rs_u!Rate = rates
        rs_u!qty = (-1 * qty)
        rs_u.Update
     Else
        rs_u!qty = rs_u!qty - qty
        rs_u.Update
    End If
    
 End If
 
 '====================================
 
 
 '================ Receive For Casting
 
 If StockFlag = "2" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.Close
        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = True Then
            rs_u.AddNew
            rs_u!itemname = Item_Name
            ItemGpSearch Item_Name
            rs_u!itemgp = itemgp
            rs_u!unit = unit
            rs_u!Rate = rates
            rs_u!qty = qty
            rs_u.Update
         Else
            rs_u!qty = rs_u!qty + qty
            rs_u.Update
        End If
    
    End If
    
 End If
 
 '====================================
 
 
 '================ Receive For Finish
 
 If StockFlag = "3" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.Close
        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = True Then
            rs_u.AddNew
            rs_u!itemname = Item_Name
            ItemGpSearch Item_Name
            rs_u!itemgp = itemgp
            rs_u!unit = unit
            rs_u!Rate = rates
            rs_u!qty = qty
            rs_u.Update
         Else
            rs_u!qty = rs_u!qty + qty
            rs_u.Update
        End If
    
    End If
    
 End If
 
 '====================================
 
 
  '================ Issue For Finish
 
 If StockFlag = "4" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.Close
        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = True Then
            rs_u.AddNew
            rs_u!itemname = Item_Name
            ItemGpSearch Item_Name
            rs_u!itemgp = itemgp
            rs_u!unit = unit
            rs_u!Rate = rates
            rs_u!qty = (-1 * qty)
            rs_u.Update
         Else
            rs_u!qty = rs_u!qty - qty
            rs_u.Update
        End If
    
    End If
    
 End If
 
 '====================================
    
    
    
    
End Sub
 
 
     
  

Private Sub cmdAdd_1_Click()
   
    txtHeating.Text = ""
    Dates.Value = Date
    txtParty.Text = ""
    txtRemarks.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    txtTotal.Text = ""
    txtLoose.Text = ""

   
   RefData Me
   vs.Clear
   SetWidth
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = True
   'cmdPrint_7.Enabled = False
   cmdSave_2.Enabled = True
   cmdUndo_5.Enabled = True
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
   txtHeating.SetFocus
   txtHeating.Text = MaxSNo("invoicea", "INVOICENO")
   cbogodown.ListIndex = 0
End Sub

Private Sub cmdDelete_3_Click()
If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
con.Execute "delete from invoiceb where INVOICENO=" & txtHeating.Text & ""
con.Execute "delete from invoicea where INVOICENO=" & txtHeating.Text & ""
Call cmdAdd_1_Click
End If
End Sub

Private Sub cmdEdit_4_Click()
   cmdDelete_3.Enabled = True
   cmdEdit_4.Enabled = False
   cmdPrint_7.Enabled = True
   cmdSave_2.Enabled = True
   cmdUndo_5.Enabled = True
   cmdAdd_1.Enabled = False
   cmdExit_12.Enabled = True
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdPrint_7_Click()

CR.Reset
CR.ReportFileName = App.Path & "/Reports/CHALLAN.rpt"
CR.ReplaceSelectionFormula "{invoiceA.invoiceno}=" & txtHeating.Text & ""
CR.WindowShowPrintSetupBtn = True
CR.WindowState = crptMaximized
CR.Action = 1

End Sub

Private Sub cmdSave_2_Click()

On Error GoTo aa1


If txtParty.Text = "" Then
MsgBox "Please Enter Binder Name !!", vbInformation
Exit Sub
End If


If MsgBox("Want to Save ?", vbYesNo + vbQuestion) = vbYes Then

If RS.State = 1 Then RS.Close
RS.Open "select * from invoicea where INVOICENO=" & txtHeating.Text & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
RS.AddNew
RS.Fields("INVOICENO").Value = txtHeating.Text
RS.Fields("INVOICEDATE").Value = Dates.Value
RS.Fields("SUBLEDGER").Value = txtParty.Text
RS.Fields("GENLEDGER").Value = "Sundry Debtors"
RS.Fields("Remarks").Value = txtRemarks.Text
RS.Fields("add1").Value = Text1.Text
RS.Fields("add2").Value = Text2.Text
RS.Fields("godown").Value = cbogodown

RS.Update
Else

RS.Fields("godown").Value = cbogodown
RS.Fields("INVOICEDATE").Value = Dates.Value
RS.Fields("SUBLEDGER").Value = txtParty.Text
RS.Fields("GENLEDGER").Value = "Sundry Debtors"
RS.Fields("Remarks").Value = txtRemarks.Text
RS.Fields("add1").Value = Text1.Text
RS.Fields("add2").Value = Text2.Text
RS.Fields("NetBook").Value = Val(txtTotal1.Text)
RS.Update
cmdSave_2.Enabled = False
cmdPrint_7.SetFocus
End If



If RS.State = 1 Then RS.Close
RS.Open "select * from invoiceb where INVOICENO=" & txtHeating.Text & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then


For i = 1 To vs.Rows - 1

If vs.TextMatrix(i, 0) <> "" Then

RS.AddNew
RS.Fields("INVOICENO").Value = txtHeating.Text
RS.Fields("INVOICEDATE").Value = Dates.Value
RS.Fields("SUBLEDGER").Value = txtParty.Text
RS.Fields("GENLEDGER").Value = "Sundry Debtors"
RS.Fields("BOOKCODE").Value = vs.TextMatrix(i, 0)
RS.Fields("TBook").Value = IIf(vs.TextMatrix(i, 1) = "", 0, vs.TextMatrix(i, 1))
RS.Fields("LoosBook").Value = vs.TextMatrix(i, 2)
RS.Fields("TotalBook").Value = Val(vs.TextMatrix(i, 3))
RS.Fields("NetBook").Value = vs.TextMatrix(i, 4)
RS.Fields("Remarks").Value = vs.TextMatrix(i, 5)
RS.Fields("Book_Code").Value = vs.TextMatrix(i, 6)
RS.Update

End If

Next

Else
con.Execute "delete from invoiceb where INVOICENO=" & txtHeating.Text & ""

For i = 1 To vs.Rows - 1

If vs.TextMatrix(i, 0) <> "" Then

RS.AddNew
RS.Fields("INVOICENO").Value = txtHeating.Text
RS.Fields("INVOICEDATE").Value = Dates.Value
RS.Fields("SUBLEDGER").Value = txtParty.Text
RS.Fields("GENLEDGER").Value = "Sundry Debtors"
RS.Fields("BOOKCODE").Value = vs.TextMatrix(i, 0)
RS.Fields("TBook").Value = vs.TextMatrix(i, 1)
RS.Fields("LoosBook").Value = vs.TextMatrix(i, 2)
RS.Fields("TotalBook").Value = Val(vs.TextMatrix(i, 3))
RS.Fields("NetBook").Value = vs.TextMatrix(i, 4)
RS.Fields("remarks").Value = vs.TextMatrix(i, 5)
RS.Fields("Book_Code").Value = vs.TextMatrix(i, 6)
RS.Update

End If

Next


End If

End If


Exit Sub
aa1:
MsgBox Err.Description


End Sub
Sub searchData()

If RS.State = 1 Then RS.Close
RS.Open "select * from invoicea where INVOICENO=" & txtHeating.Text & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
txtParty.Text = RS.Fields("SUBLEDGER").Value
txtRemarks.Text = RS.Fields("Remarks").Value & ""
Text1.Text = RS.Fields("add1").Value & ""
Text2.Text = RS.Fields("add2").Value & ""
If Not IsNull(RS.Fields("godown").Value) Then
cbogodown = RS.Fields("godown").Value & ""
Else
cbogodown.ListIndex = -1
End If

End If



If RS.State = 1 Then RS.Close
RS.Open "select * from invoiceb where INVOICENO=" & txtHeating.Text & "", con, adOpenDynamic, adLockOptimistic
For i = 1 To RS.RecordCount
If RS.EOF = False Then
vs.TextMatrix(i, 0) = RS.Fields("BOOKCODE").Value
vs.TextMatrix(i, 1) = RS.Fields("TBook").Value
vs.TextMatrix(i, 2) = RS.Fields("LoosBook").Value
vs.TextMatrix(i, 3) = RS.Fields("TotalBook").Value
vs.TextMatrix(i, 4) = RS.Fields("NetBook").Value
vs.TextMatrix(i, 5) = RS.Fields("remarks").Value & ""
vs.TextMatrix(i, 6) = RS.Fields("Book_Code").Value & ""
RS.MoveNext
End If
Next

total

End Sub
Private Sub cmdUndo_5_Click()
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = False
   cmdPrint_7.Enabled = True
   cmdSave_2.Enabled = False
   cmdUndo_5.Enabled = False
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
End Sub



Private Sub Dates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtParty.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 
If KeyCode = 27 Then
     If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
        Unload Me
     End If
 End If
End Sub
Sub TotalFinal()
   If txtTotal3.Text = "" Then
      txtTotal3.Text = 0
   End If
   
   If txtTotal2.Text = "" Then
      txtTotal2.Text = 0
   End If
   
   
    txtRawAndCasting.Text = (CDbl(txtTotal2.Text) + CDbl(txtTotal3.Text))
    txtRawAndCasting.Text = Format(txtRawAndCasting.Text, "#,###.000")
End Sub
Private Sub Form_Load()
 
 SetWidth
 AddItemInGrid
 AddItemInGrid1
 AddItemInGrid2
 AddItemInGrid3
 SetWidth
 
 Dates.Value = Date
 txtHeating.Text = MaxSNo("invoicea", "INVOICENO")
 
 Dim s As String
 
 s = ""
 
 
 If RS.State = 1 Then RS.Close
 RS.Open "select * from remarks order by head", con
 While RS.EOF = False
 If s = "" Then
 s = RS(0)
 Else
 s = s & "|" & RS(0)
 End If
 RS.MoveNext
 Wend
 
 vs.ColComboList(5) = s
 
 If RS.State = 1 Then RS.Close
 RS.Open "select * from godownmaster order by id", con
 While RS.EOF = False
       cbogodown.AddItem RS(0)
       RS.MoveNext
 Wend
 
End Sub
Sub SetWidth()
vs.Cols = 7
vs.FormatString = "Books Name|^Gaddi|^Books in a Gaddi|^Loose Books|^Total Books|Remarks"
vs.ColWidth(0) = 3200
vs.ColWidth(1) = 1500
vs.ColWidth(2) = 1500
vs.ColWidth(3) = 1200
vs.ColWidth(4) = 1200
vs.ColWidth(5) = 2000
vs.ColWidth(6) = 2000
End Sub
Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then ToDate.SetFocus
End Sub
Private Sub HeatingDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then txtParty.SetFocus
End Sub

Private Sub ListHeatingNo_Click()
  Call cmdRef_Click
  searchData
  TotalFinal
  'Frame1.Visible = False
End Sub
Private Sub ToDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then Call cmdadd_Click
End Sub
Private Sub txtGrade_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub

Private Sub txtHeating_GotFocus()
If PopUpValue1 <> "" Then
txtHeating.Text = PopUpValue1
Dates.Value = PopUpValue2
vs.Clear
SetWidth
searchData
PopUpValue1 = ""
PopUpValue2 = ""
End If
End Sub

Private Sub txtHeating_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
popuplist2 "select INVOICENO,INVOICEDATE,SUBLEDGER from invoicea order by INVOICENO", con
End If

If KeyCode = 13 Then
searchData
End If

End Sub

Private Sub txtHeating_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
        
   Dates.SetFocus

        
  End If
  

End Sub
Private Sub txtParty_GotFocus()
If PopUpValue1 <> "" Then
txtParty.Text = PopUpValue1
Text1.Text = PopUpValue2
Text2.Text = PopUpValue3
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
End If
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
popuplist2 "select SUBLEDGER as [Binder Name],Address1 as Address,Address2 as City from SLEDGER order by SUBLEDGER", con
End If
If KeyCode = 13 Then
cbogodown.SetFocus
End If

End Sub

Private Sub txtParty_LostFocus()
   Record = ""
End Sub
Private Sub txtQty_GotFocus()
     txtQty.SelLength = 10
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub
Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub
Private Sub txtSize_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtGrade.SetFocus
End Sub

Private Sub vs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs.Col = 0 Then
        cellposi
        vs.TextMatrix(vs.RowSel, 0) = cboItem.Text
     End If
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

  
  
  If KeyCode = 115 Then
  If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
    vs.RemoveItem (vs.RowSel)
    total
  End If
  End If
  
  If KeyCode = 13 Then
     
     If vs.Col = 0 Then
        vs.Editable = flexEDNone
        VsFrame.Visible = True
        cboItem.SetFocus
     Else
        vs.Editable = flexEDKbdMouse
        cellposi
     End If

  End If
  
  
  
  
  
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        
          
 If vs.Col = 0 Then
 
      vs.Editable = flexEDNone
      VsFrame.Visible = True
      cboItem.SetFocus
          
  If vs.TextMatrix(vs.RowSel, 0) <> "" Then
       SendKeys "{right}"
  End If
       
       VsFrame.Visible = False
       vs.Editable = flexEDKbdMouse
       vs.SetFocus
 '  End If
    
 End If
 If vs.Col = 1 Then
 If RS.State = 1 Then RS.Close
 RS.Open "select BKRATE from [BM] where BKDESC='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
 If RS.EOF = False Then
       vs.TextMatrix(vs.RowSel, 2) = RS.Fields(0).Value
       vs.TextMatrix(vs.RowSel, 4) = (IIf(vs.TextMatrix(vs.RowSel, 1) = "", 0, vs.TextMatrix(vs.RowSel, 1)) * IIf(vs.TextMatrix(vs.RowSel, 2) = "", 0, vs.TextMatrix(vs.RowSel, 2)) + IIf(vs.TextMatrix(vs.RowSel, 3) = "", 0, vs.TextMatrix(vs.RowSel, 3)))
End If

 SendKeys "{right}"
 SendKeys "{right}"

 End If
    
If vs.Col = 3 Then
SendKeys "{right}"
SendKeys "{right}"
vs.TextMatrix(vs.RowSel, 4) = (IIf(vs.TextMatrix(vs.RowSel, 1) = "", 0, vs.TextMatrix(vs.RowSel, 1)) * IIf(vs.TextMatrix(vs.RowSel, 2) = "", 0, vs.TextMatrix(vs.RowSel, 2)) + IIf(vs.TextMatrix(vs.RowSel, 3) = "", 0, vs.TextMatrix(vs.RowSel, 3)))

End If


If vs.Col = 5 Then
vs.TextMatrix(vs.RowSel, 4) = (IIf(vs.TextMatrix(vs.RowSel, 1) = "", 0, vs.TextMatrix(vs.RowSel, 1)) * IIf(vs.TextMatrix(vs.RowSel, 2) = "", 0, vs.TextMatrix(vs.RowSel, 2)) + IIf(vs.TextMatrix(vs.RowSel, 3) = "", 0, vs.TextMatrix(vs.RowSel, 3)))
cboItem.Text = ""
SendKeys "{home}"
SendKeys "{down}"
total
End If
    
       
total

End If

End Sub

Private Sub vs_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)

'If KeyCode = 13 Then
'
'If vs.Col = 1 Then
'If Not (KeyCode >= 48 And KeyCode <= 57) Then
'MsgBox "Enter Only Numeric Value !!", vbInformation
'vs.TextMatrix(vs.RowSel, 1) = ""
'vs.SetFocus
'Exit Sub
'End If
'End If
'
'
'If vs.Col = 2 Then
'If Not (KeyCode >= 48 And KeyCode <= 57) Then
'MsgBox "Enter Only Numeric Value !!", vbInformation
'vs.TextMatrix(vs.RowSel, 2) = ""
'vs.SetFocus
'Exit Sub
'End If
'End If
'
'
'If vs.Col = 3 Then
'If Not (KeyCode >= 48 And KeyCode <= 57) Then
'MsgBox "Enter Only Numeric Value !!", vbInformation
'vs.TextMatrix(vs.RowSel, 3) = ""
'vs.SetFocus
'Exit Sub
'End If
'End If
'
'
'If vs.Col = 4 Then
'If Not (KeyCode >= 48 And KeyCode <= 57) Then
'MsgBox "Enter Only Numeric Value !!", vbInformation
'vs.TextMatrix(vs.RowSel, 4) = ""
'vs.SetFocus
'Exit Sub
'End If
'End If
'
'
'
'
'End If

End Sub

Private Sub vs_LeaveCell()
  total
End Sub

Private Sub vs1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs1.Col = 0 Then
        cellposiVs
        vs1.TextMatrix(vs1.RowSel, 0) = cboitemvs1.Text
     End If

End Sub
Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 46 Then
    vs1.RemoveItem (vs1.RowSel)
    'Total1
    TotalFinal
  End If
  
  If KeyCode = 13 Then
     If vs1.Col = 0 Then
        vs1.Editable = flexEDNone
        Vs1Frame.Visible = True
        cboitemvs1.Visible = True
        cboitemvs1.SetFocus
     Else
        vs1.Editable = flexEDKbdMouse
        cellposiVs
     End If
  End If
End Sub

Private Sub vs1_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
        
 If vs1.Col = 0 Then
    vs1.Editable = flexEDNone
    Vs1Frame.Visible = True
    cboitemvs1.SetFocus
          
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.Open "select * from ItemMaster where ItemName='" & vs1.TextMatrix(vs1.RowSel, 0) & "'", con
    If RS.EOF = False Then
       vs1.TextMatrix(vs1.RowSel, 1) = RS.Fields("Unit").Value
       SendKeys "{right}"
       SendKeys "{right}"
       Vs1Frame.Visible = False
       vs1.Editable = flexEDKbdMouse
       vs1.SetFocus
    Else
       vs1.TextMatrix(vs1.RowSel, 1) = "Kg"
       SendKeys "{right}"
       SendKeys "{right}"
       Vs1Frame.Visible = False
       vs1.Editable = flexEDKbdMouse
       vs1.SetFocus
    
    End If
    
 End If
    
 If vs1.Col = 2 Then
           
    SendKeys "{home}"
    SendKeys "{down}"
    
    AddItemInGrid1
 End If
    
    

 'Total1
 TotalFinal

End If


End Sub
Sub AddSemifinish()
   Dim j As Integer
   
   j = 1
    
   vs3.Clear
   For i = 1 To vs1.Rows - 1
    
   If vs1.TextMatrix(i, 0) <> "" Then
      If RS.State = 1 Then RS.Close
      RS.Open "select * from ItemMaster where ItemName='" & vs1.TextMatrix(i, 0) & "'", con
      If RS.Fields("itemgp").Value = "Semi Finish (R/D)" Or RS.Fields("itemgp").Value = "Semi Finish (Store)" Then
         vs3.TextMatrix(j, 0) = vs1.TextMatrix(i, 0)
         vs3.TextMatrix(j, 1) = vs1.TextMatrix(i, 1)
         vs3.TextMatrix(j, 2) = vs1.TextMatrix(i, 2)
         j = j + 1
      End If
   End If
        
   Next
    
    
End Sub
Private Sub vs1_LeaveCell()
   'Total1
End Sub

Private Sub vs2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs2.Col = 0 Then
        'cellposiVs3
        vs2.TextMatrix(vs2.RowSel, 0) = cboItemVs3.Text
     End If

End Sub

Private Sub vs2_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 46 Then
    vs2.RemoveItem (vs2.RowSel)
    'Total2
    TotalFinal
 End If

  
  If KeyCode = 13 Then
     
     If vs2.Col = 0 Then
        vs2.Editable = flexEDNone
        Vs3Frame.Visible = True
        cboItemVs3.Visible = True
        cboItemVs3.SetFocus
     Else
        vs2.Editable = flexEDKbdMouse
        'cellposiVs3
     End If

  End If

End Sub

Private Sub vs2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        
          
 If vs2.Col = 0 Then
 
      vs2.Editable = flexEDNone
      Vs3Frame.Visible = True
      
      
          
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.Open "select * from ItemMaster where ItemName='" & vs2.TextMatrix(vs2.RowSel, 0) & "'", con
    If RS.EOF = False Then
       vs2.TextMatrix(vs2.RowSel, 1) = RS.Fields("Unit").Value
       SendKeys "{right}"
       SendKeys "{right}"
       Vs3Frame.Visible = False
       vs2.Editable = flexEDKbdMouse
       vs2.SetFocus
    Else
       vs2.TextMatrix(vs2.RowSel, 1) = "Kg"
       SendKeys "{right}"
       SendKeys "{right}"
       Vs3Frame.Visible = False
       vs2.Editable = flexEDKbdMouse
       vs2.SetFocus
    
    End If
    
 End If
 
    
    If vs2.Col = 2 Then
           
           SendKeys "{home}"
           SendKeys "{down}"
           Vs3Frame.Top = Vs3Frame.Top + 170
    End If
    
       
   'Total2

End If

End Sub
Private Sub vs2_LeaveCell()
   'Total2
End Sub
Private Sub vs3_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs3.Col = 0 Then
        'cellposiVs2
        'vs3.TextMatrix(vs3.RowSel, 0) = cboitemvscboItemVs2.Text
     End If
 
End Sub

Private Sub vs3_KeyDown(KeyCode As Integer, Shift As Integer)
    
  If KeyCode = 46 Then
    vs3.RemoveItem (vs3.RowSel)
    'Total4
  End If
  
  If KeyCode = 13 Then
     If vs3.Col = 0 Then
        
        vs3.Editable = flexEDNone
        FrameVs2.Visible = True
        cboItemVs2.Visible = True
        cboItemVs2.SetFocus
     Else
        
        vs3.Editable = flexEDKbdMouse
        
     End If
  End If

End Sub

Private Sub vs3_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
 
 If vs3.Col = 0 Then
    vs3.Editable = flexEDNone
    FrameVs2.Visible = True
    cboItemVs2.SetFocus
    
    
 
          
    Set RS = New ADODB.Recordset
    If RS.State = 1 Then RS.Close
    RS.Open "select * from ItemMaster where ItemName='" & vs3.TextMatrix(vs3.RowSel, 0) & "'", con
    If RS.EOF = False Then
       vs3.TextMatrix(vs3.RowSel, 1) = RS.Fields("Unit").Value
       SendKeys "{right}"
       SendKeys "{right}"
       FrameVs2.Visible = False
       vs3.Editable = flexEDKbdMouse
       vs3.SetFocus
    Else
       vs3.TextMatrix(vs3.RowSel, 1) = "Kg"
       SendKeys "{right}"
       SendKeys "{right}"
       FrameVs2.Visible = False
       vs3.Editable = flexEDKbdMouse
       vs3.SetFocus
    
    End If
    
 End If
    
 If vs3.Col = 2 Then
    
   If RS.State = 1 Then RS.Close
   RS.Open "select  OpeningStock from ItemMaster where ItemName='" & vs3.TextMatrix(vs.RowSel, 0) & "'", con
   If RS.EOF = False Then
      If Val(RS.Fields(0).Value) < Val(vs3.TextMatrix(vs3.RowSel, 2)) Then
         MsgBox "Stock Less !!", vbInformation
         
      End If
   End If
    
    
    SendKeys "{home}"
    SendKeys "{down}"
    
    FrameVs2.Top = FrameVs2.Top + 170
    'AddItemInGrid2
 End If
    
    

 'Total4
 
End If
 
End Sub
