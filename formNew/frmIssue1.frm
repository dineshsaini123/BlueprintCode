VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmIssue1 
   Caption         =   "Binder Book  Receive"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14220
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   14220
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtId 
      Height          =   285
      Left            =   1665
      TabIndex        =   0
      Top             =   270
      Width           =   1410
   End
   Begin VB.TextBox txtHeating 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3150
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   270
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   180
      TabIndex        =   30
      Top             =   630
      Width           =   12000
      Begin VB.TextBox txtParty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7305
         TabIndex        =   4
         Top             =   135
         Width           =   3690
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1725
         TabIndex        =   7
         Top             =   1185
         Width           =   4065
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7305
         TabIndex        =   32
         Top             =   465
         Width           =   3690
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7305
         TabIndex        =   31
         Top             =   765
         Width           =   3690
      End
      Begin VB.ComboBox cbogodown 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   810
         Width           =   960
      End
      Begin VB.TextBox txtTopay 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7320
         TabIndex        =   5
         Top             =   1080
         Width           =   1275
      End
      Begin VB.ComboBox cboFirm 
         Height          =   315
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1395
         Width           =   3705
      End
      Begin MSComCtl2.DTPicker Dates 
         Height          =   315
         Left            =   1725
         TabIndex        =   1
         Top             =   465
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   59899905
         CurrentDate     =   39500
      End
      Begin VB.Label Label1 
         Caption         =   "Binder Name "
         Height          =   300
         Index           =   2
         Left            =   6105
         TabIndex        =   38
         Top             =   165
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Date "
         Height          =   270
         Index           =   1
         Left            =   180
         TabIndex        =   37
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Remarks "
         Height          =   300
         Index           =   4
         Left            =   210
         TabIndex        =   36
         Top             =   1140
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Godown Name"
         Height          =   270
         Index           =   7
         Left            =   210
         TabIndex        =   35
         Top             =   855
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "To Pay"
         Height          =   300
         Index           =   8
         Left            =   6195
         TabIndex        =   34
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Firm Name :"
         Height          =   300
         Index           =   9
         Left            =   6195
         TabIndex        =   33
         Top             =   1395
         Width           =   1140
      End
   End
   Begin VB.TextBox txtLoose 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6360
      TabIndex        =   29
      Top             =   6630
      Width           =   1215
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3420
      TabIndex        =   28
      Top             =   6570
      Width           =   1455
   End
   Begin Crystal.CrystalReport CR 
      Left            =   7680
      Top             =   7380
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
      TabIndex        =   19
      Top             =   7050
      Width           =   7290
      Begin VB.CommandButton cmdExit_12 
         Caption         =   "E&xit"
         Height          =   480
         Left            =   6120
         TabIndex        =   26
         Top             =   255
         Width           =   1005
      End
      Begin VB.CommandButton cmdPrint_7 
         Caption         =   "&Print"
         Height          =   480
         Left            =   5100
         TabIndex        =   25
         Top             =   255
         Width           =   1005
      End
      Begin VB.CommandButton cmdUndo_5 
         Caption         =   "&Undo"
         Height          =   480
         Left            =   4095
         TabIndex        =   24
         Top             =   255
         Width           =   1005
      End
      Begin VB.CommandButton cmdEdit_4 
         Caption         =   "&Edit"
         Height          =   480
         Left            =   3090
         TabIndex        =   23
         Top             =   255
         Width           =   1005
      End
      Begin VB.CommandButton cmdDelete_3 
         Caption         =   "&Delete"
         Height          =   480
         Left            =   2085
         TabIndex        =   22
         Top             =   255
         Width           =   1005
      End
      Begin VB.CommandButton cmdSave_2 
         Caption         =   "&Save"
         Height          =   480
         Left            =   1080
         TabIndex        =   21
         Top             =   255
         Width           =   1005
      End
      Begin VB.CommandButton cmdAdd_1 
         Caption         =   "&Add"
         Height          =   480
         Left            =   75
         TabIndex        =   20
         Top             =   255
         Width           =   1005
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   -135
      Top             =   9000
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
      TabIndex        =   12
      Top             =   8805
      Visible         =   0   'False
      Width           =   465
      Begin VB.TextBox txtRawAndCasting 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3690
         TabIndex        =   13
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   570
         Width           =   1635
      End
   End
   Begin VB.TextBox txtTotal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7650
      TabIndex        =   9
      Text            =   "0"
      Top             =   6630
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   4095
      Left            =   180
      TabIndex        =   8
      Top             =   2460
      Width           =   12045
      _cx             =   21246
      _cy             =   7223
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
      Rows            =   150
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
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
         Left            =   4050
         TabIndex        =   17
         Top             =   2475
         Visible         =   0   'False
         Width           =   4155
         Begin MSDataListLib.DataCombo cboItem 
            Height          =   2310
            Left            =   0
            TabIndex        =   18
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
      Caption         =   "Challan No :"
      Height          =   270
      Index           =   0
      Left            =   180
      TabIndex        =   39
      Top             =   285
      Width           =   1530
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
      TabIndex        =   27
      Top             =   0
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
      TabIndex        =   11
      Top             =   45
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Total :"
      Height          =   195
      Index           =   6
      Left            =   2340
      TabIndex        =   10
      Top             =   6630
      Width           =   660
   End
End
Attribute VB_Name = "frmIssue1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itemgp As String
Dim rates As Double
Dim I As Integer
Dim Status As String
Dim Item_Name As String
Dim unit As String
Dim qty As Integer
Dim iitem1 As String
Dim edit As Boolean
Dim StockFlag As String

Private Sub cmdMain_Click()
Unload Me
End Sub
Sub cellposi()
 'VsFrame.Width = 3165
 'VsFrame.Top = vs.Top + ((vs.CellTop)) - 1400
 'VsFrame.Left = (vs.Left) - 200
End Sub
Sub Total()
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
'    rs_4.Open "select * from Books order by BookName", con, adOpenDynamic, adLockOptimistic
'
'    Set cboitemvs1.RowSource = rs_4
'    cboitemvs1.ListField = "BookName"
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
    rs_1.Open "select * from bm order by BKDESC", con_Binder, adOpenDynamic, adLockOptimistic
    
    Set cboItem.RowSource = rs_1
    cboItem.ListField = "BKCODE"
    cboItem.BoundColumn = "BKDESC"
    cboItem.ReFill
    
End Sub
Private Sub cboFinishItem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtQty.SetFocus
End Sub




Private Sub cboFirm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Len(cboFirm.Text) > 0 Then
txtRemarks.SetFocus
End If
End If
End Sub

Private Sub cbogodown_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then txtTopay.SetFocus
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
      
         If rs.State = 1 Then rs.close
         rs.Open "select * from ItemMaster where ItemName='" & iitem1 & "'", CON, adOpenDynamic, adLockOptimistic
         If rs.EOF = True Then
            rs.AddNew
            rs.Fields("ItemGp").value = frmAddMaster.cboGp.Text
            rs.Fields("ItemName").value = iitem1
            rs.Fields("Unit").value = "Kg"
            rs.Update
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
        
        If rs.State = 1 Then rs.close
        rs.Open "select * from ItemMaster where ItemName='" & cboitemvs1.Text & "'", CON
        If rs.EOF = True Then
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
        Set rs = New ADODB.Recordset
        If rs.State = 1 Then rs.close
        rs.Open "select * from ItemMaster where ItemName='" & cboItemVs2.Text & "'", CON
        If rs.EOF = True Then
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
        Set rs = New ADODB.Recordset
        If rs.State = 1 Then rs.close
        rs.Open "select * from ItemMaster where ItemName='" & cboItemVs3.Text & "'", CON
        If rs.EOF = True Then
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
 If rs.State = 1 Then rs.close
 rs.Open "select HeatingNo from IssueMaster where HeatingDate >=datevalue('" & fromdate.value & "') and HeatingDate <=datevalue('" & todate.value & "') order by HeatingNo", CON
 ListHeatingNo.Clear
 If rs.EOF = False Then
    While rs.EOF = False
       ListHeatingNo.AddItem rs(0)
       rs.MoveNext
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
 fromdate.SetFocus
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
Dim search As New ADODB.Recordset
    
If search.State = 1 Then search.close
search.Open "select ItemName,qty from Invoice where HeatNo='" & txtHeating.Text & "'", CON
If search.EOF = False Then
While search.EOF = False

    If rss.State = 1 Then rss.close
    rss.Open "select * from IssueRawMetrial where HeatingNo=" & txtHeating.Text & " and ItemName='" & search.Fields(0).value & "'", CON, adOpenDynamic, adLockOptimistic
    If rss.EOF = False Then
       rss.Fields("Issue").value = (CDbl(rss.Fields("Issue").value) + CDbl(search.Fields("qty").value))
       rss.Update
    End If
    
    search.MoveNext
    
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
      
      setwidth
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
    
    
    
    
    If rs.State = 1 Then rs.close
    rs.Open "select * from IssueMaster where HeatingNo=" & txtHeating.Text & "", CON
    If rs.EOF = False Then
       MsgBox "Heating No. Already Exist !!", vbInformation
       Exit Sub
    End If
    
    If txtHeating.Text = "" Then
       MsgBox "Please Enter Heating No !!", vbCritical
       txtHeating.SetFocus
       Exit Sub
    End If
    
    
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.close
    rs.Open "select * from IssueMaster where HeatingNo=" & txtHeating.Text & "", CON
    If rs.EOF = True Then
       If MsgBox("Want To Save ?", vbQuestion + vbYesNo) = vbYes Then
        '  SaveData
       End If
    Else
          MsgBox "Dublicate Heating No !!", vbCritical
    End If
End Sub
Sub ItemGpSearch(str As String)
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select ItemGp,Rate from ItemMaster where ItemName='" & str & "'", CON
    If rs1.EOF = False Then
       itemgp = rs1.Fields(0).value
       rates = rs1.Fields(1).value
    End If
    
End Sub
Sub UpdateStock()
    Dim rr As New ADODB.Recordset
    Dim rs_u As New ADODB.Recordset
    Dim openning As Double
    
 
    
    
 '================ Issue For Casting
 
 
 If StockFlag = "1" Then
    
    If rs_u.State = 1 Then rs_u.close
    rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
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
    
        If rs_u.State = 1 Then rs_u.close
        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
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
    
        If rs_u.State = 1 Then rs_u.close
        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
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
    
        If rs_u.State = 1 Then rs_u.close
        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
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
   
    edit = False
    txtHeating.Text = ""
    Dates.value = Date
    txtParty.Text = ""
    txtRemarks.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    txtTotal.Text = ""
    txtLoose.Text = ""
   txtTopay = ""
   cboFirm.ListIndex = -1
   
   'RefData Me
   vs.Clear
   setwidth
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = True
   'cmdPrint_7.Enabled = False
   cmdSave_2.Enabled = True
   cmdUndo_5.Enabled = True
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
   'txtHeating.SetFocus
   txtHeating.Text = MaxSNo("BinderBkReceive", "INVOICENO")
   txtId.Text = MaxSNo("BinderBkReceive", "INVOICENO")
   
   cbogodown.ListIndex = -1
   
   formButtonValidation cmdDelete_3, cmdEdit_4
   
   Frame1.Enabled = True
   vs.Enabled = True
   'txtHeating.SetFocus
   txtId.SetFocus
   
End Sub

Private Sub cmdDelete_3_Click()
If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
CON.Execute "delete from BookReceiveDet where INVOICENO=" & txtHeating.Text & ""
CON.Execute "delete from BinderBkReceive where INVOICENO=" & txtHeating.Text & ""
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
   Frame1.Enabled = True
   edit = True
   txtId.SetFocus
   vs.Enabled = True
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdPrint_7_Click()

CR.Reset
CR.ReportFileName = App.Path & "/Reports/CHALLAN.rpt"
CR.ReplaceSelectionFormula "{BinderBkReceive.invoiceno}=" & txtHeating.Text & ""
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

If edit = False Then
    txtHeating.Text = MaxSNo("BinderBkReceive", "INVOICENO")
    txtId.Text = txtHeating.Text
End If


If MsgBox("Want to Save ?", vbYesNo + vbQuestion) = vbYes Then

  

If rs.State = 1 Then rs.close
rs.Open "select * from BinderBkReceive where INVOICENO=" & txtId.Text & "", CON, adOpenDynamic, adLockOptimistic
If rs.EOF = True Then
rs.AddNew
rs.Fields("INVOICENO").value = txtHeating.Text
rs.Fields("INVOICEDATE").value = Dates.value
rs.Fields("SUBLEDGER").value = txtParty.Text
rs.Fields("GENLEDGER").value = "Sundry Debtors"
rs.Fields("Remarks").value = txtRemarks.Text
rs.Fields("add1").value = Text1.Text
rs.Fields("add2").value = Text2.Text
rs.Fields("godown").value = cbogodown
rs.Fields("topay").value = Val(txtTopay)
rs.Fields("firmname").value = cboFirm.Text

rs.Update
Else

rs.Fields("godown").value = cbogodown
rs.Fields("INVOICEDATE").value = Dates.value
rs.Fields("SUBLEDGER").value = txtParty.Text
rs.Fields("GENLEDGER").value = "Sundry Debtors"
rs.Fields("Remarks").value = txtRemarks.Text
rs.Fields("add1").value = Text1.Text
rs.Fields("add2").value = Text2.Text
rs.Fields("NetBook").value = Val(txtTotal1.Text)
rs.Fields("topay").value = Val(txtTopay)
rs.Fields("firmname").value = cboFirm.Text
rs.Update
cmdSave_2.Enabled = False
cmdPrint_7.SetFocus
End If



If rs.State = 1 Then rs.close
rs.Open "select * from BookReceiveDet where INVOICENO=" & txtHeating.Text & "", CON, adOpenDynamic, adLockOptimistic
If rs.EOF = True Then


For I = 1 To vs.Rows - 1

If vs.TextMatrix(I, 0) <> "" Then

rs.AddNew
rs.Fields("INVOICENO").value = txtHeating.Text
rs.Fields("INVOICEDATE").value = Dates.value
rs.Fields("SUBLEDGER").value = txtParty.Text
rs.Fields("GENLEDGER").value = "Sundry Debtors"
rs.Fields("BOOKCODE").value = vs.TextMatrix(I, 6)
rs.Fields("TBook").value = IIf(vs.TextMatrix(I, 1) = "", 0, vs.TextMatrix(I, 1))
rs.Fields("LoosBook").value = vs.TextMatrix(I, 2)
rs.Fields("TotalBook").value = Val(vs.TextMatrix(I, 3))
rs.Fields("NetBook").value = vs.TextMatrix(I, 4)
rs.Fields("Remarks").value = vs.TextMatrix(I, 5)
rs.Fields("Book_Code").value = vs.TextMatrix(I, 0)
rs.Update
cmdSave_2.Enabled = False
cmdPrint_7.SetFocus
End If

Next

Else
CON.Execute "delete from BookReceiveDet where INVOICENO=" & txtHeating.Text & ""

For I = 1 To vs.Rows - 1

If vs.TextMatrix(I, 0) <> "" Then

rs.AddNew
rs.Fields("INVOICENO").value = txtHeating.Text
rs.Fields("INVOICEDATE").value = Dates.value
rs.Fields("SUBLEDGER").value = txtParty.Text
rs.Fields("GENLEDGER").value = "Sundry Debtors"
rs.Fields("BOOKCODE").value = vs.TextMatrix(I, 6)
rs.Fields("TBook").value = vs.TextMatrix(I, 1)
rs.Fields("LoosBook").value = vs.TextMatrix(I, 2)
rs.Fields("TotalBook").value = Val(vs.TextMatrix(I, 3))
rs.Fields("NetBook").value = vs.TextMatrix(I, 4)
rs.Fields("remarks").value = vs.TextMatrix(I, 5)
rs.Fields("Book_Code").value = vs.TextMatrix(I, 0)
rs.Update
cmdSave_2.Enabled = False
End If

Next


End If
'Call cmdAdd_1_Click

End If


Exit Sub
aa1:
MsgBox Err.Description


End Sub
Sub searchData()

If rs.State = 1 Then rs.close
rs.Open "select * from BinderBkReceive where INVOICENO=" & txtHeating.Text & "", CON, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then

cmdSave_2.Enabled = False

txtParty.Text = rs.Fields("SUBLEDGER").value
txtRemarks.Text = rs.Fields("Remarks").value & ""
Text1.Text = rs.Fields("add1").value & ""
Text2.Text = rs.Fields("add2").value & ""
txtTopay = rs.Fields("topay").value

cboFirm.Text = rs.Fields("firmname").value

If Not IsNull(rs.Fields("godown").value) Then
cbogodown = rs.Fields("godown").value & ""
Else
cbogodown.ListIndex = -1
End If

End If



If rs.State = 1 Then rs.close
rs.Open "select * from BookReceiveDet where INVOICENO=" & txtHeating.Text & "", CON, adOpenDynamic, adLockOptimistic
For I = 1 To rs.RecordCount
If rs.EOF = False Then
vs.TextMatrix(I, 6) = rs.Fields("BOOKCODE").value
vs.TextMatrix(I, 1) = rs.Fields("TBook").value
vs.TextMatrix(I, 2) = rs.Fields("LoosBook").value
vs.TextMatrix(I, 3) = rs.Fields("TotalBook").value
vs.TextMatrix(I, 4) = rs.Fields("NetBook").value
vs.TextMatrix(I, 5) = rs.Fields("remarks").value & ""
vs.TextMatrix(I, 0) = rs.Fields("Book_Code").value & ""
rs.MoveNext
End If
Next

formButtonValidation cmdDelete_3, cmdEdit_4

Total

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

Private Sub Form_Activate()
'txtHeating.SetFocus
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
 
 setwidth
 AddItemInGrid
' AddItemInGrid1
' AddItemInGrid2
' AddItemInGrid3
 setwidth
 
 
 txtHeating.Text = MaxSNo("BinderBkReceive", "INVOICENO")
 txtId = txtHeating.Text
 'Dates.value = Date
 
 Dim s As String
 
 s = ""
 
 
 If rs.State = 1 Then rs.close
 rs.Open "select * from remarks order by head", CON
 While rs.EOF = False
 If s = "" Then
 s = rs(0)
 Else
 s = s & "|" & rs(0)
 End If
 rs.MoveNext
 Wend
 
 vs.ColComboList(5) = s
 
 If rs.State = 1 Then rs.close
 rs.Open "select customer_name from CustomerMaster order by customer_id", CON
 While rs.EOF = False
       cbogodown.AddItem rs(0)
       rs.MoveNext
 Wend
 
 If rs.State = 1 Then rs.close
 rs.Open "select firmname from firmname order by firmname", CON
 While rs.EOF = False
       cboFirm.AddItem rs(0)
       rs.MoveNext
 Wend
 
 
  'Frame1.Enabled = False
  ' vs.Enabled = False
 
 Dates.value = Date
 
 formButtonValidation cmdDelete_3, cmdEdit_4
 
End Sub
Sub setwidth()
vs.Cols = 7
vs.FormatString = "BookCode|^Gaddi|^Books in a Gaddi|^Loose Books|^Total Books|Remarks|BookName"
vs.ColWidth(0) = 1500
vs.ColWidth(1) = 1500
vs.ColWidth(2) = 1500
vs.ColWidth(3) = 1200
vs.ColWidth(4) = 1200
vs.ColWidth(5) = 2000
vs.ColWidth(6) = 2000
End Sub
Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then todate.SetFocus
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
vs.Clear
setwidth
searchData
Dates.value = PopUpValue2

PopUpValue1 = ""
PopUpValue2 = ""


formButtonValidation cmdDelete_3, cmdEdit_4

Frame1.Enabled = False
vs.Enabled = False

End If
End Sub

Private Sub txtHeating_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
popuplist2 "select INVOICENO as [ChallanNo],INVOICEDATE as [Date],SUBLEDGER as Binder from BinderBkReceive order by INVOICENO", CON
End If

If KeyCode = 13 Then
   Dates.SetFocus
End If

End Sub

Private Sub txtHeating_KeyPress(KeyAscii As Integer)
   
''   If KeyAscii = 13 Then
''
''  'If Dates.Enabled = True Then
''  ' Dates.SetFocus
''  'End If
''
''  End If
  

End Sub

Private Sub txtId_GotFocus()
If PopUpValue1 <> "" Then
txtHeating.Text = PopUpValue1

txtId.Text = PopUpValue1

vs.Clear
setwidth
searchData
Dates.value = PopUpValue2

PopUpValue1 = ""
PopUpValue2 = ""


formButtonValidation cmdDelete_3, cmdEdit_4



Frame1.Enabled = False
vs.Enabled = False

End If

End Sub

Private Sub txtId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
popuplist2 "select INVOICENO as [ChallanNo],INVOICEDATE as [Date],SUBLEDGER as Binder from BinderBkReceive order by INVOICENO", CON
End If

If KeyCode = 13 Then
If Dates.Enabled = True Then
   Dates.SetFocus
End If
End If
End Sub

Private Sub txtId_KeyPress(KeyAscii As Integer)
'If KeyCode = 13 Then txtId.SetFocus

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
popuplist2 "select SUBLEDGER as [Binder Name],Address1 as Address,Address2 as City from SLEDGER order by SUBLEDGER", con_Binder
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
 
Dim w%
 
 If KeyAscii = 13 Then
  vs.SetFocus

w = vs.RowSel
For I = 1 To 100
SendKeys "{up}"
w = w - 1
If w = 1 Then
   Exit For
End If
  
Next

'======================================

For I = 1 To 100
If vs.TextMatrix(I, 0) <> "" Then
SendKeys "{down}"
Else
Exit For
End If
Next



SendKeys "{home}"

End If
End Sub
Private Sub txtSize_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtGrade.SetFocus
End Sub



Private Sub txtTopay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cboFirm.SetFocus
End Sub

Private Sub vs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs.Col = 0 Then
        cellposi
      If cboItem.Text <> "" Then
        vs.TextMatrix(vs.RowSel, 0) = cboItem.Text
      End If
     End If
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

  
  
  If KeyCode = 115 Then
  If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
    vs.RemoveItem (vs.RowSel)
    Total
  End If
  End If
  
  If KeyCode = 13 Then
     
     If vs.Col = 0 Then
        'vs.Editable = flexEDNone
        'VsFrame.Visible = True
        'cboItem.SetFocus
     If Val(txtTotal) > 0 Then
        Call cmdSave_2_Click
     End If
     
     Else
        'vs.Editable = flexEDKbdMouse
        cellposi
     End If

  End If
  
  
  
  
  
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)

On Error Resume Next

If KeyCode = 13 Then
        
          
 If vs.Col = 0 Then
 
  
 
 If rs.State = 1 Then rs.close
 rs.Open "select BKRATE,BKDESC from [bm] where BKCODE='" & vs.TextMatrix(vs.RowSel, 0) & "'", con_Binder
 If rs.EOF = False Then
       vs.TextMatrix(vs.RowSel, 0) = UCase(vs.TextMatrix(vs.RowSel, 0))
       vs.TextMatrix(vs.RowSel, 2) = rs.Fields(0).value
       vs.TextMatrix(vs.RowSel, 4) = (IIf(vs.TextMatrix(vs.RowSel, 1) = "", 0, vs.TextMatrix(vs.RowSel, 1)) * IIf(vs.TextMatrix(vs.RowSel, 2) = "", 0, vs.TextMatrix(vs.RowSel, 2)) + IIf(vs.TextMatrix(vs.RowSel, 3) = "", 0, vs.TextMatrix(vs.RowSel, 3)))
       vs.TextMatrix(vs.RowSel, 6) = rs.Fields(1).value
Else
   vs.SetFocus
   Exit Sub
End If

  If vs.TextMatrix(vs.RowSel, 0) <> "" Then
       SendKeys "{right}"
  End If
 
    
 End If

If vs.Col = 1 Then

If Val(vs.TextMatrix(vs.RowSel, 1)) = 0 Then Exit Sub

 SendKeys "{right}"
 SendKeys "{right}"

 End If
    
If vs.Col = 3 Then
SendKeys "{right}"
SendKeys "{right}"
vs.TextMatrix(vs.RowSel, 4) = (IIf(vs.TextMatrix(vs.RowSel, 1) = "", 0, Val(vs.TextMatrix(vs.RowSel, 1))) * IIf(vs.TextMatrix(vs.RowSel, 2) = "", 0, vs.TextMatrix(vs.RowSel, 2)) + IIf(vs.TextMatrix(vs.RowSel, 3) = "", 0, vs.TextMatrix(vs.RowSel, 3)))

End If


If vs.Col = 5 Then
If Len(vs.TextMatrix(vs.RowSel, 5)) = 0 Then Exit Sub

vs.TextMatrix(vs.RowSel, 4) = (IIf(vs.TextMatrix(vs.RowSel, 1) = "", 0, vs.TextMatrix(vs.RowSel, 1)) * IIf(vs.TextMatrix(vs.RowSel, 2) = "", 0, vs.TextMatrix(vs.RowSel, 2)) + IIf(vs.TextMatrix(vs.RowSel, 3) = "", 0, vs.TextMatrix(vs.RowSel, 3)))
cboItem.Text = ""
SendKeys "{home}"
SendKeys "{down}"
Total
End If
    
       
Total

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
  Total
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
          
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.close
    rs.Open "select * from ItemMaster where ItemName='" & vs1.TextMatrix(vs1.RowSel, 0) & "'", CON
    If rs.EOF = False Then
       vs1.TextMatrix(vs1.RowSel, 1) = rs.Fields("Unit").value
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
   For I = 1 To vs1.Rows - 1
    
   If vs1.TextMatrix(I, 0) <> "" Then
      If rs.State = 1 Then rs.close
      rs.Open "select * from ItemMaster where ItemName='" & vs1.TextMatrix(I, 0) & "'", CON
      If rs.Fields("itemgp").value = "Semi Finish (R/D)" Or rs.Fields("itemgp").value = "Semi Finish (Store)" Then
         vs3.TextMatrix(j, 0) = vs1.TextMatrix(I, 0)
         vs3.TextMatrix(j, 1) = vs1.TextMatrix(I, 1)
         vs3.TextMatrix(j, 2) = vs1.TextMatrix(I, 2)
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
      
      
          
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.close
    rs.Open "select * from ItemMaster where ItemName='" & vs2.TextMatrix(vs2.RowSel, 0) & "'", CON
    If rs.EOF = False Then
       vs2.TextMatrix(vs2.RowSel, 1) = rs.Fields("Unit").value
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
    
    
 
          
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.close
    rs.Open "select * from ItemMaster where ItemName='" & vs3.TextMatrix(vs3.RowSel, 0) & "'", CON
    If rs.EOF = False Then
       vs3.TextMatrix(vs3.RowSel, 1) = rs.Fields("Unit").value
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
    
   If rs.State = 1 Then rs.close
   rs.Open "select  OpeningStock from ItemMaster where ItemName='" & vs3.TextMatrix(vs.RowSel, 0) & "'", CON
   If rs.EOF = False Then
      If Val(rs.Fields(0).value) < Val(vs3.TextMatrix(vs3.RowSel, 2)) Then
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
