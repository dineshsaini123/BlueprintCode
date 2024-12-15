VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmOrdermgm 
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14850
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   14850
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   990
      Left            =   2760
      TabIndex        =   5
      Top             =   6360
      Width           =   4800
      Begin VB.CommandButton cmdRef 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   810
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   90
         Width           =   1140
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   810
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   90
         Width           =   1140
      End
      Begin VB.CommandButton cmdModify 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modify"
         Enabled         =   0   'False
         Height          =   810
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   90
         Width           =   90
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   810
         Left            =   2415
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   90
         Width           =   1140
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   810
         Left            =   1236
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   90
         Width           =   1140
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print "
         Height          =   810
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   90
         Width           =   1140
      End
   End
   Begin VB.Frame VsFrame 
      Height          =   3390
      Left            =   1305
      TabIndex        =   0
      Top             =   1575
      Visible         =   0   'False
      Width           =   7950
      Begin VB.CommandButton Command4 
         Caption         =   "Search"
         Height          =   375
         Left            =   8160
         TabIndex        =   3
         Top             =   180
         Width           =   1035
      End
      Begin VB.TextBox txtv 
         Height          =   315
         Left            =   8160
         TabIndex        =   2
         Top             =   60
         Width           =   1755
      End
      Begin VB.ComboBox cboitem1 
         Height          =   315
         Left            =   8100
         TabIndex        =   1
         Top             =   540
         Width           =   3675
      End
      Begin MSDataListLib.DataCombo cboItem 
         Height          =   3285
         Left            =   0
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   60
         Width           =   7890
         _ExtentX        =   13917
         _ExtentY        =   5794
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
   Begin Crystal.CrystalReport CR 
      Left            =   300
      Top             =   6975
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   5835
      Left            =   120
      TabIndex        =   12
      Top             =   420
      Width           =   9435
      _cx             =   16642
      _cy             =   10292
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmOrdermgm.frx":0000
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
   End
   Begin VB.TextBox txtTotal1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8175
      TabIndex        =   13
      Text            =   "0"
      Top             =   5940
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Total :"
      Height          =   195
      Index           =   6
      Left            =   5580
      TabIndex        =   16
      Top             =   6900
      Width           =   660
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080C0FF&
      Height          =   240
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   19575
   End
   Begin VB.Label Label2 
      Caption         =   "F4 For Delete Record"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6360
      Width           =   2535
   End
End
Attribute VB_Name = "frmOrdermgm"
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
Dim StockFlag As String
Private Sub cmdMain_Click()
Unload Me
End Sub
Sub cellposi()
 'VsFrame.Width = 3165
 VsFrame.Top = vs.Top + ((vs.CellTop)) + 270
 VsFrame.Left = (vs.Left) + 20
End Sub
Sub cellposiVs3()
 'Vs3Frame.Width = 3165
 'Vs3Frame.Top = vs2.Top + ((vs2.CellTop)) + 400
 'Vs3Frame.Left = (vs2.Left) - 80
End Sub

Sub cellposiVs()
 Vs1Frame.Width = 3165
 Vs1Frame.Top = vs1.Top + ((vs1.CellTop)) + 250
 Vs1Frame.Left = (vs1.Left) + 50
End Sub
Sub cellposiVs2()
 'FrameVs2.Width = 3165
 'FrameVs2.Top = vs3.Top + ((vs3.CellTop)) + 250
 'FrameVs2.Left = (vs3.Left) + 50
End Sub

Sub AddItemInGrid()
    'Adodc1.ConnectionString = "filedsn=Saru"
    'Adodc1.CommandType = adCmdText
    Dim rs_1 As New ADODB.Recordset
    
    'rs_1.Open "select * from ItemMaster where ItemGp='Raw Item' or ItemGp='Scrap' or (ItemGp='Chunnus (Pointing)' or ItemGp='Chunnus (Sonicut)') order by ItemName", con, adOpenDynamic, adLockOptimistic
    rs_1.Open "select SUBLEDGER from SLEDGER where gledger='SUNDRY DEBTORS' order by SUBLEDGER", CON, adOpenDynamic, adLockOptimistic
    
    Set cboItem.RowSource = rs_1
    cboItem.ListField = "SUBLEDGER"
    cboItem.BoundColumn = "SUBLEDGER"
    cboItem.ReFill
    
End Sub
Private Sub cboFinishItem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtQty.SetFocus
End Sub
Private Sub cboItem_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        'cellposi
     If vs.Col = 2 Then
        vs.TextMatrix(vs.RowSel, 2) = cboItem.Text
     
        
        VsFrame.Visible = False
        vs.SetFocus
     ElseIf vs.Col = 3 Then
     
     End If
        
     ElseIf KeyCode = 27 Then
       
          VsFrame.Visible = False
        
     End If
End Sub
Sub saveInMaster()
         On Error Resume Next
      
         If RS.State = 1 Then RS.close
         RS.Open "select * from ItemMaster where ItemName='" & iitem1 & "'", CON, adOpenDynamic, adLockOptimistic
         If RS.EOF = True Then
            RS.AddNew
            RS.Fields("ItemGp").value = frmAddMaster.cbogp.Text
            RS.Fields("ItemName").value = iitem1
            RS.Fields("Unit").value = "Kg"
            RS.update
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
        
        If RS.State = 1 Then RS.close
        RS.Open "select * from ItemMaster where ItemName='" & cboitemvs1.Text & "'", CON
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
        
        
        cellposiVs2
        vs3.TextMatrix(vs3.RowSel, 0) = cboItemVs2.Text
        Set RS = New ADODB.Recordset
        If RS.State = 1 Then RS.close
        RS.Open "select * from ItemMaster where ItemName='" & cboItemVs2.Text & "'", CON
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
        cellposiVs3
        
        vs2.TextMatrix(vs2.RowSel, 0) = cboItemVs3.Text
        Set RS = New ADODB.Recordset
        If RS.State = 1 Then RS.close
        RS.Open "select * from ItemMaster where ItemName='" & cboItemVs3.Text & "'", CON
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

Private Sub cmdAdd_Click()
 If RS.State = 1 Then RS.close
 RS.Open "select HeatingNo from IssueMaster where HeatingDate >=datevalue('" & fromDate.value & "') and HeatingDate <=datevalue('" & toDate.value & "')", CON
 ListHeatingNo.Clear
 If RS.EOF = False Then
    While RS.EOF = False
       ListHeatingNo.AddItem RS(0)
       RS.MoveNext
    Wend
 End If
End Sub
Private Sub cboMain_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then vs.SetFocus
End Sub
Private Sub cmdDelete_Click()
   
   If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
         CON.Execute "delete from OrderMnm"
         Call cmdRef_Click
   End If
   
End Sub
Sub DeleteStock()
    
Dim rr As New ADODB.Recordset
Dim rs_u As New ADODB.Recordset
Dim openning As Double
 
'================ Issue For Casting
 
 
 If StockFlag = "1" Then
    
    If rs_u.State = 1 Then rs_u.close
    rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
    If rs_u.EOF = False Then
        rs_u!qty = rs_u!qty + qty
        rs_u.update
    End If
 
 End If
 
 
 
 '================ Receive For Casting
 
 If StockFlag = "2" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.close
        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = False Then
            rs_u!qty = rs_u!qty - qty
            rs_u.update
        End If
    
    End If
    
 End If
 

 
 '================ Receive For Finish
 
 If StockFlag = "3" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.close
        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = False Then
             rs_u!qty = rs_u!qty - qty
            rs_u.update
        End If
    
    End If
    
 End If
 
  
 
 '================ Issue For Finish
 
 If StockFlag = "4" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.close
        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = False Then
             rs_u!qty = rs_u!qty + qty
            rs_u.update
        End If
    
    End If
    
 End If
 
 '====================================
    
    
    
    
   
    
   
   
    
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFatch_Click()
AddSemifinish
Total4

End Sub

Private Sub cmdFind_Click()
 Frame1.Visible = True
 fromDate.SetFocus
End Sub

Private Sub cmdModify_Click()
   If MsgBox("Want To Modify ?", vbQuestion + vbYesNo) = vbYes Then
      'DeleteRecord txtCode.Text, "VsNo", "IssueRawMetrial"
      saveData
      Call cmdRef_Click
   End If
End Sub


Private Sub cmdRef_Click()
      'txtParty.Text = ""
      
      
      txtTotal1.Text = 0
      
      
      
      vs.Clear
      setwidth
     
      cmdModify.Enabled = False
      cmdSave.Enabled = True
      
End Sub


Private Sub Command4_Click()
   Unload Me
End Sub

Private Sub CmdSave_Click()
saveData
End Sub

Sub saveData()
    
    CON.Execute "delete from OrderMnm where len(name)>0"
    
    If RS.State = 1 Then RS.close
    RS.Open "select * from OrderMnm", CON, adOpenDynamic, adLockOptimistic
    For I = vs.FixedRows To vs.Rows - 1
         If vs.TextMatrix(I, 0) <> "" Then
            RS.AddNew
            RS.Fields("id").value = vs.TextMatrix(I, 0)
            RS.Fields("dates").value = vs.TextMatrix(I, 1)
            RS.Fields("name").value = vs.TextMatrix(I, 2)
            RS.Fields("godwn").value = vs.TextMatrix(I, 3)
            
            RS.update
          End If
    Next
 
   
 MsgBox "Saved ....", vbInformation
     
    
End Sub
Sub SearchData()
  
    vs.Rows = 2
    
    If RS.State = 1 Then RS.close
    RS.Open "select * from OrderMnm order by Name", CON, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
           
           For I = 1 To RS.RecordCount
              If RS.EOF = False Then
                
                vs.TextMatrix(I, 0) = RS.Fields("id").value
                vs.TextMatrix(I, 1) = RS.Fields("dates").value
                vs.TextMatrix(I, 2) = RS.Fields("name").value
                vs.TextMatrix(I, 3) = RS.Fields("godwn").value
                
                 
             vs.Rows = vs.Rows + 1
             RS.MoveNext
              '------------------------------------------------------------
              End If
          Next
     End If
    '-----------------------------------------
    
    'vs.Rows = vs.Rows - 1
      
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

Private Sub Command1_Click()
Dim rs1 As New ADODB.Recordset
If RS.State = 1 Then RS.close
RS.Open "select ItemGp,ItemName from ItemMaster", CON
While RS.EOF = False
CON.Execute "update IssueRawMetrial set RawCode='" & RS(0) & "' where ItemName='" & RS(1) & "'"
RS.MoveNext
Wend
End Sub

Private Sub Command2_Click()
cr.Reset
cr.ReportFileName = App.Path & "\LOTSHEET.rpt"
cr.Connect = "filedsn=chitradsn;uid=sa;pwd=sidc;"
If txtCode.Text <> "" Then
cr.ReplaceSelectionFormula "{IssueRawMetrial.VSNo}='" & txtCode.Text & "'"
End If
cr.WindowShowPrintSetupBtn = True
cr.WindowShowPrintBtn = True
cr.WindowState = crptMaximized
cr.Action = 1

End Sub

Private Sub Command3_Click()


Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

If rs2.State = 1 Then rs21.close
rs2.Open "select ItemName from IssueRawMetrial", CON

Q1 = "SELECT IssueRawMetrial.ItemName, ItemMaster.ItemName1 " & _
"FROM IssueRawMetrial LEFT JOIN ItemMaster ON IssueRawMetrial.ItemName = ItemMaster.ItemName1 WHERE ItemMaster.ItemName1 IS NULL"


If RS.State = 1 Then RS.close
RS.Open Q1, CON
While RS.EOF = False

SSS = Trim(Mid(RS(0), 1, 30))

If rs1.State = 1 Then rs1.close
rs1.Open "select ItemName from ITEMMASTER where MID(ItemName1,1,25) = '" & Trim(Mid(RS(0), 1, 25)) & "'", CON
If rs1.EOF = False Then


   CON.Execute "update IssueRawMetrial set ItemName='" & RS(0) & "' where itemname='CERAMIC DISC. CAP./330 PF/Epoxy P=10mm/Y5P/K/6KV/(Y5P331K6KDL 10E (Lead Free)'"
End If

RS.MoveNext
Wend


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
End If
End Sub
Private Sub Form_Load()
setwidth
AddItemInGrid

S11 = ""

If RS.State = 1 Then RS.close
RS.Open "select Godwn from Godownmaster group by Godwn", CON
While RS.EOF = False

If S11 = "" Then
   S11 = RS(0)
Else
  S11 = S11 & "|" & RS(0)
End If

RS.MoveNext
Wend



vs.ColComboList(3) = S11



setwidth

SearchData
BackColorFrom Me

End Sub
Sub setwidth()


vs.Cols = 4
vs.FormatString = "OrderNo|Date|Party Name|Godown"
vs.ColWidth(0) = 1000
vs.ColWidth(1) = 1200
vs.ColWidth(2) = 5200
vs.ColWidth(3) = 1000

End Sub
Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then toDate.SetFocus
End Sub
Private Sub HeatingDate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtParty.SetFocus
End Sub
Private Sub ListHeatingNo_Click()
SearchData
TotalFinal
'Frame1.Visible = False
End Sub
Private Sub Todate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call cmdAdd_Click
End Sub
Private Sub txtGrade_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub
Private Sub txtHeating_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Set RS = New ADODB.Recordset
If txtHeating.Text = "" Then Exit Sub
If RS.State = 1 Then RS.close
RS.Open "select * from IssueMaster where HeatingNo=" & txtHeating.Text & "", CON
If RS.EOF = False Then
SearchData
TotalFinal
End If
HeatingDate.SetFocus
End If
End Sub
Private Sub txtcode_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then txtParty.SetFocus
'''If KeyCode = 113 Then
'''   popuplist12 "Select Consume_NonCon as Code from CollegeMaster where name='" & cboBearer.Text & "'", CON
'''End If
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
Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

For I = 1 To vs.Rows - 1
SendKeys "{up}"
SendKeys "{home}"
Next

vs.SetFocus
End If
End Sub

Private Sub vs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If vs.Col = 2 Then
cellposi
vs.TextMatrix(vs.RowSel, 2) = cboItem.Text
End If
End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 115 Then

If MsgBox("Want to Delete ?", vbQuestion + vbYesNo) = vbYes Then
   CON.Execute "delete from OrderMnm where id=" & vs.TextMatrix(vs.RowSel, 0) & ""
   vs.RemoveItem (vs.RowSel)
   total
End If

End If

If KeyCode = 13 Then
If vs.Col = 2 Then
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

If (vs.Col = 0) Then

If vs.Rows >= 3 Then
   vs.TextMatrix(vs.RowSel, 1) = vs.TextMatrix(vs.RowSel - 1, 1)
End If
SendKeys "{right}"
   
End If


If (vs.Col = 1) Then
SendKeys "{right}"
End If




If vs.Col = 2 Then
vs.Editable = flexEDNone
 
 SendKeys "{right}"
 vs.SetFocus

End If


If vs.Col = 3 Then
    
    SendKeys "{home}"
    SendKeys "{down}"
    
    vs.Rows = vs.Rows + 1
End If



total
End If
End Sub
Sub total()
    
    On Error Resume Next
    txtTotal1.Text = 0
    For I = 1 To vs.Rows - 1
        txtTotal1.Text = (Val(txtTotal1.Text) + Val(vs.TextMatrix(I, 3)))
    Next
    
    setwidth
End Sub
Sub Total4()
    
    On Error Resume Next
    txtTotal4.Text = 0
    For I = 1 To vs3.Rows - 1
        txtTotal4.Text = Format((CDbl(txtTotal4.Text) + CDbl(vs3.TextMatrix(I, 2))), "#,###.000")
    Next
    
    
End Sub

Sub total1()
    On Error Resume Next
    txtTotal2.Text = 0
    For I = 1 To vs1.Rows - 1
        txtTotal2.Text = Format((CDbl(txtTotal2.Text) + CDbl(vs1.TextMatrix(I, 2))), "#,###.000")
    Next
End Sub
Sub total2()
    
    On Error Resume Next
    
    txtTotal3.Text = 0
    For I = 1 To vs2.Rows - 1
        txtTotal3.Text = Format((CDbl(txtTotal3.Text) + CDbl(vs2.TextMatrix(I, 2))), "#,###.000")
    Next
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
    total1
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
    If RS.State = 1 Then RS.close
    RS.Open "select * from ItemMaster where ItemName='" & vs1.TextMatrix(vs1.RowSel, 0) & "'", CON
    If RS.EOF = False Then
       vs1.TextMatrix(vs1.RowSel, 1) = RS.Fields("Unit").value
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
    
    
 End If
    
    

 total1
 TotalFinal

End If


End Sub
Sub AddSemifinish()
   Dim J As Integer
   
   J = 1
    
   vs3.Clear
   For I = 1 To vs1.Rows - 1
    
   If vs1.TextMatrix(I, 0) <> "" Then
      If RS.State = 1 Then RS.close
      RS.Open "select * from ItemMaster where ItemName='" & vs1.TextMatrix(I, 0) & "'", CON
      If RS.Fields("itemgp").value = "Semi Finish (R/D)" Or RS.Fields("itemgp").value = "Semi Finish (Store)" Then
         vs3.TextMatrix(J, 0) = vs1.TextMatrix(I, 0)
         vs3.TextMatrix(J, 1) = vs1.TextMatrix(I, 1)
         vs3.TextMatrix(J, 2) = vs1.TextMatrix(I, 2)
         J = J + 1
      End If
   End If
        
   Next
    
    
End Sub
Private Sub vs1_LeaveCell()
   total1
End Sub

Private Sub vs2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs2.Col = 0 Then
        cellposiVs3
        vs2.TextMatrix(vs2.RowSel, 0) = cboItemVs3.Text
     End If

End Sub

Private Sub vs2_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 46 Then
    vs2.RemoveItem (vs2.RowSel)
    total2
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
    If RS.State = 1 Then RS.close
    RS.Open "select * from ItemMaster where ItemName='" & vs2.TextMatrix(vs2.RowSel, 0) & "'", CON
    If RS.EOF = False Then
       vs2.TextMatrix(vs2.RowSel, 1) = RS.Fields("Unit").value
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
    
       
   total2

End If

End Sub
Private Sub vs2_LeaveCell()
   total2
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
    Total4
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
    If RS.State = 1 Then RS.close
    RS.Open "select * from ItemMaster where ItemName='" & vs3.TextMatrix(vs3.RowSel, 0) & "'", CON
    If RS.EOF = False Then
       vs3.TextMatrix(vs3.RowSel, 1) = RS.Fields("Unit").value
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
    
   If RS.State = 1 Then RS.close
   RS.Open "select  OpeningStock from ItemMaster where ItemName='" & vs3.TextMatrix(vs.RowSel, 0) & "'", CON
   If RS.EOF = False Then
      If Val(RS.Fields(0).value) < Val(vs3.TextMatrix(vs3.RowSel, 2)) Then
         MsgBox "Stock Less !!", vbInformation
         
      End If
   End If
    
    
    SendKeys "{home}"
    SendKeys "{down}"
    
    FrameVs2.Top = FrameVs2.Top + 170
    'AddItemInGrid2
 End If
    
    

 Total4
 
End If
 
End Sub


