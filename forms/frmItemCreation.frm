VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmItemCreation 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Product Master"
   ClientHeight    =   8670
   ClientLeft      =   3645
   ClientTop       =   975
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtre 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1500
      TabIndex        =   17
      Top             =   1545
      Width           =   1125
   End
   Begin VB.TextBox txtop 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3675
      TabIndex        =   16
      Top             =   1260
      Width           =   1350
   End
   Begin VB.TextBox txtrarte 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   1500
      TabIndex        =   14
      Top             =   1245
      Width           =   1125
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   4455
      TabIndex        =   13
      Top             =   1575
      Visible         =   0   'False
      Width           =   585
   End
   Begin Crystal.CrystalReport cr 
      Left            =   300
      Top             =   9000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6435
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   870
      Width           =   2325
   End
   Begin VB.ComboBox cboUnit 
      Height          =   315
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   870
      Width           =   1095
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   270
      Width           =   1635
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   270
      Width           =   1575
   End
   Begin VB.CommandButton cmdRef 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   270
      Width           =   1455
   End
   Begin VB.CommandButton cmdMain 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   855
      Width           =   2460
   End
   Begin VB.ComboBox cboCategory 
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1500
      TabIndex        =   0
      Text            =   "cboCategory"
      Top             =   165
      Width           =   3525
   End
   Begin VB.TextBox txtQuality 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   555
      Width           =   4785
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   6180
      Left            =   180
      TabIndex        =   7
      Top             =   2250
      Width           =   11175
      _cx             =   19711
      _cy             =   10901
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   16711680
      BackColorFixed  =   16761024
      ForeColorFixed  =   255
      BackColorSel    =   16448755
      ForeColorSel    =   16744448
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   0
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
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
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   -1  'True
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
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Reorder"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   150
      TabIndex        =   19
      Top             =   1635
      Width           =   1125
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Opening"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2700
      TabIndex        =   18
      Top             =   1335
      Width           =   1125
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   180
      TabIndex        =   15
      Top             =   1335
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   180
      TabIndex        =   11
      Top             =   1005
      Width           =   615
   End
   Begin VB.Shape Shape1 
      Height          =   1425
      Left            =   6360
      Top             =   135
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name   "
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   585
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Group Name "
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   195
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E98A0A&
      BackStyle       =   0  'Transparent
      Caption         =   "Esc To Exit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3525
      TabIndex        =   8
      Top             =   7650
      Width           =   1215
   End
End
Attribute VB_Name = "frmItemCreation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bb As Boolean



Private Sub cboCategory_Click()
addProduct
End Sub

Private Sub cboCategory_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     txtQuality.SetFocus
  End If
End Sub
Private Sub cboProductId_Click()
cboCategory.ListIndex = cboProductid.ListIndex
End Sub
Private Sub cboGroup_Click()
cboCategory.ListIndex = cboGroup.ListIndex
End Sub

Private Sub cboUnit_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
      txtrarte.SetFocus
   End If
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   'If SetPerHead.Value <> 0 Then
   '   txtperhead.SetFocus
   'End If
End If
End Sub

Private Sub Check1_Click()

End Sub

Private Sub cmdDel_Click()
On Error Resume Next
If RS.State = 1 Then RS.close
         RS.Open "select * from ItemCreation where ItemCode=" & vs.TextMatrix(vs.RowSel, 0) & "", CON, adOpenDynamic, adLockOptimistic
         If RS.EOF = False Then
         If MsgBox("Do U Want To Delete ?", vbInformation + vbYesNo, "Message") = vbYes Then
            RS.delete
            addProduct
            addProductUnitCourse
            Call cmdRef_Click
            cboCategory_Click
         End If
         End If
         
End Sub
Private Sub cmdMain_Click()
Unload Me
End Sub
Sub AddCategory()
    Set RS = New ADODB.Recordset
    RS.Open "select distinct(Name) from ProductMaster", CON
    cboCategory.Clear
    If RS.EOF = False Then
       While RS.EOF = False
          cboCategory.AddItem RS.Fields(0).value
          RS.MoveNext
       Wend
    End If
    
End Sub
Sub addProductUnitCourse()
         On Error Resume Next
         
         
         
         
         
         Dim rs_S As New ADODB.Recordset
         If rs_S.State = 1 Then rs_S.close
         rs_S.Open "select * from  ItemCreation where CourseName='" & cboCategory.Text & "' order by ItemName", CON, adOpenDynamic, adLockOptimistic
         
         vs.Clear
         vs.Rows = rs_S.RecordCount + 1
         I = 1
         If rs_S.EOF = False Then
             
           While rs_S.EOF = False
               
               vs.TextMatrix(I, 0) = rs_S.Fields("ItemCode").value
               vs.TextMatrix(I, 2) = rs_S.Fields("ItemName").value
               vs.TextMatrix(I, 1) = rs_S.Fields("CourseName").value
               vs.TextMatrix(I, 3) = rs_S.Fields("Unit").value
               
               I = I + 1
               rs_S.MoveNext
           Wend
         
         
         
         vs.FormatString = "|Group Name|Item Name|Unit"
         vs.ColWidth(0) = 0
         vs.ColWidth(1) = 2400
         vs.ColWidth(2) = 2400
         vs.ColWidth(3) = 1100
         
         
         End If
End Sub
Sub addProduct()
         On Error Resume Next
         
         Dim rs_S As New ADODB.Recordset
         If rs_S.State = 1 Then rs_S.close
         rs_S.Open "select * from  ItemCreation where courseName='" & cboCategory.Text & "'", CON, adOpenDynamic, adLockOptimistic
         
         vs.Cols = 8
         
         vs.Clear
         vs.Rows = rs_S.RecordCount + 1
         I = 1
         If rs_S.EOF = False Then
             
           While rs_S.EOF = False
               
               vs.TextMatrix(I, 0) = rs_S.Fields(0).value
               vs.TextMatrix(I, 1) = I
               vs.TextMatrix(I, 2) = rs_S.Fields(1).value
               vs.TextMatrix(I, 3) = rs_S.Fields(2).value
               vs.TextMatrix(I, 4) = rs_S.Fields("unit").value
               vs.TextMatrix(I, 5) = Format(rs_S.Fields("price").value, "0.00")
               vs.TextMatrix(I, 6) = Format(rs_S.Fields("opening").value, "0.000")
               vs.TextMatrix(I, 7) = Format(rs_S.Fields("orderl").value, "0.000")
               I = I + 1
               rs_S.MoveNext
           Wend
         
         
         
         vs.FormatString = "|S.N.|Group Name|ItemName|Unit|Rate|Opening|Re Order"
         vs.ColWidth(0) = 0
         vs.ColWidth(1) = 500
         vs.ColWidth(2) = 1800
         vs.ColWidth(3) = 2900
         vs.ColWidth(4) = 1000
         vs.ColWidth(5) = 1200
         vs.ColWidth(6) = 1500
         vs.ColWidth(7) = 1200
         
         End If
End Sub
     
              
     
       
Private Sub cmdNew_Click()
       
       save
       addProduct
      'addProductUnitCourse
   
       
End Sub
Sub max()
    If RS.State = 1 Then RS.close
    RS.Open "select max(ItemCode) from itemcreation", CON
    If IsNull(RS.Fields(0).value) Then
        
       txtCode.Text = 1
       Else
       txtCode.Text = RS.Fields(0).value + 1
    End If
End Sub

Private Sub Cmdprint_Click()
  cr.ReportFileName = App.Path & "/ItemList.rpt"
  cr.Connect = "filedsn=chitradsn;uid=sa;pwd=sidc;"
  If cboCategory.Text <> "" Then
  cr.ReplaceSelectionFormula "{ItemCreation.CourseName}='" & cboCategory.Text & "'"
  End If
  cr.WindowState = crptMaximized
  cr.WindowShowPrintSetupBtn = True
  cr.Action = 1
End Sub

Private Sub cmdRef_Click()
  On Error Resume Next
  txtQuality.Text = ""
  'cboCategory.ListIndex = -1
  txtPrice.Text = ""
  txtperhead.Text = ""
  txtre.Text = " "
  txtrarte.Text = ""
  txtOp.Text = ""
  cmdNew.Enabled = True
  SetPerHead.value = 0
  RAW.value = 0
  max
  cboCategory.SetFocus
End Sub
Sub save()

   Dim rs_con As New ADODB.Recordset
   
   
   If txtQuality.Text = "" Or cboCategory.Text = "" Or cboUnit.Text = "" Then
      MsgBox "Please Fillup Entry !!", vbInformation, "Message"
      Exit Sub
   End If
   
   If RS.State = 1 Then RS.close
         RS.Open "select * from ItemCreation where ItemCode=" & txtCode.Text & "", CON, adOpenDynamic, adLockOptimistic
         If RS.EOF = True Then
            RS.AddNew
            RS.Fields(0).value = txtCode.Text
            RS.Fields(1).value = cboCategory.Text
            RS.Fields(2).value = txtQuality.Text
            
            RS.Fields("Unit").value = cboUnit.Text
            RS.Fields("price").value = Val(txtrarte.Text)
            
            RS.Fields("opening").value = Val(txtOp.Text)
            RS.Fields("OrderL").value = Val(txtre.Text)
            
            If rs_con.State = 1 Then rs_con.close
            rs_con.Open "select Consume_NonCon from productMaster where name='" & cboCategory.Text & "'", CON, adOpenDynamic, adLockOptimistic
            If rs_con.EOF = False Then
               RS.Fields("head").value = rs_con.Fields(0).value
            End If
            
            
            RS.update
            max
         
         Else
            
            RS.Fields(0).value = txtCode.Text
            RS.Fields(1).value = cboCategory.Text
            RS.Fields(2).value = txtQuality.Text
            
            RS.Fields("Unit").value = cboUnit.Text
            RS.Fields("price").value = Val(txtrarte.Text)
            RS.Fields("opening").value = Val(txtOp.Text)
            RS.Fields("OrderL").value = Val(txtre.Text)
            
            If rs_con.State = 1 Then rs_con.close
            rs_con.Open "select Consume_NonCon from productMaster where name='" & cboCategory.Text & "'", CON, adOpenDynamic, adLockOptimistic
            If rs_con.EOF = False Then
               RS.Fields("head").value = rs_con.Fields(0).value
            End If
            
            
            RS.update
          
         End If
         'addProduct
         
         
         txtQuality.Text = ""
         cboUnit.ListIndex = -1
         txtrarte.Text = ""
         txtre.Text = 0
         
         cboCategory.SetFocus
         
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then
    Unload Me
 End If
End Sub

Private Sub Form_Load()
addProduct
AddCategory
'Call frmBackColor(frmItemCreation)
max


If RS.State = 1 Then RS.close
RS.Open "select * from  UnitMaster", CON, adOpenDynamic, adLockOptimistic
cboUnit.Clear
If RS.EOF = False Then
  While RS.EOF = False
    cboUnit.AddItem RS.Fields(0).value
    RS.MoveNext
Wend
End If


'Call DeletePermissin(cmdDel)
'Call SavePermissin(cmdNew)

Me.Top = 250
Me.Left = 250

End Sub
Private Sub List1_Click()
txtproduct.Text = List1.Text
End Sub
Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      cboCategory.Text = grid.TextMatrix(grid.RowSel, 1)
      txtQuality.Text = grid.TextMatrix(grid.RowSel, 2)
      search1.Visible = False
   End If
End Sub

Private Sub search1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub SetPerHead_Click()
  If SetPerHead.value = 1 Then
     txtperhead.Visible = True
     head.Visible = True
   Else
     txtperhead.Visible = False
     head.Visible = False
  End If
End Sub

Private Sub txtperhead_Change()

End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
Dim B As Boolean
If KeyAscii = 13 Then
     txtre.SetFocus
End If

''B = val_int(txtPrice.Text, KeyAscii)
''If B = False Then
''   KeyAscii = 0
''End If

End Sub

Private Sub txtop_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtre.SetFocus
End Sub

Private Sub txtQuality_GotFocus()
If PopUpValue3 <> "" Then
  
  If RS.State = 1 Then RS.close
  RS.Open "select * from itemcreation where ItemCode=" & PopUpValue3 & "", CON
  If RS.EOF = False Then
  cboCategory.Text = RS!CourseName
  txtQuality.Text = RS!itemname
  txtCode.Text = RS!ItemCode
  cboUnit.Text = RS!unit
  txtrarte.Text = RS!price
  txtOp.Text = RS!Opening
  txtre.Text = RS!OrderL
  PopUpValue3 = ""
  addProduct
  End If
End If
End Sub

Private Sub txtQuality_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then

popuplist10 "select ItemName,CourseName as GpName,ItemCode from ItemCreation", CON
End If
End Sub

Private Sub txtQuality_KeyPress(KeyAscii As Integer)
 On Error Resume Next
 If KeyAscii = 13 Then
    If Check1.value = 1 Then
       SendKeys "{tab}"
      Else
       txtPrice.SetFocus
    End If
    
      
   End If

End Sub
Sub search()
  On Error Resume Next
  cboCategory.Text = vs.TextMatrix(vs.RowSel, 2)
  txtQuality.Text = vs.TextMatrix(vs.RowSel, 3)
  txtCode.Text = vs.TextMatrix(vs.RowSel, 0)
  cboUnit.Text = vs.TextMatrix(vs.RowSel, 4)
  txtrarte.Text = vs.TextMatrix(vs.RowSel, 5)
  txtOp.Text = vs.TextMatrix(vs.RowSel, 6)
  txtre.Text = vs.TextMatrix(vs.RowSel, 7)

End Sub
Private Sub txtre_GotFocus()
   txtre.SelLength = 10
End Sub

Private Sub txtre_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Call cmdNew_Click

End Sub

Private Sub txtrarte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtOp.SetFocus
End If
End Sub

Private Sub vs_Click()
search
'Call DeletePermissin(cmdDel)
'Call SavePermissin(cmdDel)

End Sub
Private Sub vs_SelChange()
 search
End Sub
