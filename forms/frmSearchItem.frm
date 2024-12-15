VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSearchItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   5415
   ClientLeft      =   1695
   ClientTop       =   3315
   ClientWidth     =   12870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   12870
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   7065
      Top             =   5400
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
      CommandType     =   8
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
      Connect         =   "filedsn=mydsn;"
      OLEDBString     =   "filedsn=mydsn;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAs 
      BackColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   12180
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Sort Ascending (Press A)"
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdDes 
      BackColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   12180
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Sort Ascending (Press Z)"
      Top             =   585
      Width           =   600
   End
   Begin VB.CommandButton cmdFil 
      BackColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   12180
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Filter By Selection (Press F3)"
      Top             =   1170
      Width           =   600
   End
   Begin VB.CommandButton cmdUn 
      BackColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   12180
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Apply Filter  (Press F4)"
      Top             =   1755
      Width           =   600
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   12030
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   4710
      Left            =   45
      TabIndex        =   1
      Top             =   360
      Width           =   12030
      _cx             =   21220
      _cy             =   8308
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   15787206
      ForeColor       =   -2147483640
      BackColorFixed  =   12582912
      ForeColorFixed  =   16777215
      BackColorSel    =   15787206
      ForeColorSel    =   -2147483640
      BackColorBkg    =   15787206
      BackColorAlternate=   15787206
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.Label lblfooter 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   45
      TabIndex        =   2
      Top             =   5085
      Width           =   12030
   End
End
Attribute VB_Name = "frmSearchItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Boolean
Dim vsFill As New ADODB.Recordset
Dim para As String
Dim a1 As String
Dim bb As Boolean

Dim ColNumber As Integer
Dim filcombo As New ADODB.Recordset
Dim str1 As String
Dim rs1 As New ADODB.Recordset

Private Sub cboCombo_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 27 Then
     cboCombo.Visible = False
  End If
End Sub
Private Sub cmdAs_Click()
    
Screen.MousePointer = vbHourglass
    

If RS.State = 1 Then RS.close
RS.Open "select * from Querydes where tables_No='" & tblNo & "'", CON
If RS.EOF = False Then
   para = RS.Fields("" & "para" & vs.col + 1).value
   If vsFill.State = 1 Then vsFill.close
    
   If Len(RS!para6) > 0 Then
      vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & " where " & RS!para6 & " group by " & RS!fld1 & "  order by " & para & " asc", CON
   Else
      vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & "  group by " & RS!fld1 & "  order by " & RS!fld1 & " asc", CON
   End If
   Set vs.DataSource = vsFill
   
   setColHead
End If

showFooter

Screen.MousePointer = vbDefault

Me.Caption = "Ascending Successfully ..."

End Sub
Private Sub cmdDes_Click()

Screen.MousePointer = vbHourglass


If RS.State = 1 Then RS.close
RS.Open "select * from Querydes where tables_No='" & tblNo & "'", CON
If RS.EOF = False Then
   para = RS.Fields("" & "para" & vs.col + 1).value
   
   If vsFill.State = 1 Then vsFill.close
   If Len(RS!para6) > 0 Then
      vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & " where " & RS!para6 & " group by " & RS!fld1 & " order by " & para & " desc", CON
   Else
      vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & "  group by " & RS!fld1 & "  order by " & para & " desc", CON
   End If
   
   Set vs.DataSource = vsFill
   setColHead
End If

showFooter
Screen.MousePointer = vbDefault

Me.Caption = "Descending Successfully ..."

End Sub
Sub setColHead()
     
    If tblNo = 1 Or tblNo = 10 Then
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       'vs.Refresh
       vs.TextMatrix(0, 0) = "Name"
    End If
     
End Sub

Private Sub cmdFil_Click()

s1 = ""


If RS.State = 1 Then RS.close
RS.Open "select * from Querydes where tables_No='" & tblNo & "'", CON, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
    
    para = RS.Fields("" & "para" & vs.col + 1).value
    d1 = RS.Fields("" & "para" & vs.col + 1).Name
    
    If s1 = "" Then
       s1 = para & " Like " & "'" & txtSearch & "%'"
       RS.Fields("para6").value = s1
       RS.update
    Else
       s1 = para & " Like " & "'" & txtSearch & "%'"
       RS.Fields("para6").value = s1 & " or " & para & " Like " & "'" & txtSearch & "%'"
       RS.update
    End If
    
    RS.Requery
    If vsFill.State = 1 Then vsFill.close
    vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & " where " & para & " like '" & txtSearch & "%'" & " or " & RS!para6 & " group by " & RS!fld1
    Set vs.DataSource = vsFill
    
    setColHead
    
    'cboCombo.Visible = False
    
    
    
End If

showFooter
Me.Caption = "Filtered Successfully ..."


End Sub
Private Sub cmdUn_Click()

Screen.MousePointer = vbHourglass

CON.Execute "update Querydes set para6='' where tables_No='" & tblNo & "'"
If RS.State = 1 Then RS.close
RS.Open "select * from Querydes where tables_No='" & tblNo & "'", CON
If RS.EOF = False Then
   para = RS.Fields("" & "para" & vs.col + 1).value
   If vsFill.State = 1 Then vsFill.close
   vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & " group by " & RS!fld1 & "", CON, adOpenDynamic, adLockReadOnly
   Set vs.DataSource = vsFill
   setColHead
   
End If
showFooter
Me.Caption = "Unfiltered Successfully ..."


Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
vs.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then
      Unload frmSearchItem
      
      PopUpValue1 = ""
   End If
   
End Sub
Private Sub Form_Load()

'txtSearch = popupvalue4

'PopUpValue5 = ""

Set RS = New ADODB.Recordset


'=======================================================================
'=======================================================================
If popupvalue5 <> "" Then
   txtSearch = popupvalue5
   popupvalue5 = ""
End If


If REPORT_parameter <> "R" Then
If RS.State = 1 Then RS.close
RS.Open "select * from Querydes where tables_No='" & tblNo & "'", CON
If RS.EOF = False Then
   
   
   
   para = RS.Fields("" & "para" & vs.col + 1).value
   If vsFill.State = 1 Then vsFill.close
   
   If RS!haveyData = "y" Then
      vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & " where " & para & " like '" & "A" & "%' group by " & RS!fld1 & "  order by " & para, CON, adOpenKeyset, adLockReadOnly
   Else
     Set vsFill = CON.Execute("select " & RS!fld1 & " from " & RS!tblName & " where " & para & " like '" & txtSearch & "%' group by " & RS!fld1 & " order by " & para)
   End If
   Set vs.DataSource = vsFill
   
   setColHead
   showFooter


End If

'====================================================================================

Else

'===================================================================================

RS.Open "select * from Querydes where tables_No='" & tblNo & "'", CON
If RS.EOF = False Then
   
   
   para = RS.Fields("" & "para" & vs.col + 1).value
   If vsFill.State = 1 Then vsFill.close
   
   
   
   If RS!haveyData = "y" Then
      
   If REPORT_str <> "" Then
      'REPORT_str = REPORT_str & " and " & para & " like '" & "A" & "%'"
    Else
      REPORT_str = para & " like '" & "A" & "%'"
   End If
      
      vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & " where " & REPORT_str & " group by " & RS!fld1 & "  order by " & para, CON, adOpenKeyset, adLockReadOnly
   Else
   
   If REPORT_str <> "" Then
     'REPORT_str = REPORT_str & " " & para & " like '" & txtSearch & "%'"
   Else
     REPORT_str = para & " like '" & txtSearch & "%'"
   End If
     
     Set vsFill = CON.Execute("select " & RS!fld1 & " from " & RS!tblName & " where " & REPORT_str & " group by " & RS!fld1 & " order by " & para)
   End If
   Set vs.DataSource = vsFill
   
   setColHead
   showFooter


End If
'==================================================================================
'==================================================================================
End If




Me.Top = MainMenu.Top + 4000
Me.Left = MainMenu.Left + 1000

End Sub
Sub showFooter()
lblfooter.Caption = " TOTAL RECORD : " & vs.Rows - 1
End Sub
Private Sub Form_Unload(cancel As Integer)
CON.Execute "update Querydes set para6='' where tables_No='" & tblNo & "'"
End Sub
Function check_id(s As String) As String
check_id = "1"

On Error GoTo a1:

If rs1.State = 1 Then rs1.close
rs1.Open "select top 100 fyear from " & s & "", CON, adOpenStatic, adLockReadOnly

Exit Function
a1:
check_id = "2"

End Function

Private Sub txtSearch_Change()
'txtSearch = UCase(txtSearch)

str1 = "1"

If bb = True Then Exit Sub

Dim itemfound As VSFlexGrid



'=====================================================================

If REPORT_parameter <> "R" Then


If RS.State = 1 Then RS.close
RS.Open "select * from Querydes where tables_No='" & tblNo & "'", CON
If RS.EOF = False Then
    
  str1 = check_id("" & RS!tblName)
  
  para = RS.Fields("" & "para" & vs.col + 1).value
    
  If vsFill.State = 1 Then vsFill.close
    
  If str1 = "1" Then
    
    If Len(RS!para6) > 0 Then
       vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & " where " & stringyear & "  and " & para & " like '" & txtSearch.Text & "%'" & " and " & RS!para6 & " group by " & RS!fld1 & "", CON
    Else
       vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & " where  " & stringyear & "  and  " & para & " like '" & txtSearch.Text & "%'  group by " & RS!fld1 & "", CON
    End If
  
  Else
  
    If Len(RS!para6) > 0 Then
       vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & " where " & para & " like '" & txtSearch.Text & "%'" & " and " & RS!para6 & " group by " & RS!fld1 & "", CON
    Else
       vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & " where  " & para & " like '" & txtSearch.Text & "%'  group by " & RS!fld1 & "", CON
    End If
  
  End If
    
    
    
    Set vs.DataSource = vsFill
    
   
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       
       If vs.row > 0 Then
          vs.Select 1, vs.col, 1, vs.col
          vs.CellBorder RGB(255, 0, 0), 2, 3, 2, 2, 1, 1
     
       End If
       DoEvents
       DoEvents
 
    
    
    setColHead
    showFooter
    
End If

'=========================================================================
Else
'REPORT_str = ""



If RS.State = 1 Then RS.close
RS.Open "select * from Querydes where tables_No='" & tblNo & "'", CON
If RS.EOF = False Then
    
    para = RS.Fields("" & "para" & vs.col + 1).value
    
    str1 = check_id("" & RS!tblName)
    
  If vsFill.State = 1 Then vsFill.close
  
  If str1 = "2" Then
  
    If Len(RS!para6) > 0 Then
       vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & " where  " & para & " like '" & txtSearch.Text & "%'" & " and " & RS!para6 & " group by " & RS!fld1 & "", CON
    Else
       vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & " where " & para & " like '" & txtSearch.Text & "%'  group by " & RS!fld1 & "", CON
    End If
  
  Else
  
    If Len(RS!para6) > 0 Then
       vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & " where " & stringyear & "  and   " & para & " like '" & txtSearch.Text & "%'" & " and " & RS!para6 & " group by " & RS!fld1 & "", CON
    Else
       vsFill.Open "select " & RS!fld1 & " from " & RS!tblName & " where " & stringyear & "  and   " & para & " like '" & txtSearch.Text & "%'  group by " & RS!fld1 & "", CON
    End If
  
  
  End If
    
   Set vs.DataSource = vsFill
    
   
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       DoEvents
       
       If vs.row > 0 Then
          vs.Select 1, vs.col, 1, vs.col
          vs.CellBorder RGB(255, 0, 0), 2, 3, 2, 2, 1, 1
     
       End If
       DoEvents
       DoEvents
 
    
    
    setColHead
    showFooter
    
End If



End If




End Sub

Private Sub txtSearch_GotFocus()
If bb = True Then
   txtSearch.SelLength = 50
End If
c = True
HIT

End Sub
Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
   
   
   
If KeyCode = 114 Then
   Call cmdFil_Click
ElseIf KeyCode = 112 Then
   Call cmdUn_Click
ElseIf KeyCode = 115 Then
   Call cmdAs_Click
ElseIf KeyCode = 116 Then
   Call cmdDes_Click
End If
            
            
            

If (KeyCode = 37 Or KeyCode = 40 Or KeyCode = 13) Then
      vs.SetFocus
      vs.col = ColNumber
      vs.Cell(flexcpBackColor, vs.RowSel, ColNumber) = &HC00000
      vs.Cell(flexcpForeColor, vs.RowSel, ColNumber) = vbWhite
      vs.Redraw = flexRDDirect
 End If



  
   
End Sub
Private Sub txtSearch_LostFocus()
c = False
End Sub

Private Sub vs_DblClick()


      PopUpValue1 = vs.TextMatrix(vs.RowSel, vs.Cols - 1)
      PopUpValue2 = vs.TextMatrix(vs.RowSel, 0)
   
   If vs.Cols - 1 = 2 Then
      PopUpValue3 = vs.TextMatrix(vs.RowSel, 1)
   End If
   
   If vs.Cols - 1 = 3 Then
      PopUpValue3 = vs.TextMatrix(vs.RowSel, 1)
      popupvalue4 = vs.TextMatrix(vs.RowSel, 2)
   End If
   
   
   If vs.Cols - 1 = 4 Then
      PopUpValue3 = vs.TextMatrix(vs.RowSel, 1)
      popupvalue4 = vs.TextMatrix(vs.RowSel, 2)
      popupvalue5 = vs.TextMatrix(vs.RowSel, 3)
      
   End If
   
  
   
    Unload frmSearchItem
   
   'frmSearchItem.Visible = False
  
   
End Sub

Private Sub vs_GotFocus()
bb = True
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
         
         
   If KeyCode = 114 Then
      Call cmdFil_Click
   ElseIf KeyCode = 112 Then
      Call cmdUn_Click
   ElseIf KeyCode = 115 Then
      Call cmdAs_Click
   ElseIf KeyCode = 116 Then
      Call cmdDes_Click
   End If

  If (KeyCode >= 65 And KeyCode <= 90) Or (KeyCode >= 97 And KeyCode <= 122) Or (KeyCode = 32) Or (KeyCode = 8) Then
     bb = False
  End If




If (KeyCode >= 65 And KeyCode <= 90) Or (KeyCode = 8) Or KeyCode >= 48 And KeyCode <= 57 Or KeyCode = 8 Then
    c = False
    'txtSearch = Mid(vs.TextMatrix(vs.RowSel, vs.Col), 1, 1)
    If popupvalue4 <> "" Then
    If Chr(KeyCode) = 13 Then Exit Sub
      txtSearch = txtSearch + LCase(Chr(KeyCode))
      popupvalue4 = ""
    Else
      txtSearch = Chr(KeyCode)
    End If
    
    ColNumber = vs.col
    txtSearch.SetFocus
    SendKeys "{right}"
    Exit Sub
End If





''If KeyCode = 113 Then
''
''
''  cboCombo.Text = ""
''  cboCombo.Visible = True
''  cboCombo.ZOrder
''  cboCombo.Width = vs.ColWidth(vs.Col)
''  cboCombo.Left = vs.CellLeft
''  cboCombo.Top = vs.Top + vs.CellTop
''
''
''If rs.State = 1 Then rs.Close
''rs.Open "select * from Querydes where tables_No='" & tblNo & "'", CON
''If rs.EOF = False Then
''   para = rs.Fields("" & "para" & vs.Col + 1).Value
''   Adodc1.RecordSource = "select " & para & " from " & rs!tblName & " order by " & para
''
''   Adodc1.Refresh
''   cboCombo.ReFill
''   cboCombo.ListField = "" & para
''   cboCombo.SetFocus
''End If
''
''End If



If KeyCode = 13 Then
 
      PopUpValue1 = vs.TextMatrix(vs.RowSel, vs.Cols - 1)
      PopUpValue2 = vs.TextMatrix(vs.RowSel, 0)
   If vs.Cols - 1 = 2 Then
      PopUpValue3 = vs.TextMatrix(vs.RowSel, 1)
   End If
   
   If vs.Cols - 1 = 3 Then
      PopUpValue3 = vs.TextMatrix(vs.RowSel, 1)
      popupvalue4 = vs.TextMatrix(vs.RowSel, 2)
   End If
   
   Unload frmSearchItem
End If

End Sub
Private Sub vs_LeaveCell()
  'vs.Cell(flexcpBackColor, vs.RowSel, vs.Col) = &HF0E4C6
 vs.Cell(flexcpBackColor, vs.RowSel, vs.col) = &HF0E4C6
 vs.Cell(flexcpForeColor, vs.RowSel, vs.col) = vbBack
 If bb = True Then
    vs.Select vs.RowSel, vs.col, vs.RowSel, vs.col
    vs.CellBorder &HF0E4C6, 2, 3, 2, 2, 1, 1
End If
End Sub
Private Sub vs_LostFocus()
bb = False
End Sub
Private Sub vs_SelChange()
On Error Resume Next
vs.Cell(flexcpBackColor, vs.RowSel, vs.col) = &HC00000
vs.Cell(flexcpForeColor, vs.RowSel, vs.col) = vbWhite
If bb = True Then
  txtSearch.Text = vs.TextMatrix(vs.RowSel, vs.col)
End If

End Sub
