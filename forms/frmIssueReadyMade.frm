VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmIssueReadyMade 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Issue To Deptt."
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14940
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   14940
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtbillno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "0"
      Top             =   300
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Caption         =   "Show Balance Demand"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11250
      TabIndex        =   26
      Top             =   7800
      Visible         =   0   'False
      Width           =   3180
      Begin VB.ListBox Listbalance 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2190
         Left            =   360
         Style           =   1  'Checkbox
         TabIndex        =   28
         Top             =   1575
         Width           =   2655
      End
      Begin VB.CommandButton cmdshow 
         BackColor       =   &H80000013&
         Caption         =   "S&how"
         Height          =   375
         Left            =   420
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1080
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker todate1 
         Height          =   315
         Left            =   1080
         TabIndex        =   29
         Top             =   630
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   393216
         Format          =   75497473
         CurrentDate     =   38923
      End
      Begin MSComCtl2.DTPicker fromdate1 
         Height          =   315
         Left            =   1080
         TabIndex        =   30
         Top             =   300
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   393216
         Format          =   75497473
         CurrentDate     =   38923
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "To :"
         Height          =   255
         Left            =   480
         TabIndex        =   32
         Top             =   675
         Width           =   405
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "From :"
         Height          =   225
         Left            =   480
         TabIndex        =   31
         Top             =   315
         Width           =   570
      End
   End
   Begin MSDataListLib.DataCombo Cmbmedi 
      Height          =   1725
      Left            =   270
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3043
      _Version        =   393216
      Appearance      =   0
      Style           =   1
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Issue No. Searching"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8100
      Left            =   11220
      TabIndex        =   19
      Top             =   45
      Width           =   3180
      Begin VB.ComboBox cbogp 
         Height          =   315
         Left            =   180
         TabIndex        =   16
         Top             =   1575
         Width           =   2850
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H80000000&
         Caption         =   "S&earch"
         Height          =   375
         Left            =   1395
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ListBox listno 
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
         Height          =   5685
         Left            =   180
         TabIndex        =   20
         Top             =   1935
         Width           =   2865
      End
      Begin MSComCtl2.DTPicker todate 
         Height          =   315
         Left            =   1395
         TabIndex        =   21
         Top             =   675
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   393216
         Format          =   75497473
         CurrentDate     =   38923
      End
      Begin MSComCtl2.DTPicker fromdate 
         Height          =   315
         Left            =   1395
         TabIndex        =   24
         Top             =   315
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         _Version        =   393216
         Format          =   75497473
         CurrentDate     =   38923
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "From Date :"
         Height          =   225
         Left            =   480
         TabIndex        =   23
         Top             =   375
         Width           =   915
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "To Date :"
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   675
         Width           =   915
      End
   End
   Begin Crystal.CrystalReport cr 
      Left            =   225
      Top             =   8235
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Exit"
      Height          =   645
      Left            =   9465
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8385
      Width           =   1035
   End
   Begin VB.TextBox txtrem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1350
      TabIndex        =   7
      Top             =   1140
      Width           =   6345
   End
   Begin VB.TextBox txtparty 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5625
      TabIndex        =   6
      Text            =   "ChitraExport"
      Top             =   825
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   645
      Left            =   8430
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8385
      Width           =   1035
   End
   Begin VB.CommandButton cmdmodify 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Modify"
      Enabled         =   0   'False
      Height          =   645
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8385
      Width           =   1065
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   645
      Left            =   5295
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8385
      Width           =   1035
   End
   Begin VB.CommandButton cmdref 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   645
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8385
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   645
      Left            =   7395
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8385
      Width           =   1035
   End
   Begin VB.ComboBox cbodeptt 
      Height          =   315
      Left            =   1350
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   780
      Width           =   4245
   End
   Begin MSComCtl2.DTPicker dtpdate1 
      Height          =   345
      Left            =   6375
      TabIndex        =   11
      Top             =   285
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   75497475
      CurrentDate     =   38338
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   6375
      Left            =   225
      TabIndex        =   8
      Top             =   1770
      Width           =   10785
      _cx             =   19024
      _cy             =   11245
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   16711680
      BackColorFixed  =   12632256
      ForeColorFixed  =   0
      BackColorSel    =   12648447
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   16711680
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   8388608
      SheetBorder     =   8388608
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   8
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
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
      ExplorerBar     =   0
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   510
         Left            =   6210
         TabIndex        =   12
         Top             =   6375
         Width           =   2655
      End
   End
   Begin VB.Label demand 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   5730
      TabIndex        =   33
      Top             =   480
      Width           =   1275
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   18
      Top             =   1155
      Width           =   1275
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Issue No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   210
      TabIndex        =   17
      Top             =   300
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5235
      TabIndex        =   15
      Top             =   300
      Width           =   1230
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   525
      Left            =   4950
      TabIndex        =   14
      Top             =   3300
      Width           =   1245
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFCB97&
      BackStyle       =   0  'Transparent
      Caption         =   "Department  :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   13
      Top             =   855
      Width           =   1395
   End
End
Attribute VB_Name = "frmIssueReadyMade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim editflag As Boolean
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim billno As Long
Dim bal, price As Double
Dim demandNo
Private Sub cboParty_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys "{tab}"
   End If
End Sub

Private Sub cbodeptt_GotFocus()
   If PopUpValue1 <> "" Then
   cboDeptt.Text = PopUpValue1
   PopUpValue1 = ""
   End If
End Sub

Private Sub cbodeptt_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode >= 65 And KeyCode <= 122 Then
   popuplist10 "select name from deptt order by name", CON
End If
 


End Sub

Private Sub cbodeptt_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    txtrem.SetFocus
 End If
End Sub

Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then
   For J = 1 To fg.Rows - 1
       fg.TextMatrix(J, 2) = 0
   Next
End If
End Sub

Private Sub Cmbmedi_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 27 Then
      Cmbmedi.Visible = False
      fg.SetFocus
   End If
End Sub

Private Sub Cmbmedi_KeyPress(KeyAscii As Integer)
      
  On Error GoTo aa1
    
    
    If KeyAscii = 13 Then
     
    If fg.Col = 0 Then
       Cmbmedi.Visible = False
       fg.TextMatrix(fg.RowSel, 0) = Cmbmedi.Text
       
       If fg.TextMatrix(fg.RowSel, 0) <> "" Then
       fg.SetFocus
       SendKeys "{right}"
       End If
       
       If rs.State = 1 Then rs.Close
       rs.Open "select * from ItemCreation where " & stringyear & " and  CourseName='" & Cmbmedi.Text & "'", CON
       If rs.EOF = False Then
          fg.TextMatrix(fg.RowSel, 3) = rs.Fields("unit").Value
       End If
       
       
       
    Else
       Cmbmedi.Visible = False
       fg.TextMatrix(fg.RowSel, 1) = Cmbmedi.Text
       
       checkBalance
       fg.TextMatrix(fg.RowSel, 6) = bal
       fg.TextMatrix(fg.RowSel, 5) = demandNo
       fg.TextMatrix(fg.RowSel, 4) = price
       
       fg.SetFocus
       SendKeys "{right}"
       'SendKeys "{right}"
       
       If Cmbmedi.BoundText = "" Then Exit Sub
       
       If rs.State = 1 Then rs.Close
       rs.Open "select unit from ItemCreation where " & stringyear & " and  ItemCode=" & Cmbmedi.BoundText & "", CON
       If rs.EOF = False Then
          fg.TextMatrix(fg.RowSel, 3) = rs.Fields("unit").Value
       End If
    End If
      
      
    
    End If
    
    
 
 Exit Sub
aa1:
  
  MsgBox "Err " & Err.DESCRIPTION
   
    
    
End Sub
Private Sub cmdExit_Click()
CR.Reset
CR.ReportFileName = App.Path & "\Reports\issuedeptt.RPT"
CR.Connect = "FILEDSN=hotel;pwd=java;"
CR.ReplaceSelectionFormula "{issuedeppt.Deppt}='" & cboDeptt.Text & "'"
CR.WindowShowCloseBtn = True
CR.WindowShowPrintBtn = True
CR.WindowControlBox = True
CR.WindowShowPrintSetupBtn = True
CR.WindowShowProgressCtls = True
CR.WindowState = crptMaximized
CR.Action = 1
End Sub

Private Sub cmdModify_Click()

If MsgBox("Are you sure to Modify ?", vbQuestion + vbYesNo) = vbYes Then
   
   deleteFinishPurchase
   save
End If

End Sub
Sub max()
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "select max(billno) from IssueDeppt", CON
    If IsNull(rs.Fields(0).Value) Then
       txtbillno.Text = 1
       Else
       txtbillno.Text = rs.Fields(0).Value + 1
    End If
End Sub


Private Sub cmdPrint_Click()
Unload Me
End Sub

Private Sub cmdprint1_Click()

End Sub

Private Sub cmdRef_Click()
     fg.Clear
     
     SeWidth
     max
     fg.Rows = 5
     cmdSave.Enabled = True
     
     
     
     txtparty.Text = "ChitraExport"
     txtrem.Text = ""
     dtpdate1.Value = Date
  
     cboDeptt.ListIndex = -1
     For I = 0 To Listbalance.ListCount - 1
     Listbalance.Selected(I) = False
     Next
     Frame2.Enabled = True
     
     cboDeptt.Text = ""
     
     dtpdate1.SetFocus
     
End Sub

Private Sub cmdSave_Click()
  If MsgBox("Do U Want Save ?", vbQuestion + vbYesNo) = vbYes Then
     On Error GoTo save:
      save
     Exit Sub
save:
     MsgBox "" & Err.DESCRIPTION
  End If
 
End Sub
Sub save()
Dim rs1 As New ADODB.Recordset
   If txtparty.Text = "" Then
      MsgBox "Please Select Party Name !!", vbExclamation
      Exit Sub
   End If
   
   
   
   Set rs = New ADODB.Recordset
   If rs.State = 1 Then rs.Close
   rs.Open "select * from IssueDeppt where " & stringyear & " and  billno=" & txtbillno.Text & "", CON, adOpenDynamic, adLockOptimistic
   If rs.EOF = True Then
   
   For I = 1 To fg.Rows - 1
   
   If Val(fg.TextMatrix(I, 2)) > 0 Then
           
      rs.addNew
      rs!billno = Trim(txtbillno.Text)
      rs!Supplier = Trim(txtparty.Text)
      rs!Dates = dtpdate1.Value
      rs!Remarks = txtrem.Text
      rs!deppt = cboDeptt.Text
      
      rs!gp = Trim(fg.TextMatrix(I, 0))
      rs!itemname = Trim(fg.TextMatrix(I, 1))
      rs!qty = fg.TextMatrix(I, 2)
      rs!unit = Trim(fg.TextMatrix(I, 3))
      rs!price = fg.TextMatrix(I, 4)
      rs!demandNo = IIf(fg.TextMatrix(I, 5) = "", 0, fg.TextMatrix(I, 5))
      rs.Update
      
      
      End If
   
    
   Next
   
   
   
   
   
   
   Else
   
      MsgBox "This IssueDeppt No Already Exist", vbInformation
   
   End If

   Call cmdRef_Click

End Sub
Sub search()
 
 
  On Error Resume Next
   
  fg.Clear
  SeWidth
  Dim rs1 As New ADODB.Recordset
  
  
  If listno.Text <> "" Then
     txtbillno.Text = listno.Text
  End If
 
   
   Set rs = New ADODB.Recordset
   If rs.State = 1 Then rs.Close
   rs.Open "select * from IssueDeppt where " & stringyear & " and  billno=" & txtbillno.Text & " order by ItemName", CON, adOpenDynamic, adLockOptimistic
   If rs.EOF = False Then
      billno = rs!demandNo
      cmdmodify.Enabled = True
      Command2.Enabled = True
      txtparty.Text = rs!Supplier
         
      dtpdate1.Value = rs!Dates
      cboDeptt.Text = rs!deppt
      txtrem.Text = rs!Remarks
      fg.Rows = rs.RecordCount + 1
      
      For I = 1 To rs.RecordCount
      
      fg.TextMatrix(I, 0) = rs!gp
      fg.TextMatrix(I, 1) = rs!itemname
      fg.TextMatrix(I, 3) = rs!unit
      fg.TextMatrix(I, 2) = Format(rs!qty, "0.000")
      fg.TextMatrix(I, 4) = rs!price
      

      
      rs.MoveNext
      Next

   End If
        
  txtparty.SetFocus
      
  'fg.Rows = fg.Rows + 20
      
   
End Sub
Sub searchDemand()
 
 

   
End Sub
Private Sub cmdSearch_Click()
    Me.listno.Clear
    
    Set rs = Nothing
    If cbogp = "" Then
    Set rs = CON.Execute("Select BillNo from IssueDeppt where " & stringyear & " and  convert(smalldatetime,Dates,103) >= convert(smalldatetime,'" & fromdate.Value & "',103) and convert(smalldatetime,Dates,103) <= convert(smalldatetime,'" & todate.Value & "',103) Group By BillNo Order by BillNo")
    Else
    Set rs = CON.Execute("Select BillNo from IssueDeppt where " & stringyear & " and  gp='" & cbogp & "' and convert(smalldatetime,Dates,103) >= convert(smalldatetime,'" & fromdate.Value & "',103) and convert(smalldatetime,Dates,103) <= convert(smalldatetime,'" & todate.Value & "',103) Group By BillNo Order by BillNo")
    End If
    Do While Not rs.EOF = True
            Me.listno.AddItem (rs("BillNo"))
    rs.MoveNext
    Loop
End Sub
Private Sub Command2_Click()
If MsgBox("Are you sure to delete", vbQuestion + vbYesNo) = vbYes Then
   deleteFinishPurchase
   
   Call cmdRef_Click

End If

End Sub
Sub deleteFinishPurchase()
  
  
  Set mrs = Nothing
  Set mrs = CON.Execute("Delete  from IssueDeppt where " & stringyear & " and  billno=" & txtbillno.Text & "")
  
  
  cmdmodify.Enabled = False
  Command2.Enabled = False

End Sub
Private Sub CommandButton4_Click()
Unload Me
End Sub
Private Sub dtpdate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      cboParty.SetFocus
   End If
End Sub

Sub Total()
  '  txtTotal.Text = 0
  '  For i = 1 To fg.Rows - 1
  '     If fg.TextMatrix(i, 0) <> "" Then
  '        txtTotal.Text = (Val(txtTotal.Text) + Val(fg.TextMatrix(i, 5)))
  '     End If
  '  Next
    
  '  txtTotal.Text = txtTotal.Text
    
End Sub

Private Sub date1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      txtTime2.SetFocus
   End If
End Sub
Private Sub datePLA_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then txtPLArs.SetFocus
   
End Sub
Private Sub daterg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtregRs.SetFocus
End Sub
Private Sub dtpdate1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then cboDeptt.SetFocus
End Sub

Private Sub fg_EnterCell()
If Me.fg.Col = 2 Or Me.fg.Col = 4 Then
  Me.fg.Editable = flexEDKbd
Else
  Me.fg.Editable = flexEDNone
End If
End Sub

Private Sub fg_KeyDown(KeyCode As Integer, Shift As Integer)


'On Error GoTo aa1:


 If KeyCode = 13 Then
           
     If fg.Col = 1 Then
           
           
           Cmbmedi.Text = ""
           cellposi
           Dim filldata As New ADODB.Recordset
            filldata.Open "select ItemCode,ItemName from ItemCreation where " & stringyear & " and  CourseName='" & fg.TextMatrix(fg.RowSel, 0) & "' order by ItemName", CON
            
            Set Cmbmedi.RowSource = filldata
            Cmbmedi.ListField = "ItemName"
            Cmbmedi.BoundColumn = "ItemCode"
            Cmbmedi.ReFill
            
            Cmbmedi.Visible = True
            Cmbmedi.SetFocus
        End If
  End If


 



If KeyCode = 13 Then
If Me.fg.Col = 5 Then
    Me.fg.Rows = Me.fg.Rows + 1
    'Me.fg.Row = Me.fg.Row + 1
      SendKeys "{down}"
        SendKeys "{home}"
End If
End If



If KeyCode = 46 Then
   fg.RemoveItem (fg.RowSel)
   Total
End If



   If KeyCode = 13 Then
     If fg.Col = 0 Then
       Cmbmedi.Text = ""
       cellposi
       fillcmb
       fg.Editable = flexEDNone
       Cmbmedi.Visible = True
       If fg.Row >= 2 Then
          Cmbmedi.Text = fg.TextMatrix(fg.RowSel - 1, 0)
       End If
       Cmbmedi.SetFocus
     ElseIf fg.Col = 1 Then
            
       fg.Editable = flexEDNone
       SendKeys "{right}"
       'SendKeys "{right}"
     
     ElseIf fg.Col = 2 Then
       
       fg.Rows = fg.Rows + 1
       'SendKeys "{home}"
       SendKeys "{down}"
       fg.Col = 0
       fg.Editable = flexEDNone
     ElseIf fg.Col = 3 Then
       fg.Editable = flexEDKbdMouse
     End If
     
     
     
   End If

   
'  Exit Sub
'aa1:
  
'  MsgBox "" & Err.Description





End Sub

Private Sub fg_KeyUp(KeyCode As Integer, Shift As Integer)
  'If KeyCode = 13 Then
  
   ' If fg.TextMatrix(fg.RowSel, 4) <> "" Then
    ' fg.TextMatrix(fg.RowSel, 5) = (CDbl(fg.TextMatrix(fg.RowSel, 4)) * CDbl(fg.TextMatrix(fg.RowSel, 3)))
  'End If

 'End If
 'End If

End Sub


Sub cellposi()
  Cmbmedi.Width = fg.CellWidth
  Cmbmedi.TOP = fg.TOP + fg.CellTop
  Cmbmedi.Left = fg.Left + fg.CellLeft
End Sub

Private Sub Form_Activate()
'txtparty.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

'If KeyCode = 27 Then
''   If SearchFrame.Visible = True Then
''    SearchFrame.Visible = False
''    If Val(txtTotal.Text) > 0 Then
''    Call cmdSave_Click
''    End If
''    Unload Me
''    Else
''    Unload Me
''   End If
'End If
'
End Sub
Function checkBalance()
   
'''   bal = 0
'''   demandNo = ""
'''   price = 0
'''
'''   Set rs = New ADODB.Recordset
'''   If rs.State = 1 Then rs.Close
'''   rs.Open "select * from demand where " & stringyear & " and  Deppt='" & cbodeptt.Text & "'" & _
'''   " and gp='" & fg.TextMatrix(fg.RowSel, 0) & "' and ItemName='" & fg.TextMatrix(fg.RowSel, 1) & "'", CON, adOpenDynamic, adLockOptimistic
'''   If rs.EOF = False Then
'''
'''      If rs1.State = 1 Then rs1.Close
'''      rs1.Open "select sum(Qty) from issuedeppt where " & stringyear & " and  Deppt='" & cbodeptt.Text & "'" & _
'''      " and gp='" & fg.TextMatrix(fg.RowSel, 0) & "' and ItemName='" & fg.TextMatrix(fg.RowSel, 1) & "'", CON, adOpenDynamic, adLockOptimistic
'''      If Not IsNull(rs1(0)) Then
'''         bal = rs1(0)
'''      Else
'''         bal = 0
'''      End If
'''
'''      price = rs!price
'''      bal = (rs.Fields("qty").Value - bal)
'''      demandNo = rs!billno
'''
'''  End If


End Function
Private Sub Form_Load()
''Main



Me.fg.ColComboList(0) = lst

SeWidth

max

dtpdate1.Value = Date


fillcmb




'Call frmBackColor(Me)

'Call cmdSearch_Click

fromdate.Value = Date
todate.Value = Date

fromdate1.Value = Date
todate1.Value = Date

If mrs.State = 1 Then mrs.Close
mrs.Open "select distinct(CourseName) from ItemCreation", CON
While mrs.EOF = False
cbogp.AddItem mrs(0)
mrs.MoveNext
Wend



End Sub
Sub fillcmb()

Dim filldata As New ADODB.Recordset


filldata.Open "select distinct(CourseName) from ItemCreation order by CourseName", CON

Set Cmbmedi.RowSource = filldata

Cmbmedi.ListField = "CourseName"
Cmbmedi.BoundColumn = "CourseName"
Cmbmedi.ReFill

'fg.Rows = 5

End Sub
Sub SeWidth()
    
    fg.Cols = 7
    fg.FormatString = "Group|Item Name|Quantity|Unit|Price"               '>Demand No|>Balance Qty"
    fg.ColWidth(0) = 3000
    fg.ColWidth(1) = 3700
    fg.ColWidth(2) = 1200
    fg.ColWidth(3) = 1200
    fg.ColWidth(4) = 1500
    'fg.ColWidth(5) = 1400
    'fg.ColWidth(6) = 1500
    
    
End Sub
Private Sub FromDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      todate.SetFocus
   End If
End Sub

Private Sub List1_Click()
'''''''''''''''''''''''

'''                On Error Resume Next
'''                Set mrs = Nothing
'''                Set mrs = con.Execute("Select * from RawPurchaseMain where " & stringyear & " and  billno=" & Me.List1.Text & "")
'''                If mrs.EOF = True Then
'''                   MsgBox "IssueDeppt No. dose't exist.", vbInformation
'''                   cmdSave.Enabled = True
'''                Else
'''                   dtpdate.Value = mrs.Fields("Dates").Value
'''                   cboParty.Text = mrs.Fields("PartyName").Value
'''                   txtbillno.Text = mrs.Fields("billno").Value
'''                   txtTotal.Text = mrs.Fields("amt").Value
'''                   cmdSave.Enabled = False
'''                End If
'''
'''
'''                Set mrs = Nothing
'''                Set mrs = con.Execute("Select * from RawPurchase where " & stringyear & " and  billno=" & Me.List1.Text & "")
'''
'''                Me.fg.Rows = 1
'''                i = 1
'''                            Do While Not mrs.EOF = True
'''
'''                                Me.fg.Rows = Me.fg.Rows + 1
'''                                Me.fg.TextMatrix(i, 0) = mrs("Itemcode")
'''                                Me.fg.TextMatrix(i, 1) = mrs("Itemname")
'''                                Me.fg.TextMatrix(i, 2) = mrs("Unit")
'''                                Me.fg.TextMatrix(i, 3) = mrs("Qty")
'''
'''
'''                                mrs.MoveNext
'''                                i = i + 1
'''                            Loop
'''
'''                          cmdmodify.Enabled = True
'''                          Command2.Enabled = True
'''
'''                          Total
'''
'''                       ' Call DeletePermissin(Command2)
'''                       ' Call SavePermissin(cmdSave)
'''                        'Call ModifyPermissin(cmdmodify)
'''
'''
'''''''''''''''''''''''''''
End Sub

Private Sub Option1_Click()
'Me.txtchequeno.Text = ""
'Me.txtchequeno.Enabled = False
'Me.txtchequeno.Visible = False
'Me.Label5.Enabled = False
'Me.Label5.Visible = False
'
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.fg.SetFocus
Me.fg.Col = 0
End If
End Sub

Private Sub Option2_Click()
'Me.Label5.Enabled = True
'Me.Label5.Visible = True
'Me.txtchequeno.Enabled = True
'Me.txtchequeno.Visible = True
'Me.txtchequeno.SetFocus
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.fg.SetFocus
Me.fg.Col = 0
End If
End Sub

Private Sub Option3_Click()
'Me.txtchequeno.Text = ""
'Me.txtchequeno.Enabled = False
'Me.txtchequeno.Visible = False
'Me.Label5.Enabled = False
'Me.Label5.Visible = False
End Sub

Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Me.fg.SetFocus
Me.fg.Col = 0
End If
End Sub

Private Sub txtchequeno_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
'If KeyCode = 13 Then
'Me.fg.SetFocus
'Me.fg.Col = 0
'ElseIf KeyCode = vbKeyUp Then
'Me.Option1.SetFocus
'End If
End Sub

Private Sub SearchVs_KeyDown(KeyCode As Integer, Shift As Integer)
  
'''  If KeyCode = 38 Then
'''    If SearchVs.Row = 0 Then
'''     txtSearch.SetFocus
'''    End If
'''  ElseIf KeyCode = 13 Then
'''
'''     fg.SetFocus
'''
'''
'''    Set mrs = Nothing
'''    Set mrs = con.Execute("Select * from ItemCreation where " & stringyear & " and  ItemCode='" & Me.SearchVs.TextMatrix(Me.SearchVs.RowSel, 0) & "'")
'''    If Not mrs.EOF = True Then
'''       Me.fg.TextMatrix(fg.RowSel, 0) = mrs("ItemCode")
'''       Me.fg.TextMatrix(fg.RowSel, 1) = mrs("ItemName")
'''       Me.fg.TextMatrix(fg.RowSel, 2) = mrs("Unit")
'''       SendKeys "{right}"
'''    End If
'''
'''
'''
'''
'''     SearchFrame.Visible = False
'''  End If

End Sub

Private Sub Listbalance_Click()

searchDemand

End Sub

Private Sub listno_Click()
  search
  Frame2.Enabled = False
  cmdSave.Enabled = False
End Sub

Private Sub todate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then cmdSearch.SetFocus
End Sub

Private Sub txtSearch_Change()

Set mrs = Nothing
Set mrs = CON.Execute("Select ItemCode from ItemCreation where " & stringyear & " and  ItemName like '" & txtSearch.Text & "%' order by ItemName")
If mrs.EOF = False Then
   Set SearchVs.DataSource = mrs
End If

End Sub

Private Sub txtSearch_GotFocus()
  txtSearch.BackColor = &HFFC0C0
End Sub
Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 40 Then
     SearchVs.SetFocus
  End If
End Sub

Private Sub txtSearch_LostFocus()
  txtSearch.BackColor = &HFFFFFF
End Sub
Sub searchdate()

'If rs.State = 1 Then rs.Close
'rs.Open "Select * from staff where " & stringyear & " and  BrokerName='" & PopUpValue1 & "'", con, adOpenDynamic, adLockOptimistic
'If rs.EOF = False Then
'
'   txtparty.Text = rs!BrokerName
'
'   PopUpValue1 = ""
'
'End If

End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text3_Change()

End Sub

Private Sub txtdutyinword2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtTotalcases.SetFocus
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtbillno_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
      search
   End If
End Sub

Private Sub Label29_Click()

End Sub

Private Sub txtExiceDuty_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
       txtPLANo.SetFocus
   End If
End Sub

Private Sub txtgrno_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      txtPono.SetFocus
   End If
End Sub

Private Sub txtModeTr_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      txtgrno.SetFocus
   End If
End Sub

Private Sub txtparty_GotFocus()
   'searchdate
   If PopUpValue1 <> "" Then
   txtparty.Text = PopUpValue1
   PopUpValue1 = ""
   End If
End Sub

Private Sub txtparty_KeyDown(KeyCode As Integer, Shift As Integer)
''    If KeyCode >= 65 And KeyCode <= 122 Then
''       txtparty.Text = ""
''       PopupList12 "Select distinct(Name) as Sites  from CollegeMaster", CON
''    End If
    
''    If KeyCode = 13 Then
''       cbodeptt.Clear
''       If rs.State = 1 Then rs.Close
''       rs.Open "select Name from deptt order by Name", con
''       If rs.EOF = False Then
''       While rs.EOF = False
''          cbodeptt.AddItem rs.Fields(0).Value
''          rs.MoveNext
''       Wend
''       End If
''       dtpdate1.SetFocus
''    End If

    If KeyCode = 13 Then
       cboDeptt.Clear
       If rs.State = 1 Then rs.Close
       rs.Open "select distinct(Consume_NonCon) as Dept from CollegeMaster where " & stringyear & " and  name='" & txtparty.Text & "'", CON
       If rs.EOF = False Then
       While rs.EOF = False
          cboDeptt.AddItem rs.Fields(0).Value
          rs.MoveNext
       Wend
       End If
       dtpdate1.SetFocus
    End If

End Sub
Private Sub txtPLANo_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      datePLA.SetFocus
   End If
End Sub

Private Sub txtPLArs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtRG23.SetFocus
End Sub

Private Sub txtPono_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      fg.SetFocus
   End If
End Sub

Private Sub txttime_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      date1.SetFocus
   End If
End Sub

Private Sub txtrem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
    fg.SetFocus
   fg.Row = 1
    End If
End Sub

