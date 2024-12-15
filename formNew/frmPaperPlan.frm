VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPaperPlan 
   ClientHeight    =   9852
   ClientLeft      =   60
   ClientTop       =   396
   ClientWidth     =   19860
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9852
   ScaleWidth      =   19860
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport cr 
      Left            =   5256
      Top             =   9360
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtRem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1008
      MaxLength       =   150
      TabIndex        =   26
      Top             =   720
      Width           =   4896
   End
   Begin VB.TextBox txtpath 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      Left            =   5976
      MaxLength       =   150
      TabIndex        =   24
      Top             =   720
      Width           =   6696
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   10476
      Top             =   9180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAddFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add &File"
      Height          =   576
      Left            =   5976
      Picture         =   "frmPaperPlan.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   72
      Width           =   1236
   End
   Begin VB.CommandButton cmdImportExcel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Import Excel Data"
      Enabled         =   0   'False
      Height          =   576
      Left            =   7272
      Picture         =   "frmPaperPlan.frx":073D
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   72
      Width           =   1452
   End
   Begin VB.CommandButton cmdRepQty 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export to Excel"
      Height          =   384
      Left            =   11340
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9252
      Visible         =   0   'False
      Width           =   444
   End
   Begin VB.CommandButton cmdOp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add &Paper Opening"
      Height          =   444
      Left            =   12024
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9216
      Visible         =   0   'False
      Width           =   144
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
      Height          =   444
      Left            =   11808
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9216
      Visible         =   0   'False
      Width           =   168
   End
   Begin VB.TextBox txtOrdNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   6
      Top             =   228
      Width           =   1044
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFD7AE&
      Height          =   816
      Left            =   12732
      TabIndex        =   1
      Top             =   216
      Width           =   6828
      Begin VB.CommandButton cmdprin2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Printer Wise (Print) - 2"
         Height          =   585
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   180
         Width           =   1128
      End
      Begin VB.CommandButton cmddel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete File"
         Height          =   585
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   180
         Width           =   1020
      End
      Begin VB.CommandButton cmdNewFile 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add New File"
         Height          =   585
         Left            =   36
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   180
         Width           =   1128
      End
      Begin VB.CommandButton cmdBillCancel 
         Caption         =   "&Order Cancel"
         Height          =   585
         Left            =   12240
         TabIndex        =   5
         Top             =   165
         Width           =   75
      End
      Begin VB.CommandButton CommandQuit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         Height          =   588
         Left            =   5724
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   1020
      End
      Begin VB.CommandButton Printcmd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Paper Size Wise (Print)"
         Height          =   585
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1164
      End
      Begin VB.CommandButton ok 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Printer Wise (Print) - 1"
         Height          =   585
         Left            =   1176
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   1128
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7980
      Left            =   72
      TabIndex        =   0
      Top             =   1104
      Width           =   19476
      _cx             =   34354
      _cy             =   14076
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   12582847
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483645
      BackColorAlternate=   -2147483643
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
      Rows            =   2
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
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
   Begin Crystal.CrystalReport cr1 
      Left            =   13848
      Top             =   9228
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComCtl2.DTPicker txtOrdDate 
      Height          =   312
      Left            =   2604
      TabIndex        =   7
      Top             =   228
      Width           =   1308
      _ExtentX        =   2307
      _ExtentY        =   550
      _Version        =   393216
      CalendarBackColor=   16776960
      Format          =   513474561
      CurrentDate     =   38372
   End
   Begin VSFlex7Ctl.VSFlexGrid vs2 
      Height          =   348
      Left            =   12420
      TabIndex        =   17
      Top             =   9180
      Visible         =   0   'False
      Width           =   1020
      _cx             =   1799
      _cy             =   609
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   12582847
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483645
      BackColorAlternate=   -2147483643
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
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
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   14640
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   6360
         Width           =   195
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   252
      Left            =   180
      TabIndex        =   27
      Top             =   720
      Width           =   1152
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PLAN No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   252
      Left            =   180
      TabIndex        =   16
      Top             =   228
      Width           =   1152
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   " Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   312
      Left            =   2112
      TabIndex        =   15
      Top             =   228
      Width           =   1380
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press F3 Search For"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   14
      Top             =   4980
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 Search For English"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   4980
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press F4 Key to delete a record"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   9135
      Width           =   2805
   End
   Begin VB.Label txtTReam 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   336
      Left            =   15048
      TabIndex        =   11
      Top             =   9312
      Visible         =   0   'False
      Width           =   828
   End
   Begin VB.Label txtTSheet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   336
      Left            =   15888
      TabIndex        =   10
      Top             =   9312
      Visible         =   0   'False
      Width           =   828
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   14448
      TabIndex        =   9
      Top             =   9372
      Visible         =   0   'False
      Width           =   672
   End
   Begin VB.Label lblPaper_det 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12900
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmPaperPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mrpeat As Boolean
Public orderchk As Boolean
Public partchk As Boolean
Public bindchk As Boolean
Dim RS As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Dim rs2_ As New ADODB.Recordset
Dim rs1_ As New ADODB.Recordset
Dim qty, wast_

Dim page_sum As Double
Public gridchk As Boolean
Public bo_ As Boolean
Dim sheet, ream, westage, Cover_sheet, Cover_ream
Public mode As String
Dim flag As Boolean

Sub paperstatementcalc()
End Sub
Sub try()
gridchk = True
vs.Row = 1
vs.Col = 8
If vs.text = "" Then
gridchk = False
vs.SetFocus
End If
End Sub
Sub part()
partchk = True
End Sub
Sub order()
End Sub
Sub grid_ini()

    
    
    
    Me.vs.Cols = 19
    
    vs.FormatString = "SNo|Books|Class|Pages|Forms|Bsize|||PrintRun|Price|Paper|PSize||Printer|PaperCons.|CoverPrinter|CoverSize|CoverPaper|CoverCons."
    
    'Me.vs.rows = 2
    Me.vs.ColWidth(0) = 600
    Me.vs.ColWidth(1) = 2300
    Me.vs.ColWidth(2) = 900
    Me.vs.ColWidth(3) = 600
    Me.vs.ColWidth(4) = 600
    Me.vs.ColWidth(5) = 900
    Me.vs.ColWidth(6) = 0
    Me.vs.ColWidth(7) = 0
    Me.vs.ColWidth(8) = 800
    Me.vs.ColWidth(9) = 600
    Me.vs.ColWidth(10) = 1800
    Me.vs.ColWidth(11) = 1700
    Me.vs.ColWidth(12) = 1000
    Me.vs.ColWidth(12) = 0
    Me.vs.ColWidth(13) = 1200
    Me.vs.ColWidth(14) = 1100
    
    Me.vs.ColWidth(15) = 1400
    Me.vs.ColWidth(16) = 1200
    Me.vs.ColWidth(17) = 1800
    Me.vs.ColWidth(18) = 1200
   
       
   
       
    
    
End Sub
Sub adddisab()
'Me.Edit.Enabled = False
'Me.Printcmd.Enabled = False
'Me.cmdBillCancel.Enabled = False
End Sub

Private Sub bill_no_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
popuplist1 "Select bill_id as [Bill No],dat as [Date] from billmaster where " & stringyear & " and categories='Main' order by cint(bill_id) asc", con
End If
End Sub

Private Sub bill_no_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   sendkeys "{tab}"
End If

End Sub

Private Sub bill_no_LostFocus()
End Sub
Private Sub Binder_id_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
If Binder_id.text = "" Then
MsgBox "Please Choose the Printers"
Binder_id.SetFocus
Exit Sub
Else
'Me.vs.Col = 1

If vs.Enabled = True Then
vs.SetFocus
End If
'SendKeys "{tab}"
End If

End If

End Sub

Private Sub Binder_id_LostFocus()
Label10.Visible = False
End Sub
Private Sub binder_name_Click()
westage = 0
If RS.State = 1 Then RS.close
RS.Open "select Address,westage from Godownmaster  where " & stringyear & " and godwn='" & binder_name & "'", con, adOpenKeyset, adLockReadOnly
If RS.EOF = False Then
    lblAdd.Caption = RS(0) & ""
    westage = RS(1)
Else
    lblAdd.Caption = ""
End If

End Sub
Private Sub binder_name_GotFocus()

If PopUpValue1 <> "" Then
   binder_name = PopUpValue1
   lblAdd = PopUpValue2
   PopUpValue1 = ""
   PopUpValue2 = ""
End If

End Sub

Private Sub binder_name_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   'popuplist1 "select Godwn as [Printer Name],Address from Godownmaster where Binder_Printer<>'g' order by Godwn", con
   popuplist1 "select Godwn as [Printer Name],Address from Godownmaster where  Len(Godwn) > 5 order by Godwn", con
   
End If
End Sub

Private Sub binder_name_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'
'    If binder_name.text = "" Then
'       MsgBox "Please Choose the Printers", vbInformation
'       Me.binder_name.SetFocus
'       Exit Sub
'    End If
'
'    sendkeys "{tab}"
'
'
'End If

End Sub


Private Sub cmdAddFile_Click()

maxNo_
vs.Clear
grid_ini
txtrem.text = ""
txtpath.text = ""


cd.ShowOpen
txtpath.text = cd.filename

cmdImportExcel.Enabled = True

End Sub

Private Sub cmdBillCancel_Click()
   If Me.txtOrdNo.text = "" Then Exit Sub
   
   If RS.State = 1 Then RS.close
   RS.Open "select * from billmaster where " & stringyear & " and firm_id = '" + "Chitra" + "' and bill_id = '" + Me.txtOrdNo.text + "'", con
   If RS.EOF = True Then
      MsgBox "Bill No Already Exist !!", vbInformation
      Exit Sub
   End If
   
   If RS.State = 1 Then RS.close
   RS.Open "select * from billmaster where " & stringyear & " and firm_id = '" + "Chitra" + "' and bill_id = '" + Me.txtOrdNo.text + "'", con, adOpenDynamic, adLockOptimistic
   If RS.EOF = False Then
   If MsgBox("Want To Order Cancel", vbQuestion + vbYesNo) = vbYes Then
      RS!OrderCancel = "Yes"
      RS.update
   End If
   End If
End Sub

Private Sub Command1_Click()
flag = True
frmbook.Show
End Sub

Private Sub Command2_Click()
flag = True
printingmaster.Option3.value = True
printingmaster.mname.Caption = "Size Master"
printingmaster.Show
End Sub
Private Sub Command3_Click()
flag = True
CustomerMaster.Show
End Sub
Private Sub Command4_Click()
flag = True
BinderMaster.Show
End Sub
Sub Ream_Sheet(Form_ As String, quan_ As String)

Dim Tot, wastage_per, Form, wream


wastage_per = 0
wsheet = 0
Form = Val(Form_)
quan = Val(quan_)
wastage_per = 5

Tot = Form * quan
Tot = (((105 / 100) * Tot) / 1000)

'If Val(wastage_per) > 0 Then
'   a1 = (Tot * wastage_per / 100)
'   wream = Int(a1)
'
'   ''wsheet = Round((a1 - Int(a1)) * 500)
'   aa1 = Int(a1)
'   aa2 = a1
'
'   wsheet = Round(aa2 - aa1, 2)
'
'End If




  
If Tot > 0 Then
   
   ream = Int(Tot)
   
   sheet = Round((Tot - ream), 3)
   
   'sheet = sheet + wsheet
   
   
   'ream = ream + wream
   
'   If sheet > 499 Then
'      wream = Int(sheet / 500)
'      sheet = sheet - (wream * 500)
'      ream = ream + wream
'   End If
 
  
End If



End Sub
Sub Cover_ReamSheet(Form_ As String, quan_ As String)

Dim Tot, wastage_per, Form, wream

 
wastage_per = 0
wsheet = 0
Form = Val(Form_)
quan = Val(quan_)
wastage_per = 5

'Tot = Form * quan
Tot = ((105 / 100) * (quan / 4) / 500)


'If Val(wastage_per) > 0 Then
'   a1 = (Tot * wastage_per / 100)
'   wream = Int(a1)
'    wsheet = Round((a1 - Int(a1)) * 500)
'End If


  
If Tot > 0 Then
   
   Cover_ream = Int(Tot)
   Cover_sheet = Round((Tot - Int(Tot)), 3)
   'Cover_sheet = Cover_sheet + wsheet
   'Cover_ream = Cover_ream + wream
   
   'If Cover_sheet > 499 Then
   '   wream = Int(Cover_sheet / 500)
   '   Cover_sheet = Cover_sheet - (wream * 500)
   '   Cover_ream = Cover_ream + wream
   'End If
 
  
End If









End Sub


Private Sub cmdDel_Click()
If MsgBox("Want to Delete File ", vbYesNo) = vbYes Then
con.Execute "delete from PaperConsumptionPlan where id='" & txtOrdNo.text & "'"

maxNo_
vs.Clear
grid_ini
txtrem.text = ""
txtpath.text = ""

End If
End Sub

Private Sub cmdImportExcel_Click()


Dim sconn As String
Dim I As Integer

sFile = Me.txtpath.text
sconn = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & sFile

txtTotQty = 0
txtBillQty = 0

Dim rs_fatch As New ADODB.Recordset
Dim rs_em As New ADODB.Recordset


dis = 0
I = 0
k1 = 1

Dim sno As Integer
Dim Qty_, forms_


sno = 1
Qty_ = ""
forms_ = ""

Set rs_em = New ADODB.Recordset
rs_em.Open "select * from PaperConsumptionPlan where id=" & txtOrdNo.text & "", con, adOpenDynamic, adLockOptimistic
If rs_em.EOF = False Then
   con.Execute "delete from PaperConsumptionPlan where id='" & txtOrdNo.text & "'"
End If


Set RS = New ADODB.Recordset
RS.Open "SELECT * FROM [sheet1$]", sconn
While RS.EOF = False

ream = 0
sheet = 0
Cover_ream = 0
Cover_sheet = 0




v1 = IIf(IsNull(RS(0)), 0, RS(0))

forms_ = IIf(IsNull(RS(5)), 0, RS(5))    'form 5
Qty_ = IIf(IsNull(RS(13)), 0, RS(15))   '11 qty

If (v1 > 0) Then


If (Qty_ > 0 And forms_ > 0) Then
  Ream_Sheet RS(5), RS.Fields(15).value
End If


If (Qty_ > 0) Then
   Cover_ReamSheet RS(5), RS.Fields(15).value



sum1 = ream + sheet
sum1_ = Cover_ream + Cover_sheet

con.Execute "insert into PaperConsumptionPlan(Id,Dates,Sno,Books,Class,Pages,Forms,BSize,PrintRun," & _
"Price,Paper,PSize,Printer,Consumption,FileDesc,ream,sheet,CoverPrinter,CoverSize,CoverPaper,CoverConsumption,CoverReam,CoverSheet) " & _
" values('" & txtOrdNo.text & "','" & Format(txtOrdDate.value, "MM/dd/yyyy") & "','" & sno & "','" & RS(2) & "','" & RS(3) & "'," & _
"'" & RS(4) & "','" & RS(5) & "','" & RS(16) & "','" & RS.Fields(15).value & "','" & RS(17) & "'," & _
"'" & RS(18) & "','" & RS(19) & "','" & RS(21) & "','" & sum1 & "','" & txtrem.text & "'," & _
"'" & ream & "','" & sheet & "','" & RS(23) & "','" & RS(24) & "','" & RS(25) & "','" & sum1_ & "','" & Cover_ream & "','" & Cover_sheet & "')"
'RS(9)  7     11 =>13
sno = sno + 1

End If



End If




RS.MoveNext
Wend


Dim vsFill As New ADODB.Recordset
Set vsFill = New ADODB.Recordset
d1 = 1
Me.vs.Clear
vs.rows = 2

vsFill.Open "select SNo,Books,Class,Pages,Forms,BSize,Vender,Remarks,PrintRun," & _
"Price,Paper,PSize,Colour,Printer,Consumption,CoverPrinter,CoverSize,CoverPaper," & _
"CoverConsumption from PaperConsumptionPlan where Id='" & txtOrdNo.text & "' order by Sno", con
While vsFill.EOF = False

vs.TextMatrix(d1, 0) = vsFill!sno
vs.TextMatrix(d1, 1) = vsFill!Books
vs.TextMatrix(d1, 2) = vsFill!Class
vs.TextMatrix(d1, 3) = vsFill!Pages
vs.TextMatrix(d1, 4) = vsFill!Forms
vs.TextMatrix(d1, 5) = vsFill!bsize
'vs.TextMatrix(d1, 6) = vsFill!Vender
'vs.TextMatrix(d1, 7) = vsFill!remarks
vs.TextMatrix(d1, 8) = vsFill!PrintRun
vs.TextMatrix(d1, 9) = vsFill!Price

vs.TextMatrix(d1, 10) = vsFill!paper
vs.TextMatrix(d1, 11) = vsFill!PSize
'vs.TextMatrix(d1, 12) = vsFill!Colour
vs.TextMatrix(d1, 13) = vsFill!Printer
vs.TextMatrix(d1, 14) = vsFill!Consumption

vs.TextMatrix(d1, 15) = vsFill!CoverPrinter & ""
vs.TextMatrix(d1, 16) = vsFill!Coversize & ""
vs.TextMatrix(d1, 17) = vsFill!CoverPaper & ""
vs.TextMatrix(d1, 18) = vsFill!CoverConsumption & ""







vs.rows = vs.rows + 1
vsFill.MoveNext
d1 = d1 + 1

Wend


grid_ini



MsgBox "Data import Successfully", vbInformation




End Sub

Private Sub cmdNewFile_Click()

maxNo_
vs.Clear
grid_ini

txtrem.text = ""
txtpath.text = ""

End Sub
Sub maxNo_()

Set RS = New ADODB.Recordset
RS.Open "select max(id) from PaperConsumptionPlan", con, adOpenDynamic, adLockOptimistic
If Not IsNull(RS.Fields(0).value) Then
   txtOrdNo.text = RS(0) + 1
Else
   txtOrdNo.text = 1
End If

End Sub

Private Sub cmdprin2_Click()
Dim reem, sheet, a1, per
con.Execute "delete from tmps_LEDGER1"

sheet = 0


c_ream = 0
c_sheet = 0


'If RS.State = 1 Then RS.close
'RS.Open "select  Printer,Paper,sum(CoverReam),sum(CoverSheet)  from PaperConsumptionPlan where ID='" & txtOrdNo.text & "' group by Printer,Paper", con



If rs1.State = 1 Then rs1.close
rs1.Open "select  Printer,Paper,PSize,sum(Ream),sum(Sheet) size from PaperConsumptionPlan WHERE ID='" & txtOrdNo.text & "'  group by Printer,Paper,PSize"
While rs1.EOF = False
   
   sheet = 0
   a1 = rs1(4)
   ream = Int(rs1(3))
   
   sheet = a1 - Int(a1)
   ream = ream + Int(a1)
   
   '-------------------------------------------------------
   c_ream = 0
   c_sheet = 0
   
   
'   RS.MoveFirst
'   RS.Find "Paper='" & rs1!paper & "'"
'   If RS.EOF = False Then
'      a1_ = RS(3)
'      c_ream = Int(RS(2))
'      c_sheet = a1_ - Int(a1_)
'      c_ream = c_ream + Int(a1_)
'   End If
   
   '-------------------------------------------------------

   
   con.Execute "insert into tmps_LEDGER1(SUBLEDGER,DESCFORINVOICE,address1,address2,address3,phone,owner,setupid,fyear) values('" & rs1!Printer & "','" & rs1!PSize & "','" & rs1!paper & "','" & ream & "','" & sheet & "','" & c_ream & "','" & c_sheet & "'," & setupid & ",'" & session & "')"
  
   rs1.MoveNext
   
Wend


'=====================================================

If rs1.State = 1 Then rs1.close
rs1.Open "select  CoverPrinter as Printer,CoverPaper as Paper,CoverSize as PSize,sum(CoverReam),sum(CoverSheet) size from PaperConsumptionPlan WHERE ID='" & txtOrdNo.text & "'  group by CoverPrinter,CoverPaper,CoverSize"
While rs1.EOF = False
   
   sheet = 0
   a1 = rs1(4)
   ream = Int(rs1(3))
   
   sheet = a1 - Int(a1)
   ream = ream + Int(a1)
   
   '-------------------------------------------------------
   c_ream = ream
   c_sheet = sheet
   sheet = 0
   ream = 0
   
   
'   RS.MoveFirst
'   RS.Find "Paper='" & rs1!paper & "'"
'   If RS.EOF = False Then
'      a1_ = RS(3)
'      c_ream = Int(RS(2))
'      c_sheet = a1_ - Int(a1_)
'      c_ream = c_ream + Int(a1_)
'   End If
   
   '-------------------------------------------------------

   
   con.Execute "insert into tmps_LEDGER1(SUBLEDGER,DESCFORINVOICE,address1,address2,address3,phone,owner,setupid,fyear) values('" & rs1!Printer & "','" & rs1!PSize & "','" & rs1!paper & "','" & ream & "','" & sheet & "','" & c_ream & "','" & c_sheet & "'," & setupid & ",'" & session & "')"
  
   rs1.MoveNext
   
Wend





con.Execute "update tmps_LEDGER1 set YEAROPENING = (convert(float,ADDRESS2)+convert(float ,ADDRESS3)+convert(float ,phone)+convert(float ,Owner)) "
con.Execute "update tmps_LEDGER1 set DISCATEGORY = (convert(float,ADDRESS2)+convert(float ,ADDRESS3)),DISTCODE=(convert(float ,phone)+convert(float ,Owner))"


DSNNew


If MsgBox("Want to View", vbYesNo) = vbYes Then

    CR.Reset
    CR.ReportFileName = rptPath & "/Consumption_Printer_PaperwiseNew.rpt"
    CR.Connect = constr
    CR.WindowShowPrintSetupBtn = True
    CR.WindowShowPrintBtn = True
    CR.WindowState = crptMaximized
    CR.Action = 1
    Screen.MousePointer = vbDefault
    
End If


End Sub

Private Sub cmdView_Click()

Dim rss_ As New ADODB.Recordset

con.Execute "delete from temp"
con.Execute "delete from PrinterWisePaper"

If binder_name = "" Then
  Exit Sub
End If

Dim opReam, opSheet


For J = 1 To vs.rows - 1
  
  If vs.TextMatrix(J, 0) <> "" Then
     con.Execute "insert into temp(text,col1,col2) values('" & vs.TextMatrix(J, 16) & "'," & vs.TextMatrix(J, 12) & "," & vs.TextMatrix(J, 13) & ")"
  End If
Next



'Cal_ReamAndSheet

If rss_.State = 1 Then rss_.close
rss_.Open "select op_ream,op_sheet,paper from  PrinterWisePaperOp where Printer='" & binder_name.text & "'", con, adOpenDynamic, adLockOptimistic


If RS.State = 1 Then RS.close
RS.Open "select text,sum(col1),sum(col2)  from temp group by text", con
While RS.EOF = False

Cal_ReamAndSheet RS(1), RS(2)

con.Execute "insert into PrinterWisePaper(Printer,Paper,Con_Ream,Con_Sheet) values('" & binder_name.text & "' ,'" & RS(0) & "'," & ream_tot & "," & sheet_tot & ")"

opReam = ream_tot
opSheet = sheet_tot
sheet_ = 0


rss_.MoveFirst
rss_.Find "paper='" & RS(0) & "'"
If rss_.EOF = False Then
  
  con.Execute "update PrinterWisePaper set OP_Ream=" & rss_(0) & ",OP_sheet=" & rss_(1) & " where paper='" & RS(0) & "' and printer='" & binder_name.text & "'"
    
  sheet_ = (rss_(0) * 500) + rss_(1)
  opSheet = sheet_ - ((opReam * 500) + opSheet)
   
  Cal_ReamAndSheetNew 0, opSheet
   
  ' Cal_ReamAndSheet opReam, opSheet
  con.Execute "update PrinterWisePaper set Bal_Ream=" & ream_tot & ",Bal_Sheet=" & sheet_tot & " where paper='" & RS(0) & "' and printer='" & binder_name.text & "'"
   
End If


RS.MoveNext
Wend


Dim rsfill As New ADODB.Recordset
If rsfill.State = 1 Then rsfill.close
rsfill.Open "select Paper,OP_Ream,OP_Sheet,Con_Ream,Con_Sheet,Bal_Ream,Bal_Sheet  from PrinterWisePaper order by paper", con
Set vs2.DataSource = rsfill






End Sub
Public Function Cal_ReamAndSheetNew(ByVal rm_ As Long, ByVal st_ As Long)
    
Dim cal_ream
Dim D As Integer
Dim recSheet, tmpReam, tmpSheet
recSheet = 0
tmpReam = 0
tmpSheet = 0


recSheet = (rm_ * 500) + st_

If recSheet < 0 Then

recSheet = Abs(recSheet)

tmpReam = Int(recSheet / 500)
tmpSheet = (Round(((recSheet / 500) - Int(recSheet / 500)), 3) * 1000 / 2)
ream_tot = (tmpReam * -1)
sheet_tot = (tmpSheet * -1)

Else

tmpReam = Int(recSheet / 500)
tmpSheet = (Round(((recSheet / 500) - Int(recSheet / 500)), 3) * 1000 / 2)
ream_tot = tmpReam
sheet_tot = tmpSheet


End If


End Function

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'binder_name.Enabled = True
'Me.binder_name.SetFocus
'End If
End Sub

Private Sub Form_Activate()
'Add_Click
End Sub
Sub MaxNo()
'    If RS.State = 1 Then RS.close
'    RS.Open "select max(val(bill_id)) from billmaster", con, adOpenDynamic, adLockOptimistic
'    If Not IsNull(RS.Fields(0).value) Then
'       bill_no.text = RS.Fields(0).value
'    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
sendkeys "{tab}"
End If
End Sub
Sub addpaper_(Optional bk As String, Optional rw As String)

'If bk <> "" Then
'   If rs1_.State = 1 Then rs1_.close
'   rs1_.Open "select Inn_pcode,text_pcode,Exam_pcode,Supp_pcode,Title_pcode from  BookMaster where bookno='" & bk & "'", con
'   If vs.TextMatrix(rw, 3) = "Inner" Then
'      bk = rs1_!Inn_pcode
'   ElseIf vs.TextMatrix(rw, 3) = "Text" Then
'      bk = rs1_!text_pcode
'   ElseIf vs.TextMatrix(rw, 3) = "Exam" Then
'      bk = rs1_!Exam_pcode
'   ElseIf vs.TextMatrix(rw, 3) = "Supp" Then
'      bk = rs1_!supp_pcode
'   ElseIf vs.TextMatrix(rw, 3) = "Title" Then
'      bk = rs1_!title_pcode
'   End If
'End If




s = ""
If rs2_.State = 1 Then rs2_.close
If bk = "" Then
   rs2_.Open "select * from PaperMakeMaster order by papermaker_name", con, adOpenStatic, adLockReadOnly
Else
   rs2_.Open "select * from PaperMakeMaster where papermaker_id='" & bk & "' order by papermaker_name", con, adOpenStatic, adLockReadOnly
End If

While rs2_.EOF = False
If s = "" Then

     s = rs2_!papermaker_name
     If rs2_!eco <> "" Then
       s = s & "-" & rs2_!eco
    End If
    
    If rs2_!SizeValue1 <> "" Then
       s = s & "-" & rs2_!SizeValue1 & "X" & rs2_!SizeValue2
    End If
    
    If rs2_!GSM <> "" Then
       s = s & "-" & rs2_!GSM
    End If
    s = s & "=>" & rs2_!papermaker_id
Else
    If rs2_!papermaker_name <> "" Then
       s = s & "|" & rs2_!papermaker_name
    End If
    If rs2_!eco <> "" Then
       s = s & "-" & rs2_!eco
    End If
    If rs2_!SizeValue1 <> "" Then
       s = s & "-" & rs2_!SizeValue1 & "X" & rs2_!SizeValue2
    End If
    
    If rs2_!GSM <> "" Then
       s = s & "-" & rs2_!GSM
    End If
    
    s = s & "=>" & rs2_!papermaker_id
End If
rs2_.MoveNext
Wend

'If bk = "" Then
'vs.ColComboList(16) = s
'Else
'PopUpValue6 = s
'End If


End Sub

Sub addbinder()

st_ = ""
If rs1.State = 1 Then rs1.close
rs1.Open "select Godwn as [Binder Name],Address from Godownmaster where " & stringyear & " and (Binder_Printer='b' or Binder_Printer='pb') order by Godwn", con, adOpenStatic, adLockReadOnly
While rs1.EOF = False
  If st_ = "" Then
     st_ = rs1(0)
  Else
     st_ = st_ & "|" & rs1(0)
  End If
  rs1.MoveNext
Wend

vs.ColComboList(14) = st_

End Sub
Private Sub Form_Load()
wast_ = 3
Me.Left = 10
Me.top = 10
 
Me.Width = 18000
Me.Height = 11000

grid_ini
bkfont = "e"
fid = "chitra"





maxNo_




txtOrdDate.value = Format(Date, "dd/MM/yyyy")
BackColorFrom Me




'mode = ""
'
'If RS.State = 1 Then RS.close
'RS.Open "select max(convert(int,ord_no)) from Plan_main where " & stringyear, con
'If Not IsNull(RS(0)) Then
'   txtOrdNo.text = RS(0) + 1
'Else
'   txtOrdNo.text = 1
'End If



End Sub


Private Sub Godown_Click()
  If Godown.value = 1 Then
     Label2.Caption = "GoDown Id"
  Else
     Label2.Caption = "Printer Id"
  End If
End Sub



Private Sub txtFirmName_GotFocus()
'' If PopUpValue1 <> "" Then
''    txtFirmName.Text = PopUpValue1
''    PopUpValue1 = ""
'' End If
End Sub

Private Sub txtFirmName_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyCode = 113 Then
''   popuplist1 "select FirmName,Add1,Add2 from FirmMaster order by firmname", con
''End If
If KeyCode = 13 Then binder_name.SetFocus
End Sub

Private Sub ok_Click()




Dim reem, sheet, a1, per
con.Execute "delete from tmps_LEDGER1"

sheet = 0


c_ream = 0
c_sheet = 0


'If RS.State = 1 Then RS.close
'RS.Open "select  Printer,Paper,sum(CoverReam),sum(CoverSheet)  from PaperConsumptionPlan where ID='" & txtOrdNo.text & "' group by Printer,Paper", con



If rs1.State = 1 Then rs1.close
rs1.Open "select  Printer,Paper,PSize,sum(Ream),sum(Sheet) size from PaperConsumptionPlan WHERE ID='" & txtOrdNo.text & "'  group by Printer,Paper,PSize"
While rs1.EOF = False
   
   sheet = 0
   a1 = rs1(4)
   ream = Int(rs1(3))
   
   sheet = a1 - Int(a1)
   ream = ream + Int(a1)
   
   '-------------------------------------------------------
   c_ream = 0
   c_sheet = 0
   
   
'   RS.MoveFirst
'   RS.Find "Paper='" & rs1!paper & "'"
'   If RS.EOF = False Then
'      a1_ = RS(3)
'      c_ream = Int(RS(2))
'      c_sheet = a1_ - Int(a1_)
'      c_ream = c_ream + Int(a1_)
'   End If
   
   '-------------------------------------------------------

   
   con.Execute "insert into tmps_LEDGER1(SUBLEDGER,DESCFORINVOICE,address1,address2,address3,phone,owner,setupid,fyear) values('" & rs1!Printer & "','" & rs1!PSize & "','" & rs1!paper & "','" & ream & "','" & sheet & "','" & c_ream & "','" & c_sheet & "'," & setupid & ",'" & session & "')"
  
   rs1.MoveNext
   
Wend


'=====================================================

If rs1.State = 1 Then rs1.close
rs1.Open "select  CoverPrinter as Printer,CoverPaper as Paper,CoverSize as PSize,sum(CoverReam),sum(CoverSheet) size from PaperConsumptionPlan WHERE ID='" & txtOrdNo.text & "'  group by CoverPrinter,CoverPaper,CoverSize"
While rs1.EOF = False
   
   sheet = 0
   a1 = rs1(4)
   ream = Int(rs1(3))
   
   sheet = a1 - Int(a1)
   ream = ream + Int(a1)
   
   '-------------------------------------------------------
   c_ream = ream
   c_sheet = sheet
   sheet = 0
   ream = 0
   
   
'   RS.MoveFirst
'   RS.Find "Paper='" & rs1!paper & "'"
'   If RS.EOF = False Then
'      a1_ = RS(3)
'      c_ream = Int(RS(2))
'      c_sheet = a1_ - Int(a1_)
'      c_ream = c_ream + Int(a1_)
'   End If
   
   '-------------------------------------------------------

   
   con.Execute "insert into tmps_LEDGER1(SUBLEDGER,DESCFORINVOICE,address1,address2,address3,phone,owner,setupid,fyear) values('" & rs1!Printer & "','" & rs1!PSize & "','" & rs1!paper & "','" & ream & "','" & sheet & "','" & c_ream & "','" & c_sheet & "'," & setupid & ",'" & session & "')"
  
   rs1.MoveNext
   
Wend





con.Execute "update tmps_LEDGER1 set YEAROPENING = (convert(float,ADDRESS2)+convert(float ,ADDRESS3)+convert(float ,phone)+convert(float ,Owner)) "
con.Execute "update tmps_LEDGER1 set DISCATEGORY = (convert(float,ADDRESS2)+convert(float ,ADDRESS3)),DISTCODE=(convert(float ,phone)+convert(float ,Owner))"


DSNNew


If MsgBox("Want to View", vbYesNo) = vbYes Then

    CR.Reset
    CR.ReportFileName = rptPath & "/Consumption_Printer_Paperwise.rpt"
    CR.Connect = constr
    CR.WindowShowPrintSetupBtn = True
    CR.WindowShowPrintBtn = True
    CR.WindowState = crptMaximized
    CR.Action = 1
    Screen.MousePointer = vbDefault
    
End If



End Sub

Private Sub txtOrdDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then txtFirmName.SetFocus
End Sub

Private Sub txtOrdNo_GotFocus()
 If PopUpValue1 <> "" Then
    txtOrdNo.text = PopUpValue1
    txtOrdDate.value = PopUpValue2
    txtrem.text = PopUpValue3
    
    searchData
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
 End If
End Sub
Private Sub txtOrdNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   popuplist1 "Select distinct Id,Dates,FileDesc from PaperConsumptionPlan order by id", con
End If

If KeyCode = 13 Then txtOrdDate.SetFocus

End Sub

Sub addTotalReam()

Dim ream_ As Long
Dim sheet_ As Long
Dim per_

ream_ = 0
sheet_ = 0
per_ = 0

For I = 1 To vs.rows - 1
  If vs.TextMatrix(I, 0) <> "" Then
     ream_ = ream_ + Val(vs.TextMatrix(I, 12))
     sheet_ = sheet_ + Val(vs.TextMatrix(I, 13))
  End If
Next



If sheet_ > 499 Then
   per_ = Int(sheet_ / 500)
   sheet_ = sheet_ - per_ * 500
End If


txtTReam = ream_ + per_
txtTSheet = sheet_


End Sub
Function cheqePrinter(P1 As String, I As Integer, bno As String) As Boolean
   
   Dim ss1 As New ADODB.Recordset
   
   Select Case I
   
   Case 1
   
     If ss1.State = 1 Then ss1.close
     ss1.Open "select * from bookmaster where " & stringyear & " and Inn_Printer='" & binder_name.text & "' and bookno='" & bno & "'", con, adOpenDynamic, adLockReadOnly
     If ss1.EOF = False Then
        cheqePrinter = True
     Else
        cheqePrinter = False
     End If
     
   Case 2
   
     
     If ss1.State = 1 Then ss1.close
     ss1.Open "select * from bookmaster where " & stringyear & " and text_Printer='" & binder_name.text & "' and bookno='" & bno & "'", con, adOpenDynamic, adLockReadOnly
     If ss1.EOF = False Then
        cheqePrinter = True
     Else
        cheqePrinter = False
     End If
     
   
   Case 3
   
   
     If ss1.State = 1 Then ss1.close
     ss1.Open "select * from bookmaster where " & stringyear & " and Exam_Printer='" & binder_name.text & "' and bookno='" & bno & "'", con, adOpenDynamic, adLockReadOnly
     If ss1.EOF = False Then
        cheqePrinter = True
     Else
        cheqePrinter = False
     End If
   
   
   
   Case 4
   
     
     If ss1.State = 1 Then ss1.close
     ss1.Open "select * from bookmaster where " & stringyear & " and Supp_Printer='" & binder_name.text & "' and bookno='" & bno & "'", con, adOpenDynamic, adLockReadOnly
     If ss1.EOF = False Then
        cheqePrinter = True
     Else
        cheqePrinter = False
     End If
     
   
   Case 5
   
     
     If ss1.State = 1 Then ss1.close
     ss1.Open "select * from bookmaster where " & stringyear & " and Title_Printer='" & binder_name.text & "' and bookno='" & bno & "'", con, adOpenDynamic, adLockReadOnly
     If ss1.EOF = False Then
        cheqePrinter = True
     Else
        cheqePrinter = False
     End If
   
   
   
   End Select

End Function

Private Sub txtOrdNo_LostFocus()
If txtOrdNo <> "" Then
con.Execute "delete from tmporder where orderno='" & txtOrdNo & "'"
End If

End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next

If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      'If vs.Row > 1 Then
         vs.RemoveItem vs.Row
         addTotalReam
         PopUpValue6 = "y"
         vs.SetFocus
      'Else
      '   MsgBox "You cannnot delete first and fixed row you can edit it", vbInformation
      'End If
   End If
End If


End Sub
Sub Ream_SheetNew(r As Integer)

Dim Tot, wastage_per, Form, wream


wastage_per = 0

Form = Val(vs.TextMatrix(r, 10))
quan = Val(vs.TextMatrix(r, 2))
wastage_per = wast_   ' Val(vs.TextMatrix(r, 11))

Tot = Form * quan
Tot = Tot / 1000

If Val(wastage_per) > 0 Then
   a1 = (Tot * wastage_per / 100)
   wream = Int(a1)
   wsheet = Round((a1 - Int(a1)) * 500)
   
End If

  
If Tot > 0 Then
   
   ream = Int(Tot)
   sheet = Round((Tot - Int(Tot)) * 500, 0)
   sheet = sheet + wsheet
   
   If sheet > 499 Then
      wream = Int(sheet / 500)
      sheet = sheet - (wream * 500)

      ream = ream + wream
   Else
      ream = ream + wream
   End If
 
   
  
End If




End Sub

Sub calc()
''prow = vs.Row
'''grid_ini
''Dim unit, unit1 As Double
''Dim rate, atrate As Double
''Dim amt As Double
''Dim Tot, quan, sheet, totrim As Double
''Dim wrim, wper, wsheet, wtemp As Double
''Dim lent As Integer
''tmpplt = 0
''vs.Col = 5: quan = Val(vs.Text)
''vs.Col = 11: per = Val(vs.Text)
''vs.Col = 6: atrate = Val(vs.Text)
'''manual calculations start
'''If Checkmanual.value <> 1 Then
'''  tmp = quan / 1100
'''  If tmp < 1 Then tmp = 1
'''  If Int(tmp) = tmp Then rate = atrate * tmp
'''    If Int(tmp) < tmp Then
'''        If (tmp - Int(tmp)) <= 0.5 Then rate = atrate * (Int(tmp) + 0.5)
'''        If (tmp - Int(tmp)) > 0.5 Then rate = atrate * (Int(tmp) + 1)
'''    End If
'''   vs.Col = 7: vs.Text = rate
'''    tmpplt = quan / 11000
'''    If Int(tmpplt) < tmpplt Then
'''        tmpplt = Int(tmpplt) + 1
'''   End If
'''   vs.Col = 14: vs.Text = tmpplt
'''End If
''  vs.Col = 3: If Val(vs.Text) >= 1 Then unit1 = Val(vs.Text) Else unit1 = 1
''  unit = Val(vs.Text)
''  vs.Col = 7: rate = Val(vs.Text)
''
'''If Checkmanual.value <> 1 Then
'''tmpplate1 = 0
'''If Int(unit) < unit Then
'''X = unit - Int(unit)
'''If X <= 0.5 Then tmpplate1 = Int(unit) + 1
'''If X > 0.5 Then tmpplate1 = Int(unit) + 2
'''Else
'''tmpplate1 = unit
'''End If
'''vs.Col = 14: tmpplt = vs.Text * tmpplate1
'''vs.Col = 14: vs.Text = tmpplt
'''Else
'''vs.Col = 14:  tmpplt = vs.Text
'''End If
''  If quan > 0 And quan <= 1100 Then
''  amt = Round(Val(tmpplt) * rate, 0)
''  Else
''  amt = Round(unit1 * rate, 0)
''  End If
''  vs.Col = 8: vs.Text = amt
''vs.Col = 15
''pltamt = Val(tmpplt) * Val(vs.Text)
''vs.Col = 16: vs.Text = pltamt
''
''  wastflag = "K"
''  tmpwast = 0
''
''  If per <> 15 Then
''  tmpwast = per * quan / 100
''  Else
''  tmpwast = 15
''  End If
''  tmpact = 0
''  tmpact = per * 1000 / 100
''  If tmpwast > 0 And tmpwast <= tmpact Then
''  wastflag = "N"
''  wast = 15
''  vs.Col = 11: vs.Text = 15
''  End If
''  per = 0
''
''  Tot = unit * quan
''  If Tot > 0 Then
''    Tot = Tot / 1000
''    sheet = (Tot - Int(Tot)) * 1000 / 2 'ghgfhfh
''
''    vs.Col = 11: wrim = Val(vs.Text)
''    If wastflag = "N" Then
''    tmptotw1 = 0
'' tmptotw1 = tmpplt * wast
''
''    wsheet = Round(tmptotw1, 0)
''   tmpwreams = 0
''tmpwsheets = 0
''If wsheet > 499 Then
''per = Int(wsheet / 500)
''wsheet = wsheet - per * 500
''End If
''    Else
''    per = Tot * wrim / 100
''    wsheet = (per - Int(per)) * 1000 / 2
''    End If
''  Else
''    If Tot > 500 Then
''        tmptot = Tot
''        sheet = Tot - 500
''        Tot = 1
''        vs.Col = 11: wrim = Val(vs.Text)
''         If wastflag = "N" Then
''    per = unit * wast
''    Else
''        per = tmptot * wrim / 100
''        End If
''        If per > 500 Then
''            wsheet = per - 1
''            per = 1
''        Else
''        wsheet = Round(per, 0)
''        per = 0
''        End If
''    Else
''    sheet = Tot
''     vs.Col = 11: wrim = Val(vs.Text)
''      If wastflag = "N" Then
''    per = unit * wast
''    Else
''        per = Tot * wrim / 100
''        End If
''        Tot = 0
''        wsheet = per
''        per = 0
''    End If
''  End If
''  vs.Col = 9
''  vs.Text = Int(Tot)
''  vs.Col = 10
''  vs.Text = Round(sheet, 0)
''
''vs.Col = 12: vs.Text = Int(per)
''vs.Col = 13: vs.Text = Round(wsheet)
''Dim amount, reams, sheets, plate
''Dim ptamt As Double
''reams = 0
''sheets = 0
''amount = 0
''wreams = 0
''wsheets = 0
''plate = 0
''ptamt = 0
''
''vs.Col = 8
''For i = 1 To vs.Rows - 1
''vs.Row = i
''vs.Col = 8
''amount = amount + Val(vs.Text)
''vs.Col = 9
''reams = Round(reams + Val(vs.Text), 0)
''vs.Col = 10
''sheets = Round(sheets + Val(vs.Text), 0)
''vs.Col = 12
''wreams = Round(wreams + Val(vs.Text), 0)
''vs.Col = 13
''wsheets = Round(wsheets + Val(vs.Text), 0)
''vs.Col = 14
''plate = plate + Val(vs.Text)
''vs.Col = 16
''ptamt = ptamt + Val(vs.Text)
''Next i
''Me.total = Round(amount, 0)
''
'''''Me.totalplate = Round(plate, 0)
''
''Me.pltamt = Round(ptamt, 0)
''TMPreams = 0
''TMPsheets = 0
''If sheets > 499 Then
''TMPreams = Int(sheets / 500)
''TMPsheets = TMPreams * 500
''End If
''Me.totalream = Round(reams, 0) + Round(TMPreams, 0)
''Me.totalsht = Round(sheets, 0) - Round(TMPsheets, 0)
''tmpwreams = 0
''tmpwsheets = 0
''If wsheets > 499 Then
''tmpwreams = Int(wsheets / 500)
''tmpwsheets = tmpwreams * 500
''End If
''
''Me.totalwream = Round(wreams, 0) + Round(tmpwreams, 0)
''Me.totalwsht = Round(wsheets, 0) - Round(tmpwsheets, 0)
''
''tmpgreams = 0
''tmpgsheets = 0
''greams = 0
''gsheets = 0
''greams = Round(Val(Me.totalream), 0) + Round(Val(Me.totalwream), 0)
''gsheets = Round(Val(Me.totalsht), 0) + Round(Val(Me.totalwsht), 0)
''If gsheets > 499 Then
''tmpgreams = Int(gsheets / 500)
''tmpgsheets = tmpgreams * 500
''End If
''Me.gtotalreams = greams + tmpgreams
''Me.gtotalsheets = gsheets - tmpgsheets
''Me.gtotalamt = Round(Val(total) + Val(pltamt), 0)
''vs.Row = prow
''
'''If vs.TextMatrix(vs.RowSel, 0) = "" Then
'''   vs.Col = 1
'''End If
''
'''Me.total = Val(total) + amt
''
End Sub
Private Sub vs_KeyPress(KeyAscii As Integer)




If vs.Col <> 1 Then
'Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
'Label14.Visible = False
Label16.Visible = False
End If
If vs.Col = 1 Then
'Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
'Label14.Visible = True
Label16.Visible = True
End If



If vs.Col = 1 Or vs.Col = 20 Then
        If KeyAscii = 8 Then
        If Len(Trim(vs.text)) <> 0 Then
                vs.text = Left(vs.text, (Len(vs.text) - 1))
        End If
            'ElseIf (KeyAscii = 13) Or (KeyAscii >= 32 And KeyAscii <= 126) Then
            ElseIf (KeyAscii >= 32 And KeyAscii <= 126) Then
            vs.text = vs.text + Chr(KeyAscii)
        End If
End If


If vs.Col = 2 Then

        If Len(Trim(vs.text)) = 80 Then
    KeyAscii = 0
    End If
        If KeyAscii = 8 Then
            If Len(Trim(vs.text)) <> 0 Then
            vs.text = Left(vs.text, (Len(vs.text) - 1))
        End If
            ElseIf (KeyAscii <> 13) Or (KeyAscii >= 32 And KeyAscii <= 126) Then
            vs.text = vs.text + Chr(KeyAscii)
        End If


End If

If vs.Col = 4 Then
    If KeyAscii = 119 Or KeyAscii = 87 Then
       vs.text = "Web"
       web.Visible = False
    End If
    If KeyAscii = 115 Or KeyAscii = 83 Then
       vs.text = "Sheet"
       web.Visible = False
    End If

End If

If (vs.Col = 3 Or vs.Col = 4 Or vs.Col = 5 Or vs.Col = 6 Or vs.Col = 7 Or vs.Col = 11 Or vs.Col = 12 Or vs.Col = 13 Or vs.Col = 14 Or vs.Col = 15) Then
    If Len(Trim(vs.text)) = 10 Then
    KeyAscii = 0
    End If
    If KeyAscii = 8 Then
        If Len(Trim(vs.text)) <> 0 Then
                vs.text = Left(vs.text, (Len(vs.text) - 1))
        End If
    End If
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 Then
                KeyAscii = 0
            End If
        End If
    End If
    If KeyAscii <> 8 Then
        vs.text = vs.text + Chr(KeyAscii)
    End If
End If



End Sub
Function addPage() As Double
   

page_sum = 0
   
For P1 = 4 To 4
   If Val(vs.TextMatrix(vs.RowSel, P1)) > 0 Then
      page_sum = page_sum + Val(vs.TextMatrix(vs.RowSel, P1))
   End If
Next

addPage = page_sum

End Function
Sub addType(bk As String)
s11 = ""

Dim k1 As Integer
k1 = 6

If rs1.State = 1 Then rs1.close
rs1.Open "select Head1,Head2,Head3,Head4,txtHead6,txtHead7 from BookMaster where bookno='" & bk & "'", con
If rs1.EOF = False Then
For Q1 = 1 To 6

If Q1 >= 5 Then

    If Not IsNull(rs1.Fields("txthead" & k1).value) Then
    If rs1.Fields("txthead" & k1).value <> "" Then
        If s11 = "" Then
           s11 = rs1.Fields("txthead" & k1).value
        Else
           s11 = s11 & "|" & rs1.Fields("txthead" & k1).value
        End If
    End If
    End If
    k1 = k1 + 1
Else

    If Not IsNull(rs1.Fields("head" & Q1).value) Then
    If rs1.Fields("head" & Q1).value <> "" Then
        If s11 = "" Then
           s11 = rs1.Fields("head" & Q1).value
        Else
           s11 = s11 & "|" & rs1.Fields("head" & Q1).value
        End If
    End If
    End If


End If



Next

End If

vs.ColComboList(3) = s11

If s11 = "" Then

If rs1.State = 1 Then rs1.close
rs1.Open "select * from MasterTbl where category='bkpart'", con
While rs1.EOF = False
    If s11 = "" Then
       s11 = rs1.Fields("name").value
    Else
       s11 = s11 & "|" & rs1.Fields("name").value
    End If
 rs1.MoveNext
Wend

vs.ColComboList(3) = s11

End If



  
   
End Sub
Sub fatchRawsData(r1 As Integer, r2 As Integer)
    
    Dim k1 As Integer
    
    k1 = r1
    
    For m1 = 0 To 2
           vs.TextMatrix(r2, m1) = vs.TextMatrix(r1, m1)
    Next
    
    
    For m1 = 1 To 3
          ' vs.TextMatrix(r2, m1) = vs.TextMatrix(r1, m1)
           
           
         If vs.TextMatrix(m1, 0) <> "" Then
            Ream_SheetNew (k1)
            vs.TextMatrix(k1, 12) = ream
            vs.TextMatrix(k1, 13) = sheet
            k1 = k1 + 1
         End If
           
    Next
    
    
    vs.TextMatrix(r2, 5) = vs.TextMatrix(r1, 5)
    vs.TextMatrix(r2, 9) = vs.TextMatrix(r1, 9)
    vs.TextMatrix(r2, 14) = vs.TextMatrix(r1, 14)
    vs.TextMatrix(r2, 18) = vs.TextMatrix(r1, 18)
    vs.TextMatrix(r2, 19) = vs.TextMatrix(r1, 19)
    vs.TextMatrix(r2, 20) = vs.TextMatrix(r1, 20)
    vs.TextMatrix(r2, 11) = vs.TextMatrix(r1, 11)
    
    
    
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)

'On Error Resume Next
'
'If KeyCode = 13 Then
'
'If vs.Col = 0 Then
'
'   If PopUpValue6 = "y" Then
'      PopUpValue6 = ""
'      Exit Sub
'   End If
'
'   If RS.State = 1 Then RS.close
'   RS.Open "select book,book_unit,DivideValue,HeadData1,HeadData2,HeadData3,HeadData4,HeadData5,bookfont,Price,trimsize,binder,binding from BookMaster " & _
'   " where " & stringyear & " and BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
'   If RS.EOF = False Then
'
'      addbinder
'
'
'
'      vs.TextMatrix(vs.RowSel, 0) = UCase(vs.TextMatrix(vs.RowSel, 0))
'      vs.TextMatrix(vs.RowSel, 1) = UCase(RS!Book)
'
'      qty = InputBox("Enter Qty ")
'      vs.TextMatrix(vs.RowSel, 2) = qty
'
'
'
'      vs.TextMatrix(vs.RowSel, 18) = RS!trimsize & ""
'      vs.TextMatrix(vs.RowSel, 19) = "No"
'
'      vs.TextMatrix(vs.RowSel, 9) = Val(RS!DivideValue)
'
'      vs.TextMatrix(vs.RowSel, 5) = RS!Price & ""
'      vs.TextMatrix(vs.RowSel, 11) = wast_
'
'      vs.TextMatrix(vs.RowSel, 14) = RS!Binder & ""
'      vs.TextMatrix(vs.RowSel, 20) = RS!Binding & ""
'
'
'      If vs.TextMatrix(vs.RowSel, 5) = "" Then
'        If RS.State = 1 Then RS.close
'        RS.Open "select RATE from books where BOOKCODE='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
'        If RS.EOF = False Then
'           vs.TextMatrix(vs.RowSel, 5) = RS(0)
'        End If
'      End If
'      addType vs.TextMatrix(vs.RowSel, 0)
'
'
'
'
'
'
'      'Add All Daya================================================================
'
'
'       If rs1.State = 1 Then rs1.close
'       rs1.Open "select * from tmporder where (orderNo='" & txtOrdNo & "' and Bkcode='" & vs.TextMatrix(vs.RowSel, 0) & "')", con, adOpenDynamic, adLockOptimistic
'       If rs1.EOF = True Then
'          rs1.AddNew
'          rs1!orderNo = Trim(txtOrdNo)
'          rs1!bkcode = Trim(vs.TextMatrix(vs.RowSel, 0))
'          rs1.update
'       Else
'          GoTo dinesh
'       End If
'
'
'       bo_ = False
'       l1 = 1
'
'       If rs1.State = 1 Then rs1.close
'       rs1.Open "select Head1,Head2,Head3,Head4,txthead6 as Head5,HeadData1,HeadData2,HeadData3,HeadData4,txtheadData6 as HeadData5,pcode1,pcode2,pcode3,pcode4,pcode5,pcode6,pcode7,color1,color2,color3,color4,color5,color6,color7 from BookMaster " & _
'       " where " & stringyear & " and BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
'       If rs1.EOF = False Then
'        For b1 = 1 To 5
'          PopUpValue6 = ""
'          If Not IsNull(rs1.Fields("Head" & b1)) Then
'
'          If bo_ = False Then
'
'              If b1 >= 6 Then
'               vs.TextMatrix(vs.RowSel, 3) = rs1.Fields("txtHead" & b1)
'               vs.TextMatrix(vs.RowSel, 4) = rs1.Fields("txtHeadData" & b1)
'              Else
'               vs.TextMatrix(vs.RowSel, 3) = rs1.Fields("Head" & b1)
'               vs.TextMatrix(vs.RowSel, 4) = rs1.Fields("HeadData" & b1)
'              End If
'
'               addpaper_ rs1.Fields("pcode" & b1), vs.RowSel
'               vs.TextMatrix(vs.RowSel, 16) = PopUpValue6
'               vs.TextMatrix(vs.RowSel, 7) = rs1.Fields("color" & b1)
'               bo_ = True
'
'               vs.TextMatrix(vs.RowSel, 10) = Round(Val(vs.TextMatrix(vs.RowSel, 4)) / Val(vs.TextMatrix(vs.RowSel, 9)), 2)
'
''               Ream_SheetNew (vs.RowSel)
''               vs.TextMatrix(vs.RowSel, 12) = ream
''               vs.TextMatrix(vs.RowSel, 13) = sheet
'
'
'          Else
'
'           If rs1.Fields("Head" & b1) <> "" Then
'
'
'
'              If b1 >= 6 Then
'               vs.TextMatrix(vs.RowSel + l1, 3) = rs1.Fields("txtHead" & b1)
'               vs.TextMatrix(vs.RowSel + l1, 4) = rs1.Fields("txtHeadData" & b1)
'              Else
'               vs.TextMatrix(vs.RowSel + l1, 3) = rs1.Fields("Head" & b1)
'               vs.TextMatrix(vs.RowSel + l1, 4) = rs1.Fields("HeadData" & b1)
'              End If
'
'
'
'               vs.TextMatrix(vs.RowSel + l1, 10) = Round(Val(vs.TextMatrix(vs.RowSel + l1, 4)) / Val(vs.TextMatrix(vs.RowSel, 9)), 2)
'
''               Ream_SheetNew (vs.RowSel + l1)
''               vs.TextMatrix(vs.RowSel + l1, 12) = ream
''               vs.TextMatrix(vs.RowSel + l1, 13) = sheet
'
'
'
'               fatchRawsData vs.RowSel, vs.RowSel + l1
'               addpaper_ rs1.Fields("pcode" & b1), vs.RowSel + l1
'               vs.TextMatrix(vs.RowSel + l1, 16) = PopUpValue6
'               vs.TextMatrix(vs.RowSel + l1, 7) = rs1.Fields("color" & b1)
'
'
'               l1 = l1 + 1
'
'           End If
'
'          End If
'
'          End If
'
'
'
'
'        Next
'       End If
'       PopUpValue6 = ""
'      '============================================================================
'dinesh:
'
'
'
'
'
'      sendkeys "{down}"
'      sendkeys "{down}"
'      'SendKeys "{right}"
'
'   End If
'
'ElseIf vs.Col = 2 Then
'      sendkeys "{right}"
'ElseIf vs.Col = 3 Then
'
'
'
'     If rs1.State = 1 Then rs1.close
'     rs1.Open "select Head1,Head2,Head3,Head4,Head5,HeadData1,HeadData2,HeadData3,HeadData4,HeadData5 from BookMaster " & _
'     " where " & stringyear & " and BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
'     If rs1.EOF = False Then
'        For b1 = 1 To 5
'            If vs.TextMatrix(vs.RowSel, 3) = rs1.Fields("Head" & b1) Then
'               vs.TextMatrix(vs.RowSel, 4) = rs1.Fields("HeadData" & b1)
'               bo_ = True
'            End If
'        Next
'
'
'
'     End If
'
'     sendkeys "{right}"
'     vs.TextMatrix(vs.RowSel, 8) = addPage
'
'ElseIf vs.Col = 4 Then
'
'
'        If bo_ = False Then
'        If MsgBox("Want to add in master", vbQuestion + vbYesNo) = vbYes Then
'            For b1 = 1 To 5
'                If (IsNull(rs1.Fields("Head" & b1)) Or rs1.Fields("Head" & b1) = "") Then
'                   con.Execute "update BookMaster set " & rs1.Fields("Head" & b1).Name & "='" & vs.TextMatrix(vs.RowSel, 3) & "'," & rs1.Fields("Headdata" & b1).Name & "='" & vs.TextMatrix(vs.RowSel, 4) & "'" & " where BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'"
'                   addType vs.TextMatrix(vs.RowSel, 0)
'                   GoTo aaaa:
'                End If
'            Next
'        End If
'        End If
'
'aaaa:
'      vs.SetFocus
'
'      sendkeys "{right}"
'      vs.TextMatrix(vs.RowSel, 8) = addPage
'ElseIf vs.Col = 5 Then
'
'
'
'  If rs1.State = 1 Then rs1.close
'  rs1.Open "select PRICE from BookMaster " & _
'  " where BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
'  If rs1.EOF = False Then
'
'  If Val(vs.TextMatrix(vs.RowSel, 5)) <> rs1(0) Then
'    If MsgBox("Want to add in master", vbQuestion + vbYesNo) = vbYes Then
'       con.Execute "update BookMaster set price='" & vs.TextMatrix(vs.RowSel, 5) & "' where BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'"
'       vs.SetFocus
'    End If
'  End If
'  End If
'
'
'
'   If vs.TextMatrix(vs.RowSel, 5) <> "" Then
'      sendkeys "{right}"
'   End If
'
'   vs.TextMatrix(vs.RowSel, 8) = addPage
'
'
'ElseIf vs.Col = 7 Then
'
'      sendkeys "{right}"
'      vs.TextMatrix(vs.RowSel, 8) = addPage
'
'ElseIf vs.Col = 9 Then
'      If RS.State = 1 Then RS.close
'      RS.Open "select DivideValue from BookMaster where bookno='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
'      If RS.EOF = False Then
'        If (Val(RS!DivideValue) <> Val(vs.TextMatrix(vs.RowSel, 9))) Then
'          If MsgBox("Want to add in master", vbQuestion + vbYesNo) = vbYes Then
'             con.Execute "update BookMaster set DivideValue='" & vs.TextMatrix(vs.RowSel, 9) & "' where BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'"
'          End If
'        End If
'      End If
'
'      vs.SetFocus
'
'      sendkeys "{right}"
'
'      'Ream_Sheet
'
'      vs.TextMatrix(vs.RowSel, 8) = addPage
'
'
'ElseIf vs.Col = 10 Then
'      sendkeys "{right}"
'ElseIf vs.Col = 11 Then
'      sendkeys "{right}"
'      Ream_Sheet
'      vs.TextMatrix(vs.RowSel, 12) = ream
'      vs.TextMatrix(vs.RowSel, 13) = sheet
'      sendkeys "{right}"
'      sendkeys "{right}"
'ElseIf vs.Col = 14 Then
'
'    If vs.TextMatrix(vs.RowSel, 14) <> "" Then
'       'SendKeys "{home}"
'       'SendKeys "{down}"
'       sendkeys "{right}"
'       addTotalReam
'    End If
'
'ElseIf vs.Col = 16 Then
'
'    'If vs.TextMatrix(vs.RowSel, 15) <> "" Then
'    '  SendKeys "{home}"
'    '  SendKeys "{down}"
'      sendkeys "{right}"
'      addTotalReam
'    'End If
'
'ElseIf vs.Col = 17 Then
'
'      sendkeys "{right}"
'      addTotalReam
'
'ElseIf vs.Col = 18 Then
'
'
'''       If rs1.State = 1 Then rs1.close
'''       rs1.Open "select trimsize from BookMaster " & _
'''        " where BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
'''       If rs1.EOF = False Then
'''
'''        If rs1(0) = "" Or IsNull(rs1(0)) Then
'''          If MsgBox("Want to add in master", vbQuestion + vbYesNo) = vbYes Then
'''             con.Execute "update BookMaster set trimsize='" & vs.TextMatrix(vs.RowSel, 18) & "' where BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'"
'''             vs.SetFocus
'''          End If
'''       End If
'''       End If
'
'
'      sendkeys "{right}"
'      addTotalReam
'
'ElseIf vs.Col = 19 Then
'
'      sendkeys "{right}"
'      addTotalReam
'
'
'ElseIf vs.Col = 20 Then
'
'      sendkeys "{home}"
'      sendkeys "{down}"
'      addTotalReam
'
'
'End If
'End If
'
'



End Sub

Private Sub vs_SelChange()

If (vs.Col = 0 Or vs.Col = 2 Or vs.Col = 3 Or vs.Col = 4 Or vs.Col = 5 Or vs.Col = 6 Or vs.Col = 7 Or vs.Col = 9 Or vs.Col = 11 Or vs.Col = 14 Or vs.Col = 16) Then
   vs.Editable = flexEDKbdMouse
   
ElseIf vs.Col = 18 Then
   vs.Editable = flexEDKbdMouse
   
ElseIf vs.Col = 19 Then
   vs.Editable = flexEDKbdMouse
ElseIf vs.Col = 20 Then
   vs.Editable = flexEDKbdMouse
ElseIf vs.Col = 21 Then
   vs.Editable = flexEDKbdMouse
ElseIf vs.Col = 17 Then
   vs.Editable = flexEDKbdMouse
Else
   vs.Editable = flexEDNone
End If




End Sub

Private Sub ItemCode_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  sendkeys "{down}"
  PopUpValue1 = Trim(Mid(ItemCode.text, InStr(ItemCode.text, "=>") + 2))
  Call searchData
  
ElseIf KeyAscii = 27 Then
  ItemCode.Visible = False
End If


End Sub
Sub searchData()

Dim vsFill As New ADODB.Recordset

Set vsFill = New ADODB.Recordset

vs.Clear
grid_ini

d1 = 1

vsFill.Open "select SNo,Books,Class,Pages,Forms,BSize,Vender,Remarks,PrintRun,Price,Paper," & _
"PSize,Colour,Printer,Consumption,FileDesc,CoverPrinter,CoverSize,CoverPaper,CoverConsumption" & _
" from PaperConsumptionPlan where Id='" & txtOrdNo.text & "' order by Sno", con
While vsFill.EOF = False

vs.TextMatrix(d1, 0) = vsFill!sno
vs.TextMatrix(d1, 1) = vsFill!Books
vs.TextMatrix(d1, 2) = vsFill!Class
vs.TextMatrix(d1, 3) = vsFill!Pages
vs.TextMatrix(d1, 4) = vsFill!Forms
vs.TextMatrix(d1, 5) = vsFill!bsize
'vs.TextMatrix(d1, 6) = vsFill!Vender
'vs.TextMatrix(d1, 7) = vsFill!remarks
vs.TextMatrix(d1, 8) = vsFill!PrintRun
vs.TextMatrix(d1, 9) = vsFill!Price

vs.TextMatrix(d1, 10) = vsFill!paper
vs.TextMatrix(d1, 11) = vsFill!PSize
'vs.TextMatrix(d1, 12) = vsFill!Colour
vs.TextMatrix(d1, 13) = vsFill!Printer
vs.TextMatrix(d1, 14) = vsFill!Consumption

vs.TextMatrix(d1, 15) = vsFill!CoverPrinter & ""
vs.TextMatrix(d1, 16) = vsFill!Coversize & ""
vs.TextMatrix(d1, 17) = vsFill!CoverPaper & ""
vs.TextMatrix(d1, 18) = vsFill!CoverConsumption & ""

vs.rows = vs.rows + 1
vsFill.MoveNext
d1 = d1 + 1

Wend


End Sub


Private Sub order_no_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   TextPaperSize.SetFocus
End If
End Sub

Private Sub order_no_LostFocus()
 'Binder_id.SetFocus
End Sub

Private Sub party_id_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   popuplist1 "Select customer_id as [Printer Id],customer_name as [Printer Name],city from customerMaster where " & stringyear & " and Category='Printer'", con
ElseIf KeyCode = 13 Then
   TextPaperSize.SetFocus
End If
End Sub
Private Sub party_id_LostFocus()
Label10.Visible = False
End Sub
Private Sub Printcmd_Click()

c_ream = 0
c_sheet = 0


'If RS.State = 1 Then RS.close
'RS.Open "select  Paper,PSize,sum(CoverReam),sum(CoverSheet) size from PaperConsumptionPlan where ID='" & txtOrdNo.text & "' group by paper,PSize"


Dim reem, sheet, a1, per
con.Execute "delete from tmps_LEDGER1"

sheet = 0

If rs1.State = 1 Then rs1.close
rs1.Open "select  Paper,PSize,sum(Ream),sum(Sheet)  from PaperConsumptionPlan where ID='" & txtOrdNo.text & "' group by paper,PSize"
While rs1.EOF = False
   
   
   sheet = 0
   a1 = rs1(3)
   ream = Int(rs1(2))
   
   sheet = a1 - Int(a1)
   ream = ream + Int(a1)
   
   '-------------------------------------------------------
   c_ream = 0
   c_sheet = 0
   
   'RS.MoveFirst
   'RS.Find "Paper='" & rs1!paper & "'"
   'If RS.EOF = False Then
   '   a1_ = RS(3)
   '   c_ream = Int(RS(2))
   '   c_sheet = a1_ - Int(a1_)
   '   c_ream = c_ream + Int(a1_)
   'End If
   
   '-------------------------------------------------------
   
   
   con.Execute "insert into tmps_LEDGER1(address1,address2,address3,phone,owner,setupid,fyear,DESCFORINVOICE) values('" & rs1!paper & "','" & ream & "','" & sheet & "','" & c_ream & "','" & c_sheet & "'," & setupid & ",'" & session & "','" & rs1!PSize & "')"
  
   rs1.MoveNext
   
Wend

'==========================================================


sheet = 0

If rs1.State = 1 Then rs1.close
rs1.Open "select  CoverPaper,CoverSize,sum(CoverReam),sum(CoverSheet)  from PaperConsumptionPlan where ID='" & txtOrdNo.text & "' group by CoverPaper,CoverSize"
While rs1.EOF = False
   
   
   sheet = 0
   a1 = rs1(3)
   ream = Int(rs1(2))
   
   sheet = a1 - Int(a1)
   ream = ream + Int(a1)
   
   '-------------------------------------------------------
   c_ream = ream
   c_sheet = sheet
   
   sheet = 0
   ream = 0
   
   'RS.MoveFirst
   'RS.Find "Paper='" & rs1!paper & "'"
   'If RS.EOF = False Then
   '   a1_ = RS(3)
   '   c_ream = Int(RS(2))
   '   c_sheet = a1_ - Int(a1_)
   '   c_ream = c_ream + Int(a1_)
   'End If
   
   '-------------------------------------------------------
   
   
   con.Execute "insert into tmps_LEDGER1(address1,address2,address3,phone,owner,setupid,fyear,DESCFORINVOICE) values('" & rs1!CoverPaper & "','" & ream & "','" & sheet & "','" & c_ream & "','" & c_sheet & "'," & setupid & ",'" & session & "','" & rs1!Coversize & "')"
  
   rs1.MoveNext
   
Wend





'==========================================================

con.Execute "update tmps_LEDGER1 set YEAROPENING = (convert(float,ADDRESS2)+convert(float ,ADDRESS3)+convert(float ,phone)+convert(float ,Owner)) "
con.Execute "update tmps_LEDGER1 set DISCATEGORY = (convert(float,ADDRESS2)+convert(float ,ADDRESS3)),DISTCODE=(convert(float ,phone)+convert(float ,Owner))"


DSNNew


If MsgBox("Want to View", vbYesNo) = vbYes Then
    CR.Reset
    CR.ReportFileName = rptPath & "/Consumption_Paperwise.rpt"
    CR.Connect = constr
    CR.WindowShowPrintSetupBtn = True
    CR.WindowShowPrintBtn = True
    CR.WindowState = crptMaximized
    CR.Action = 1
    Screen.MousePointer = vbDefault
End If


End Sub

Private Sub commandQuit_Click()
Unload Me
End Sub

Private Sub Textfirmname_GotFocus()
  ref
End Sub
Sub ref()
'mode = ""
'
'frmbill.Enabled = True
'
'
'
'Me.Add.SetFocus
'
''End If
'PopUpValue1 = ""
'PopUpValue2 = ""
'PopUpValue3 = ""

End Sub
Private Sub Textfirmname_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
popuplist1 "Select firm_name,firm_id, firmpictfilename from firmMaster", con
End If
End Sub

Private Sub textpapersize_GotFocus()

If PopUpValue1 <> "" Then
   TextPaperSize.text = PopUpValue1
   txtPcode.text = PopUpValue6
   
   lblPaper_det.Caption = "Paper Name : " & PopUpValue2 & vbCrLf & "Paper Type : " & PopUpValue3 & vbCrLf & "Real/Sheets : " & popupvalue4 & vbCrLf & "Qulality && G.S.M. : " & popupvalue5
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   

End If

End Sub

Private Sub textpapersize_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   
  searchType = "paper"
   value = "Select SizeValue1 + 'X'+ SizeValue2 as [Paper Size]," & _
   "papermaker_name as [Paper Name],PType,Size as [Sheet/Real],Eco + ' - ' + GSM  as [Quality & GSM],papermaker_Id as Code from " & _
   " papermakemaster where " & stringyear & " and papermaker_id <> '' order by SizeValue1"
    popuplistModel10 value, con
       

End If


End Sub
Sub Clearvalue()

PopUpValue6 = ""

ok.Enabled = True
Edit.Enabled = False
delete.Enabled = False
'cancel.Enabled = False

vs.Clear
grid_ini

txtFirmName.ListIndex = -1
txtPcode = ""
lblPaper_det.Caption = ""
Me.txtOrdNo.text = ""
TextPaperSize = ""
lblAdd.Caption = ""
binder_name = ""

txtTReam = ""
txtTSheet = ""

PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
popupvalue4 = ""

End Sub
Private Sub TextPaperSize_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   vs.SetFocus
   vs.Col = 0
End If



End Sub


Private Sub txtNote_GotFocus()
 
'Call vs_GotFocus


End Sub

Private Sub txtNote_KeyDown(KeyCode As Integer, Shift As Integer)
       If KeyCode = 27 Then
          If txtNote.Visible = True Then
             vs.TextMatrix(vs.RowSel, 20) = txtNote.text
             txtNote.text = ""
             txtNote.Visible = False
             vs.Col = 20
             vs.SetFocus
          End If
       End If
End Sub

Private Sub VSFlexGrid1_Click()

End Sub


