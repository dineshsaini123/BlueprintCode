VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmbill 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Printing Order"
   ClientHeight    =   9588
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   19980
   Icon            =   "printing.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9588
   ScaleWidth      =   19980
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtbalreams 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   13275
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   8100
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ComboBox txtFirmName 
      Height          =   315
      ItemData        =   "printing.frx":000C
      Left            =   7020
      List            =   "printing.frx":000E
      TabIndex        =   2
      Top             =   240
      Width           =   4695
   End
   Begin VB.TextBox TextPaperSize 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   13080
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtPcode 
      Height          =   285
      Left            =   13860
      TabIndex        =   28
      Top             =   1188
      Visible         =   0   'False
      Width           =   672
   End
   Begin VB.ComboBox binder_name 
      Height          =   288
      Left            =   7020
      Style           =   1  'Simple Combo
      TabIndex        =   3
      Top             =   720
      Width           =   4695
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   6012
      Left            =   48
      TabIndex        =   5
      Top             =   1668
      Width           =   19788
      _cx             =   34904
      _cy             =   10604
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
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   420
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"printing.frx":0010
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
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   14640
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   6360
         Width           =   195
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   9450
      Top             =   8685
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFD7AE&
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   8280
      Width           =   9180
      Begin VB.CommandButton cmdPrint_Slip 
         BackColor       =   &H00FFFFFF&
         Caption         =   "P&rint Slip"
         Height          =   585
         Left            =   6870
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton cancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   585
         Left            =   4620
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton ok 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   585
         Left            =   1245
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton Printcmd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   585
         Left            =   5745
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton delete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   585
         Left            =   3495
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton Edit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   585
         Left            =   2370
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton Add 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   585
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton CommandQuit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         Height          =   585
         Left            =   7995
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton cmdBillCancel 
         Caption         =   "&Order Cancel"
         Height          =   585
         Left            =   12240
         TabIndex        =   15
         Top             =   165
         Width           =   75
      End
   End
   Begin VB.TextBox txtOrdNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1020
      MaxLength       =   50
      TabIndex        =   0
      Top             =   180
      Width           =   1410
   End
   Begin MSComCtl2.DTPicker txtOrdDate 
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   195
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   572
      _Version        =   393216
      CalendarBackColor=   16776960
      Format          =   160432129
      CurrentDate     =   38372
   End
   Begin VB.Label lblpaperbal 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance Reams :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   11520
      TabIndex        =   31
      Top             =   8145
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Firm Name"
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
      Height          =   285
      Left            =   5940
      TabIndex        =   30
      Top             =   240
      Width           =   1470
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
      Left            =   12540
      TabIndex        =   29
      Top             =   1260
      Visible         =   0   'False
      Width           =   495
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
      Height          =   255
      Left            =   6000
      TabIndex        =   19
      Top             =   7800
      Width           =   675
   End
   Begin VB.Label lblAdd 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   7035
      TabIndex        =   18
      Top             =   1155
      Width           =   4680
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
      Height          =   330
      Left            =   7440
      TabIndex        =   17
      Top             =   7740
      Width           =   825
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
      Height          =   330
      Left            =   6600
      TabIndex        =   16
      Top             =   7740
      Width           =   825
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
      TabIndex        =   14
      Top             =   9180
      Width           =   2805
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
      Left            =   -240
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   45
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
      Left            =   -360
      TabIndex        =   12
      Top             =   5160
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Paper Size"
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
      Height          =   225
      Left            =   12540
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Orderl Date"
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
      Height          =   315
      Left            =   2880
      TabIndex        =   10
      Top             =   210
      Width           =   1380
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Printer Name"
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
      Height          =   285
      Left            =   5910
      TabIndex        =   9
      Top             =   765
      Width           =   1470
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order No."
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
      Height          =   255
      Left            =   90
      TabIndex        =   8
      Top             =   210
      Width           =   1155
   End
End
Attribute VB_Name = "frmbill"
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

Dim page_sum As Double
Public gridchk As Boolean
Public bo_ As Boolean
Dim sheet, ream, westage
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

    
    Me.vs.Clear
    
    Me.vs.Cols = 23
    
    
    Me.vs.rows = 100
    Me.vs.ColWidth(0) = 1000
    Me.vs.ColWidth(1) = 800
    Me.vs.ColWidth(2) = 1800
    Me.vs.ColWidth(3) = 600
    Me.vs.ColWidth(4) = 1000 '"Inner" '600
    Me.vs.ColWidth(5) = 600  'Text 600
    Me.vs.ColWidth(6) = 500  'Exam 600
    Me.vs.ColWidth(7) = 0    'Supp 600
    Me.vs.ColWidth(8) = 800
    Me.vs.ColWidth(9) = 0     '"T.Page" 700
    Me.vs.ColWidth(10) = 650
    Me.vs.ColWidth(11) = 700
    Me.vs.ColWidth(12) = 650
    Me.vs.ColWidth(13) = 600
    Me.vs.ColWidth(14) = 600
    Me.vs.ColWidth(15) = 1700
    Me.vs.ColWidth(16) = 0
    Me.vs.ColWidth(17) = 1800
    Me.vs.ColWidth(18) = 1400
    
    Me.vs.ColWidth(20) = 600
    Me.vs.ColWidth(21) = 1100
    Me.vs.ColWidth(22) = 1100
    
    
    
    
    Me.vs.TextMatrix(0, 0) = "Code"
    Me.vs.TextMatrix(0, 1) = "Fresh/ITC"
    Me.vs.TextMatrix(0, 2) = "Particulars"
    Me.vs.TextMatrix(0, 3) = "Qty"
    Me.vs.TextMatrix(0, 4) = "Book Part"  '"Inner"
    Me.vs.TextMatrix(0, 5) = "Page Qty"  '"Text"
    Me.vs.TextMatrix(0, 6) = "Price"  '"Exam."
    Me.vs.TextMatrix(0, 7) = ""  '"Supp."
    Me.vs.TextMatrix(0, 8) = "Color"
    Me.vs.TextMatrix(0, 9) = ""  '"T.Page"
    
    Me.vs.TextMatrix(0, 10) = "DIV(8/16)"
    Me.vs.TextMatrix(0, 11) = "T.From"
    Me.vs.TextMatrix(0, 12) = "Wast(%)"
    Me.vs.TextMatrix(0, 13) = "Reams"
    Me.vs.TextMatrix(0, 14) = "Sheet"
    Me.vs.TextMatrix(0, 15) = "Binder"
    Me.vs.TextMatrix(0, 16) = ""
    Me.vs.TextMatrix(0, 17) = "Paper Size"
    Me.vs.TextMatrix(0, 18) = "Remarks"
    Me.vs.TextMatrix(0, 19) = "TrimSize"
    Me.vs.TextMatrix(0, 20) = "CD"
    Me.vs.TextMatrix(0, 21) = "Binding"
    Me.vs.TextMatrix(0, 22) = "Printer"
    
    
   
'   For k1 = 0 To vs.Cols - 1
'     vs.Cell(flexcpFontSize, k1) = 11
'   Next
       
    
    
End Sub
Sub adddisab()
Me.Edit.Enabled = False
Me.Printcmd.Enabled = False
Me.cmdBillCancel.Enabled = False
End Sub
Private Sub Add_Click()




Clearvalue

mode = ""

txtOrdDate.value = Format(Date, "dd/MM/yyyy")

txtOrdNo.text = MaxOrderNo(txtFirmName)


If txtOrdNo <> "" Then
con.Execute "delete from tmporder where orderno='" & txtOrdNo & "'"
End If

txtOrdNo.SetFocus

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
If KeyAscii = 13 Then

    If binder_name.text = "" Then
       MsgBox "Please Choose the Printers", vbInformation
       Me.binder_name.SetFocus
       Exit Sub
    End If
    
    sendkeys "{tab}"


End If

End Sub

Private Sub cancel_Click()

mode = ""
Me.binder_name.text = ""
grid_ini


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
Private Sub CommandPAPERSTATEMENT_Click()
''''paperstatementcalc
''''        If Textfirmid.Text = "MITTAL" Then
''''        frmbill.cr1.ReportFileName = App.Path & "\mpaperstat.rpt"
''''        ElseIf Textfirmid.Text = "DAYAL" Then
''''        frmbill.cr1.ReportFileName = App.Path & "\dpaperstat.rpt"
''''        Else
''''        frmbill.cr1.ReportFileName = App.Path & "\paperstat.rpt"
''''        End If
''''frmbill.cr1.DataFiles(0) = ""
''''frmbill.cr1.DataFiles(1) = ""
''''frmbill.cr1.DataFiles(2) = ""
''''frmbill.cr1.DataFiles(3) = ""
''''frmbill.cr1.DataFiles(4) = ""
''''frmbill.cr1.DataFiles(0) = App.Path & "\" + Trim(main.directory) + "\data.mdb"
''''frmbill.cr1.DataFiles(1) = App.Path & "\" + Trim(main.directory) + "\data.mdb"
''''frmbill.cr1.DataFiles(2) = App.Path & "\" + Trim(main.directory) + "\data.mdb"
''''frmbill.cr1.DataFiles(3) = App.Path & "\" + Trim(main.directory) + "\data.mdb"
''''{billMASTER.bill_id} = "151" AND {BILLMASTER.FIRM_ID} = "NEERAJ"
''''MRSF = "{billMASTER.firm_id} = '" & frmbill.Textfirmid & "' and  {billMASTER.bill_id} = '" & frmbill.bill_no & "'"
''''MRSF = ""
''''frmbill.cr1.ReplaceSelectionFormula (MRSF)
''''frmbill.cr1.Destination = 1
''''frmbill.cr1.Action = 1
End Sub

Private Sub cmdPrint_Slip_Click()

DSNNew

cr1.Reset
cr1.ReportFileName = rptPath & "/BinderSlip.rpt"
'cr1.Connect = "filedsn=chitradsn;uid= " & sql_user  & ";pwd=sidc;"
cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
cr1.ReplaceSelectionFormula "{orderPrint_Main.ord_no}='" & txtOrdNo.text & "' and {orderPrint_Main.FirmName}='" & txtFirmName.text & "'"
cr1.WindowShowPrintSetupBtn = True
cr1.WindowState = crptMaximized
cr1.Action = 1

End Sub

Private Sub Delete_Click()


X = MsgBox("Are you sure you wish to delete the selected Bill ", 4, "Confirmation")
If X = 6 Then
   sq = "delete  from OrderPrint_Main where FirmName='" & txtFirmName & "' and Ord_No = '" + txtOrdNo.text + "' and " & stringyear
   con.Execute sq
   sq = "delete  from OrderPrint_Det where FirmName='" & txtFirmName & "' and Ord_No = '" + txtOrdNo.text + "' and " & stringyear
   con.Execute sq
   Call Add_Click
End If


End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
binder_name.Enabled = True
Me.binder_name.SetFocus
End If
End Sub
Private Sub Edit_Click()
Me.ok.Enabled = True
delete.Enabled = True
cancel.Enabled = True
Me.Edit.Enabled = False
mode = "edit"
ok.SetFocus
End Sub

Private Sub Form_Activate()
'Add_Click

'Me.Width = 20076
'Me.Height = 10032

Me.WindowState = vbMaximized

End Sub
Sub MaxNo()
    If RS.State = 1 Then RS.close
    RS.Open "select max(val(bill_id)) from billmaster", con, adOpenDynamic, adLockOptimistic
    If Not IsNull(RS.Fields(0).value) Then
       bill_no.text = RS.Fields(0).value
    End If
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

If bk = "" Then
vs.ColComboList(17) = s
Else
PopUpValue6 = s
End If


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

vs.ColComboList(15) = st_

End Sub
Private Sub Form_Load()

Me.Left = 10
Me.top = 10
 
Me.Width = 13965
Me.Height = 9990

grid_ini
bkfont = "e"
fid = "chitra"


'If UserName = "admin" Then
'   vs.Width = 19700
'Else
'   vs.Width = 15000
'End If



'
txtFirmName.Clear
If RS.State = 1 Then RS.close
RS.Open "select FirmName,Add1,Add2 from FirmMaster order by firmname", con, adOpenStatic, adLockReadOnly
While RS.EOF = False
  txtFirmName.AddItem RS(0)
  RS.MoveNext
Wend

txtFirmName.ListIndex = 0

'-------------------------------------------------
addbinder
'-------------------------------------------------
addpaper_

s21 = ""

s21 = "Single Color|Double Color|Four Color"
vs.ColComboList(8) = s21


s21 = ""
If RS.State = 1 Then RS.close
RS.Open "select  name from MasterTbl where Category='bkpart'", con
While RS.EOF = False
   If s21 = "" Then
      s21 = RS(0)
   Else
      s21 = s21 & "|" & RS(0)
   End If
RS.MoveNext
Wend
vs.ColComboList(4) = s21






s21 = ""
If RS.State = 1 Then RS.close
RS.Open "select godwn from Godownmaster where len(godwn)>8  order by godwn", con, adOpenKeyset, adLockReadOnly
While RS.EOF = False
   If s21 = "" Then
      s21 = RS(0)
   Else
      s21 = s21 & "|" & RS(0)
   End If
RS.MoveNext
Wend
vs.ColComboList(22) = s21








txtOrdDate.value = Format(Date, "dd/MM/yyyy")
BackColorFrom Me




s21 = ""
If RS.State = 1 Then RS.close
RS.Open "select  name from MasterTbl where Category='trimsize'", con
While RS.EOF = False
   If s21 = "" Then
      s21 = RS(0)
   Else
      s21 = s21 & "|" & RS(0)
   End If
RS.MoveNext
Wend
vs.ColComboList(19) = s21



vs.ColComboList(20) = "Yes|No"

s21 = ""
If RS.State = 1 Then RS.close
RS.Open "select  name from MasterTbl where Category='binding'", con
While RS.EOF = False
   If s21 = "" Then
      s21 = RS(0)
   Else
      s21 = s21 & "|" & RS(0)
   End If
RS.MoveNext
Wend
vs.ColComboList(21) = s21



'vs.ColComboList(20) = "Sectin Swing|Pin Binding|Back Cut"
'--------------------------------------------------------

mode = ""

'If RS.State = 1 Then RS.close
'RS.Open "select max(convert(int,ord_no)) from OrderPrint_main where FirmName='" & txtFirmName & "' and " & stringyear, con
'If Not IsNull(RS(0)) Then
'   txtOrdNo.Text = RS(0) + 1
'Else
'   txtOrdNo.Text = 1
'End If

txtOrdNo.text = MaxOrderNo(txtFirmName)

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
Private Sub txtFirmName_LostFocus()

If txtFirmName = "" Then Exit Sub

'If RS.State = 1 Then RS.close
'RS.Open "select max(convert(int,ord_no)) from OrderPrint_main where FirmName='" & txtFirmName & "' and " & stringyear, con
'If Not IsNull(RS(0)) Then
'   txtOrdNo.Text = RS(0) + 1
'Else
'   txtOrdNo.Text = 1
'End If

txtOrdNo.text = MaxOrderNo(txtFirmName)

End Sub

Private Sub txtOrdDate_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then txtFirmName.SetFocus
End Sub

Private Sub txtOrdNo_GotFocus()
 If PopUpValue1 <> "" Then
    txtOrdNo.text = PopUpValue1
    
    txtbalreams.Visible = True
    lblpaperbal.Visible = True

    searchData
    PopUpValue1 = ""
 End If
End Sub
Private Sub txtOrdNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   
   If txtFirmName = "" Then
     popuplist1 "Select Ord_No,convert(nvarchar,Ord_Date,103) as OrdDate,PrinterName from PrintOrderQry order by convert(int,Ord_No),Ord_Date", con
   
   Else
     popuplist1 "Select Ord_No,convert(nvarchar,Ord_Date,103) as OrdDate,PrinterName from PrintOrderQry where FirmName='" & txtFirmName & "' order by convert(int,Ord_No),Ord_Date", con
   End If
   
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
     ream_ = ream_ + Val(vs.TextMatrix(I, 13))
     sheet_ = sheet_ + Val(vs.TextMatrix(I, 14))
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
         vs.RemoveItem vs.Row
         addTotalReam
         PopUpValue6 = "y"
         vs.SetFocus
   End If
End If


End Sub
Sub Ream_Sheet()

Dim Tot, wastage_per, Form, wream


wastage_per = 0

Form = Val(vs.TextMatrix(vs.RowSel, 11))
quan = Val(vs.TextMatrix(vs.RowSel, 3))
wastage_per = Val(vs.TextMatrix(vs.RowSel, 12))

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
   ream = ream + wream
   
   If sheet > 499 Then
      wream = Int(sheet / 500)
      sheet = sheet - (wream * 500)
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




If vs.Col <> 2 Then
    Label11.Visible = False
    Label12.Visible = False
    Label16.Visible = False
End If

If vs.Col = 2 Then
    Label11.Visible = True
    Label12.Visible = True
    Label16.Visible = True
End If



If vs.Col = 2 Or vs.Col = 21 Then
        If KeyAscii = 8 Then
        If Len(Trim(vs.text)) <> 0 Then
                vs.text = Left(vs.text, (Len(vs.text) - 1))
        End If
            ElseIf (KeyAscii >= 32 And KeyAscii <= 126) Then
            vs.text = vs.text + Chr(KeyAscii)
        End If
End If


If vs.Col = 3 Then

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

If vs.Col = 5 Then

    If KeyAscii = 119 Or KeyAscii = 87 Then
       vs.text = "Web"
       web.Visible = False
    End If
    If KeyAscii = 115 Or KeyAscii = 83 Then
       vs.text = "Sheet"
       web.Visible = False
    End If

End If

If (vs.Col = 4 Or vs.Col = 5 Or vs.Col = 6 Or vs.Col = 7 Or vs.Col = 8 Or vs.Col = 12 Or vs.Col = 13 Or vs.Col = 14 Or vs.Col = 15 Or vs.Col = 16) Then
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
   
For P1 = 5 To 5
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
rs1.Open "select Head1,Head2,Head3,Head4,Head5,txtHead6 as Head6,txtHead7 as Head7," & _
"txtHead8 as Head8,txtHead9 as Head9,txtHead10 as Head10,txtHead11 as Head11,txtHead12 as Head12 from BookMaster where bookno='" & bk & "'", con

If rs1.EOF = False Then
For Q1 = 1 To 12


    If Not IsNull(rs1.Fields("head" & Q1).value) Then
    If rs1.Fields("head" & Q1).value <> "" Then
        If s11 = "" Then
           s11 = rs1.Fields("head" & Q1).value
        Else
           s11 = s11 & "|" & rs1.Fields("head" & Q1).value
        End If
    End If
    End If





Next

End If

vs.ColComboList(4) = s11

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

vs.ColComboList(4) = s11

End If



  
   
End Sub
Sub fatchRawsData(r1 As Integer, r2 As Integer)
    
    For m1 = 0 To 2
        vs.TextMatrix(r2, m1) = vs.TextMatrix(r1, m1)
    Next
    
    
    vs.TextMatrix(r2, 6) = vs.TextMatrix(r1, 6)
    
    vs.TextMatrix(r2, 10) = vs.TextMatrix(r1, 10)
    
    vs.TextMatrix(r2, 15) = vs.TextMatrix(r1, 15)
    
    vs.TextMatrix(r2, 18) = vs.TextMatrix(r1, 18)
    vs.TextMatrix(r2, 19) = vs.TextMatrix(r1, 19)
    
    vs.TextMatrix(r2, 20) = vs.TextMatrix(r1, 20)
    
    vs.TextMatrix(r2, 21) = vs.TextMatrix(r1, 21)
    
    
    
    
    
End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)

On Error Resume Next

If KeyCode = 13 Then
   
If txtFirmName.text = "" Then
   MsgBox "Select FirmName ....", vbCritical
   txtFirmName.SetFocus
   Exit Sub
End If
   
If vs.Col = 0 Then
   sendkeys "{right}"
End If
   
If vs.Col = 1 Then
   
   If PopUpValue6 = "y" Then
      PopUpValue6 = ""
      Exit Sub
   End If
   
   If RS.State = 1 Then RS.close
   RS.Open "select book,book_unit,DivideValue,HeadData1,HeadData2,HeadData3,HeadData4,HeadData5,bookfont,Price,trimsize,binder,binding,remarks from BookMaster " & _
   " where " & stringyear & " and BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "' and firmname='" & txtFirmName.text & "'", con
   If RS.EOF = True Then
      MsgBox "Please Enter Valid Book Code according to Firm Name...", vbCritical
      Exit Sub
   End If
   
   If RS.EOF = False Then
       
      addbinder
      
      
      vs.TextMatrix(vs.RowSel, 0) = UCase(vs.TextMatrix(vs.RowSel, 0))
      
      vs.TextMatrix(vs.RowSel, 2) = UCase(RS!Book)
      
      vs.TextMatrix(vs.RowSel, 18) = RS!remarks & ""
      vs.TextMatrix(vs.RowSel, 19) = RS!trimsize & ""
      vs.TextMatrix(vs.RowSel, 20) = "No"
      
      vs.TextMatrix(vs.RowSel, 10) = Val(RS!DivideValue)
      vs.TextMatrix(vs.RowSel, 11) = Round(Val(vs.TextMatrix(vs.RowSel, 9)) / Val(vs.TextMatrix(vs.RowSel, 10)), 2)
      vs.TextMatrix(vs.RowSel, 6) = RS!Price & ""
      
      vs.TextMatrix(vs.RowSel, 15) = RS!Binder & ""
      vs.TextMatrix(vs.RowSel, 21) = RS!Binding & ""
      
      
      If vs.TextMatrix(vs.RowSel, 6) = "" Then
        If RS.State = 1 Then RS.close
        RS.Open "select RATE from books where BOOKCODE='" & vs.TextMatrix(vs.RowSel, 0) & "' and firmname='" & txtFirmName.text & "'", con
        If RS.EOF = False Then
           vs.TextMatrix(vs.RowSel, 6) = RS(0)
        End If
      End If
      
      addType vs.TextMatrix(vs.RowSel, 0)
      
      
      
      
      
      
      'Add All Daya================================================================
       
       
       If rs1.State = 1 Then rs1.close
       rs1.Open "select * from tmporder where (orderNo='" & txtOrdNo & "' and Bkcode='" & vs.TextMatrix(vs.RowSel, 0) & "')", con, adOpenDynamic, adLockOptimistic
       If rs1.EOF = True Then
          rs1.AddNew
          rs1!orderNo = Trim(txtOrdNo)
          rs1!bkcode = Trim(vs.TextMatrix(vs.RowSel, 0))
          rs1.update
       Else
          GoTo dinesh
       End If
       
       
       bo_ = False
       l1 = 1
       
       If rs1.State = 1 Then rs1.close
       rs1.Open "select Head1,Head2,Head3,Head4,Head5,txthead6 as Head6,txthead7 as Head7," & _
       "txthead8 as Head8,txthead9 as Head9,txthead10 as Head10,txthead11 as Head11,txthead12 as Head12," & _
       "HeadData1,HeadData2,HeadData3,HeadData4,HeadData5,txtheadData6 as HeadData6," & _
       "txtheadData7 as HeadData7,txtheadData8 as HeadData8,txtheadData9 as HeadData9,txtheadData10 as HeadData10,txtheadData11 as HeadData11,txtheadData12 as HeadData12," & _
       "pcode1,pcode2,pcode3,pcode4,pcode5,pcode6,pcode7,pcode8,txtPCode9 as pcode9,txtPCode10 as pcode10,txtPCode11 as pcode11,txtPCode12 as pcode12," & _
       "color1,color2,color3,color4,color5,color6,color7,color8,color9,color10,color11,color12," & _
       "Inn_Printer,text_Printer,Exam_Printer,Supp_Printer,Title_Printer,cboPrinter6,cboPrinter7," & _
       "cboPrinter8,cboPrinter9,cboPrinter10,cboPrinter11,cboPrinter12 from BookMaster " & _
       " where " & stringyear & " and BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "' and firmname='" & txtFirmName.text & "'", con
       
       If rs1.EOF = False Then
       For b1 = 1 To 12
          
          PopUpValue6 = ""
          
   
          
          If Not IsNull(rs1.Fields("Head" & b1)) Then
             s1 = InStr(rs1.Fields("Head" & b1), "ITC")
          If vs.TextMatrix(vs.RowSel, 1) = "Fresh" Then
             s2 = 0
          Else
             s2 = 1
          End If
          
          If s1 <> s2 Then
             GoTo gg:
          End If
                         
          If bo_ = False Then
               
              If b1 >= 6 Then
               vs.TextMatrix(vs.RowSel, 4) = rs1.Fields("Head" & b1)
               vs.TextMatrix(vs.RowSel, 5) = rs1.Fields("HeadData" & b1)
              Else
               vs.TextMatrix(vs.RowSel, 4) = rs1.Fields("Head" & b1)
               vs.TextMatrix(vs.RowSel, 5) = rs1.Fields("HeadData" & b1)
              End If
              
               addpaper_ rs1.Fields("pcode" & b1), vs.RowSel
               vs.TextMatrix(vs.RowSel, 17) = PopUpValue6
               vs.TextMatrix(vs.RowSel, 8) = rs1.Fields("color" & b1)
               
                PopUpValue7 = b1
                
                 If Not IsNull(rs1.Fields(fatchPrinter("" & PopUpValue7))) Then
                    vs.TextMatrix(vs.RowSel, 22) = rs1.Fields("" & fatchPrinter("" & PopUpValue7))
                 End If
               
               bo_ = True
               
             
          Else
           
           
           If rs1.Fields("Head" & b1) <> "" Then
               
               
              If b1 >= 6 Then
               vs.TextMatrix(vs.RowSel + l1, 4) = rs1.Fields("Head" & b1)
               vs.TextMatrix(vs.RowSel + l1, 5) = rs1.Fields("HeadData" & b1)
              Else
               vs.TextMatrix(vs.RowSel + l1, 4) = rs1.Fields("Head" & b1)
               vs.TextMatrix(vs.RowSel + l1, 5) = rs1.Fields("HeadData" & b1)
              End If

               
               fatchRawsData vs.RowSel, vs.RowSel + l1
               addpaper_ rs1.Fields("pcode" & b1), vs.RowSel + l1
               vs.TextMatrix(vs.RowSel + l1, 17) = PopUpValue6
               vs.TextMatrix(vs.RowSel + l1, 8) = rs1.Fields("color" & b1)
               
                PopUpValue7 = b1
                
                 If Not IsNull(rs1.Fields(fatchPrinter("" & PopUpValue7))) Then
                    vs.TextMatrix(vs.RowSel + l1, 22) = rs1.Fields("" & fatchPrinter("" & PopUpValue7))
                 End If
               
               
               
               ''ok
               
               l1 = l1 + 1
               
           
           
           End If
               
          End If
          
          End If
          
gg:
          
        Next
       End If
       PopUpValue6 = ""
      '============================================================================
dinesh:



      
      
      sendkeys "{right}"
      sendkeys "{right}"
      
   End If

ElseIf vs.Col = 3 Then
      
      Dim T_ As Boolean
      T_ = False
      
      Qty1 = IIf(Val(vs.TextMatrix(vs.RowSel, 3)) = 0, 0, Val(vs.TextMatrix(vs.RowSel, 3)))
      
      If rs1.State = 1 Then rs1.close
      rs1.Open "select qty1,qty2,Wastage from WastageQty order by convert(int,Wastage) desc", con
      While rs1.EOF = False
         
         T_ = False
         
         If rs1!Qty1 >= Qty1 Then
               T_ = True
         End If
         
         If Qty1 <= rs1!Qty2 Then
               T_ = True
         End If
            
         
         If T_ = True Then
            vs.TextMatrix(vs.RowSel, 12) = rs1!Wastage
            GoTo aaa1:
         End If
         
         rs1.MoveNext
      Wend
      
aaa1:
      
      sendkeys "{right}"
      
ElseIf vs.Col = 4 Then
   
     
   
     If rs1.State = 1 Then rs1.close
     rs1.Open "select Head1,Head2,Head3,Head4,Head5,txthead6 as Head6,txthead7 as Head7,txthead8 as Head8,HeadData1,HeadData2,HeadData3,HeadData4,HeadData5,txtHeadData6 as HeadData6,txtHeadData7 as HeadData7,txtHeadData8 as HeadData8 from BookMaster " & _
     " where " & stringyear & " and BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "' and firmname='" & txtFirmName.text & "'", con
     If rs1.EOF = False Then
        For b1 = 1 To 5
            If vs.TextMatrix(vs.RowSel, 4) = rs1.Fields("Head" & b1) Then
               vs.TextMatrix(vs.RowSel, 5) = rs1.Fields("HeadData" & b1)
               bo_ = True
            End If
        Next
        
       
        
     End If

     sendkeys "{right}"
     vs.TextMatrix(vs.RowSel, 9) = addPage
     
ElseIf vs.Col = 5 Then
      
        
        
            
            
            If rs1.State = 1 Then rs1.close
             rs1.Open "select Head1,Head2,Head3,Head4,Head5,txthead6 as Head6,txthead7 as Head7,txthead8 as Head8,HeadData1,HeadData2,HeadData3,HeadData4,HeadData5,txtHeadData6 as HeadData6,txtHeadData7 as HeadData7,txtHeadData8 as HeadData8 from BookMaster " & _
             " where " & stringyear & " and BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "' and firmname='" & txtFirmName.text & "'", con
            For b1 = 1 To 5
            If (rs1.Fields("Head" & b1) = vs.TextMatrix(vs.RowSel, 4)) Then
            
                v1_ = Val(vs.TextMatrix(vs.RowSel, 5))
                v2_ = Val(rs1.Fields("Headdata" & b1).value)
                
                If v1_ <> v2_ Then
                    If MsgBox("Want to add in master", vbQuestion + vbYesNo) = vbYes Then
                         con.Execute "update BookMaster set " & rs1.Fields("Headdata" & b1).Name & "='" & vs.TextMatrix(vs.RowSel, 5) & "'" & " where BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'"
                    End If
                End If
                
                 
                GoTo aaaa:
             End If
             Next
    
    
        
        
aaaa:
      vs.SetFocus

      sendkeys "{right}"
      vs.TextMatrix(vs.RowSel, 9) = addPage
ElseIf vs.Col = 6 Then

  
   
  If rs1.State = 1 Then rs1.close
  rs1.Open "select PRICE from BookMaster " & _
  " where BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
  If rs1.EOF = False Then
  
  If Val(vs.TextMatrix(vs.RowSel, 6)) <> rs1(0) Then
    If MsgBox("Want to add in master", vbQuestion + vbYesNo) = vbYes Then
       con.Execute "update BookMaster set price='" & vs.TextMatrix(vs.RowSel, 6) & "' where BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'"
       vs.SetFocus
    End If
  End If
  End If
 

  
   If vs.TextMatrix(vs.RowSel, 6) <> "" Then
      sendkeys "{right}"
   End If
      
   vs.TextMatrix(vs.RowSel, 9) = addPage
      
      
ElseIf vs.Col = 8 Then
      
      sendkeys "{right}"
      
      vs.TextMatrix(vs.RowSel, 9) = addPage
      
ElseIf vs.Col = 10 Then
      If RS.State = 1 Then RS.close
      RS.Open "select DivideValue from BookMaster where bookno='" & vs.TextMatrix(vs.RowSel, 0) & "' and firmname='" & txtFirmName.text & "'", con
      If RS.EOF = False Then
        If (Val(RS!DivideValue) <> Val(vs.TextMatrix(vs.RowSel, 10))) Then
          If MsgBox("Want to add in master", vbQuestion + vbYesNo) = vbYes Then
             con.Execute "update BookMaster set DivideValue='" & vs.TextMatrix(vs.RowSel, 10) & "' where BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'"
          End If
        End If
      End If
      
      vs.SetFocus
      vs.TextMatrix(vs.RowSel, 11) = Round(Val(vs.TextMatrix(vs.RowSel, 9)) / Val(vs.TextMatrix(vs.RowSel, 10)), 2)
      sendkeys "{right}"
      
      Ream_Sheet

      vs.TextMatrix(vs.RowSel, 9) = addPage

      
ElseIf vs.Col = 11 Then
      sendkeys "{right}"
ElseIf vs.Col = 12 Then
      
      checkReams

      sendkeys "{right}"
      Ream_Sheet
      vs.TextMatrix(vs.RowSel, 13) = ream
      vs.TextMatrix(vs.RowSel, 14) = sheet
      addTotalReam
      sendkeys "{right}"
      sendkeys "{right}"
ElseIf vs.Col = 15 Then
      
    If vs.TextMatrix(vs.RowSel, 15) <> "" Then
       'SendKeys "{home}"
       'SendKeys "{down}"
       sendkeys "{right}"
       
    End If
    
ElseIf vs.Col = 17 Then
      
    'If vs.TextMatrix(vs.RowSel, 15) <> "" Then
    '  SendKeys "{home}"
    '  SendKeys "{down}"
      checkReams
      sendkeys "{right}"
      addTotalReam
    'End If
    
ElseIf vs.Col = 18 Then
      
      sendkeys "{right}"
      addTotalReam
      
ElseIf vs.Col = 19 Then
      
      
''       If rs1.State = 1 Then rs1.close
''       rs1.Open "select trimsize from BookMaster " & _
''        " where BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'", con
''       If rs1.EOF = False Then
''
''        If rs1(0) = "" Or IsNull(rs1(0)) Then
''          If MsgBox("Want to add in master", vbQuestion + vbYesNo) = vbYes Then
''             con.Execute "update BookMaster set trimsize='" & vs.TextMatrix(vs.RowSel, 18) & "' where BookNo='" & vs.TextMatrix(vs.RowSel, 0) & "'"
''             vs.SetFocus
''          End If
''       End If
''       End If
      
      
      sendkeys "{right}"
      addTotalReam
      
ElseIf vs.Col = 20 Then
      
      sendkeys "{right}"
      addTotalReam
      
ElseIf vs.Col = 21 Then
      
      sendkeys "{right}"
      addTotalReam
ElseIf vs.Col = 22 Then
      
      sendkeys "{home}"
      sendkeys "{down}"
      addTotalReam
    

End If
End If





End Sub
Sub checkReams()

DoEvents
DoEvents

txtbalreams.Visible = True
lblpaperbal.Visible = True


txtbalreams.text = ""

Dim ream_ As Long
Dim sheet_ As Long
Dim per_
ream_ = 0
sheet_ = 0
per_ = 0

pcode_ = Mid(vs.TextMatrix(vs.RowSel, 17), InStr(vs.TextMatrix(vs.RowSel, 17), ">") + 1)

If pcode_ <> "" Then

    str1 = "SELECT ream*-1,sheet*-1 FROM Order_Qry where PrinterName ='" & binder_name.text & "' and PCode=" & pcode_ & _
    " Union All " & _
    "select reams,sheets from paperstatement where (PaperTrans_Deliv='R' and FromGodown ='" & binder_name.text & "') and PCode=" & pcode_ & " " & _
    " Union All " & _
    " select reams,sheets from paperstatement where (PaperTrans_Deliv='D' and toGodown='" & binder_name.text & "' and pcode=" & pcode_ & ")" & _
    " Union All " & _
    " select reams*-1,sheets*-1 from paperstatement where (PaperTrans_Deliv='D' and FromGodown='" & binder_name.text & "' and pcode=" & pcode_ & ")"
    
    If rs1.State = 1 Then rs1.close
    rs1.Open str1, con
    
    While rs1.EOF = False
    
    ream_ = ream_ + rs1(0)
    sheet_ = sheet_ + rs1(1)
    
    rs1.MoveNext
    Wend
    
    
    Dim cal_ream
    Dim D As Integer
    Dim recSheet, tmpReam, tmpSheet
    recSheet = 0
    tmpReam = 0
    tmpSheet = 0
    
    
    recSheet = (ream_ * 500) + sheet_
    
    If recSheet < 0 Then
    
    recSheet = Abs(recSheet)
    
    tmpReam = Int(recSheet / 500)
    tmpSheet = (Round(((recSheet / 500) - Int(recSheet / 500)), 3) * 1000 / 2)
    tmpReam = (tmpReam * -1)
    tmpSheet = (tmpSheet * -1)
    
    Else
    
    tmpReam = Int(recSheet / 500)
    tmpSheet = (Round(((recSheet / 500) - Int(recSheet / 500)), 3) * 1000 / 2)
    
    
    End If
    
    
    
    txtbalreams.text = tmpReam
    txtbalreams.text = txtbalreams.text & "." & Format(tmpSheet, "000")
       
End If


End Sub

Private Sub vs_SelChange()

If (vs.Col = 0 Or vs.Col = 1 Or vs.Col = 2 Or vs.Col = 3 Or vs.Col = 4 Or vs.Col = 5 Or vs.Col = 6 Or vs.Col = 7 Or vs.Col = 8 Or vs.Col = 9 Or vs.Col = 11 Or vs.Col = 14 Or vs.Col = 15 Or vs.Col = 16 Or vs.Col = 17) Then
   vs.Editable = flexEDKbdMouse
   checkReams
ElseIf vs.Col = 19 Then
   vs.Editable = flexEDKbdMouse
   
ElseIf vs.Col = 20 Then
   vs.Editable = flexEDKbdMouse
ElseIf vs.Col = 21 Then
   vs.Editable = flexEDKbdMouse
ElseIf vs.Col = 22 Then
   vs.Editable = flexEDKbdMouse
ElseIf vs.Col = 18 Then
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
Private Sub ok_Click()
 
 con.Execute "delete from temp1"
 
 For I = 1 To vs.rows - 1
    
    If (Len(vs.TextMatrix(I, 22)) > 2) Then
      con.Execute "insert into temp1(text) values('" & vs.TextMatrix(I, 22) & "')"
    End If
 
 Next
 
 Dim rs10 As New ADODB.Recordset
 
 
 
 rs10.Open "select text from temp1 group by text", con
 While rs10.EOF = False
 
 
   If mode = Trim("edit") Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select * from OrderPrint_main where FirmName='" & txtFirmName & "' and Ord_No = '" + txtOrdNo.text + "' and " & stringyear, con, adOpenDynamic, adLockOptimistic
        If rs1.EOF = True Then
            txtOrdNo.text = MaxOrderNo(txtFirmName)
        End If
   Else
            txtOrdNo.text = MaxOrderNo(txtFirmName)
   End If
   
   
   binder_name = rs10(0)
   
   If rs1.State = 1 Then rs1.close
   rs1.Open "select Address,westage from Godownmaster  where godwn='" & binder_name & "'", con, adOpenKeyset, adLockReadOnly
   If rs1.EOF = False Then
      lblAdd.Caption = rs1(0) & ""
   Else
      lblAdd.Caption = ""
   End If

   saveData
   
   rs10.MoveNext
   
 Wend
 
 

End Sub
Sub saveData()

On Error GoTo aa10
 
If binder_name.text = "" Then
   MsgBox "Please Choose the Printers", vbInformation
   Me.binder_name.SetFocus
   Exit Sub
End If

 
 
If mode = Trim("edit") Then
    'sq = "delete  from OrderPrint_Main where Ord_No = '" + txtOrdNo.Text + "' and " & stringyear
    sq = "delete  from OrderPrint_Main where FirmName='" & txtFirmName & "' and Ord_No = '" + txtOrdNo.text + "' and " & stringyear
    con.Execute sq
    sq = "delete  from OrderPrint_Det where FirmName='" & txtFirmName & "' and Ord_No = '" + txtOrdNo.text + "' and " & stringyear
    con.Execute sq
End If
   
If RS.State = 1 Then RS.close
RS.Open "select * from OrderPrint_main where FirmName='" & txtFirmName & "' and Ord_No = '" + txtOrdNo.text + "' and " & stringyear, con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
   RS.AddNew
End If

RS!firmname = txtFirmName.text
RS!Ord_No = txtOrdNo.text
RS!Ord_Date = txtOrdDate.value
RS!PrinterName = binder_name.text
RS!TotalReam = Val(txtTReam.Caption)
RS!TotalSheet = Val(txtTSheet.Caption)
RS!PaperSize = Trim(TextPaperSize)
RS!OrderCancel = "n"
RS!papercode = Trim(txtPcode.text)
RS!Address = Trim(lblAdd.Caption)

RS!fyear = session
RS!setupid = setupid
RS.update


If RS.State = 1 Then RS.close
RS.Open "select * from OrderPrint_det where FirmName='" & txtFirmName & "' and Ord_No = '" + txtOrdNo.text + "' and " & stringyear, con, adOpenDynamic, adLockOptimistic
For K = 1 To vs.rows - 1

If Trim(binder_name) = Trim(vs.TextMatrix(K, 22)) Then

If vs.TextMatrix(K, 1) <> "" Then

    RS.AddNew
    RS!firmname = txtFirmName.text
    RS!Ord_No = txtOrdNo.text
    
    RS!bcode = vs.TextMatrix(K, 0)
    
    RS!Frase_Itc = vs.TextMatrix(K, 1)
    
    RS!Bookname = Trim(vs.TextMatrix(K, 2))
    RS!qty = Val(vs.TextMatrix(K, 3))
    
    
    RS!Book_part = vs.TextMatrix(K, 4)
    RS!PageCount = Val(vs.TextMatrix(K, 5))
    
    
    
    RS!rate = vs.TextMatrix(K, 6)
    RS!supp = Val(vs.TextMatrix(K, 7))
    RS!Title = vs.TextMatrix(K, 8)
    
    RS!tpage = Val(vs.TextMatrix(K, 9))
    RS!DivdeBy = Val(vs.TextMatrix(K, 10))
    If vs.TextMatrix(K, 11) <> "" Then
       RS!TForm = vs.TextMatrix(K, 11)
    End If
    RS!WastPer = Val(vs.TextMatrix(K, 12))
    RS!TotalReam = Val(vs.TextMatrix(K, 13))
    RS!TotalSheet = Val(vs.TextMatrix(K, 14))
    RS!Binder = vs.TextMatrix(K, 15)
    RS!Hindi_English = vs.TextMatrix(K, 16)
    
    aa = Mid(vs.TextMatrix(K, 17), InStr(vs.TextMatrix(K, 17), ":") + 1)
    aa1 = Mid(aa, 1, 5)
    
    RS!PaperSize = vs.TextMatrix(K, 17)
    RS!Size = Trim(aa1)
    RS!pcode = Mid(vs.TextMatrix(K, 17), InStr(vs.TextMatrix(K, 17), "=>") + 2)
    
    RS!remarks = vs.TextMatrix(K, 18)
    
    RS!trimsize = vs.TextMatrix(K, 19)
    RS!cd = vs.TextMatrix(K, 20)
    RS!binder_ = vs.TextMatrix(K, 21)
    
    
    
    RS!fyear = session
    RS!setupid = setupid
    RS.update
    
    con.Execute "delete from tmporder where Bkcode='" & vs.TextMatrix(K, 0) & "' and orderno='" & txtOrdNo & "'"
    
    
    
         
End If


End If

Next

'===================
d1 = 0
s1_ = ""
    
If RS.State = 1 Then RS.close
RS.Open "select BCode from OrderPrint_det where FirmName='" & txtFirmName & "' and Ord_No = '" + txtOrdNo.text + "' and " & stringyear & " order by printorder ", con, adOpenDynamic, adLockOptimistic
For k1 = 1 To RS.RecordCount
    
    
    
    
    If k1 > 1 Then
       If s1_ = RS!bcode Then
          GoTo aa1
       End If
    End If
    d1 = d1 + 1
    con.Execute "update OrderPrint_det set [inner]=" & d1 & " where FirmName='" & txtFirmName & "' and Ord_No = '" + txtOrdNo.text + "' and BCode='" & RS!bcode & "'"
    
aa1:
    
    s1_ = RS!bcode
    RS.MoveNext
    



Next






MsgBox "Data Saved ....", vbInformation
ok.Enabled = False
Edit.Enabled = True


Exit Sub
aa10:
MsgBox "" & err.DESCRIPTION



End Sub
Sub searchData()



grid_ini

If txtOrdNo.text <> "" Then
con.Execute "delete from tmporder where orderNo = '" & txtOrdNo.text & "'"
End If

 
   
If RS.State = 1 Then RS.close
RS.Open "select * from OrderPrint_main where FirmName='" & txtFirmName & "' and Ord_No = '" + txtOrdNo.text + "' and " & stringyear, con
If RS.EOF = False Then
    
    
    con.Execute "insert into tmporder SELECT distinct Ord_No,BCode FROM OrderPrint_Det where FirmName='" & txtFirmName & "' and Ord_No = '" + txtOrdNo.text + "'"
    
    
    ok.Enabled = False
    Edit.Enabled = True
    
    delete.Enabled = False
    cancel.Enabled = False
    
    txtFirmName.text = RS!firmname & ""
    txtPcode.text = RS!papercode & ""
    
    txtOrdNo.text = RS!Ord_No
    txtOrdDate.value = RS!Ord_Date
    binder_name.text = RS!PrinterName
    txtTReam.Caption = RS!TotalReam
    txtTSheet.Caption = RS!TotalSheet
    TextPaperSize = RS!PaperSize
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select address from Godownmaster where " & stringyear & " and Godwn ='" & binder_name.text & "'", con, adOpenKeyset, adLockReadOnly
    If rs1.EOF = False Then
       lblAdd.Caption = rs1!Address & ""
    End If
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from PaperMakeMaster where " & stringyear & " and papermaker_id ='" & txtPcode & "'", con
    If rs1.EOF = False Then
       lblPaper_det.Caption = "Paper Name : " & rs1!papermaker_name & vbCrLf & "Paper Type : " & rs1!ptype & vbCrLf & "Real/Sheets : " & rs1!Size & vbCrLf & "Qulality && G.S.M. : " & rs1!eco & " - " & rs1!GSM
    End If
    
    
    
End If




If RS.State = 1 Then RS.close
RS.Open "select * from OrderPrint_det where  FirmName='" & txtFirmName & "' and Ord_No = '" + txtOrdNo.text + "' and " & stringyear & " order by printorder", con
For K = 1 To vs.rows - 1
    
If RS.EOF = False Then
    vs.TextMatrix(K, 0) = RS!bcode
    
       
    If rs1.State = 1 Then rs1.close
    rs1.Open "select book,bookfont from BookMaster where BookNo='" & RS!bcode & "' and " & stringyear, con, adOpenKeyset, adLockReadOnly

    If rs1.EOF = False Then
        
         vs.TextMatrix(K, 2) = rs1!Book
         
     End If
    
    
    vs.TextMatrix(K, 1) = RS!Frase_Itc & ""
    vs.TextMatrix(K, 3) = RS!qty
    
    vs.TextMatrix(K, 4) = RS!Book_part & ""   'New Code
    vs.TextMatrix(K, 5) = RS!PageCount & ""
    
    vs.TextMatrix(K, 6) = RS!rate & ""
    
    vs.TextMatrix(K, 7) = RS!supp
    vs.TextMatrix(K, 8) = RS!Title
    vs.TextMatrix(K, 9) = RS!tpage
    vs.TextMatrix(K, 10) = RS!DivdeBy
    vs.TextMatrix(K, 11) = RS!TForm & ""
    vs.TextMatrix(K, 12) = RS!WastPer
    vs.TextMatrix(K, 13) = RS!TotalReam
    vs.TextMatrix(K, 14) = RS!TotalSheet
    vs.TextMatrix(K, 15) = RS!Binder
    '16 Blank
    vs.TextMatrix(K, 17) = RS!PaperSize & ""
    vs.TextMatrix(K, 18) = RS!remarks & ""
    
    vs.TextMatrix(K, 19) = RS!trimsize & ""
    vs.TextMatrix(K, 20) = RS!cd & ""
    vs.TextMatrix(K, 21) = RS!binder_ & ""
    
    vs.TextMatrix(K, 22) = binder_name
    
    
    
    RS.MoveNext
End If

Next

addTotalReam


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

'On Error GoTo abc:

Dim reem, sheet, a1, per
con.Execute "delete from tmps_LEDGER1"

sheet = 0

If rs1.State = 1 Then rs1.close
rs1.Open "select  papersize,sum(TotalReam),sum(TotalSheet) size from OrderPrint_Det where FirmName='" & txtFirmName & "' and Ord_No = '" + txtOrdNo.text + "' and " & stringyear & " group by papersize"
While rs1.EOF = False
   
   sheet = 0
   a1 = rs1(2)
   ream = Int(rs1(1))
   
   
   
   If a1 > 499 Then
      per = Int(a1 / 500)
      r_sheet = a1 - per * 500
      ream = ream + per
      sheet = sheet + r_sheet
   Else
      sheet = a1
   End If
    
   
   
   hh1 = InStr(rs1(0), "=>")
   p_p = Mid(rs1(0), 1, hh1 - 1) & " GSM"
   
   
   con.Execute "insert into tmps_LEDGER1(address1,address2,address3,setupid,fyear) values('" & p_p & "','" & ream & "','" & sheet & "'," & setupid & ",'" & session & "')"
  
   rs1.MoveNext
   
Wend

DoEvents
DoEvents
DoEvents


DSNNew

Screen.MousePointer = vbHourglass

cr1.Reset


cr1.ReportFileName = rptPath & "/printorderbp.rpt"
cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
cr1.ReplaceSelectionFormula "{orderPrint_Main.ord_no}='" & txtOrdNo.text & "' and {orderPrint_Main.FirmName}='" & txtFirmName.text & "'"
If RS.State = 1 Then RS.close
RS.Open "select * from firmmaster where firmname='" & txtFirmName & "'", con
If RS.EOF = False Then
 cr1.Formulas(0) = "firmname='" & RS!firmname & "'"
 
 
'If UCase(Trim(txtFirmName.Text)) = "BLUEPRINT EDUCATION" Then
    cr1.Formulas(2) = "add2_='" & RS!add1 & "'"
    cr1.Formulas(3) = "add2='" & RS!add2 & "'"

'Else
'    cr1.Formulas(2) = "add2_='" & RS!add1 & "'"
'     cr1.Formulas(3) = "add2='" & RS!add2 & "'"

'End If
 
End If
cr1.WindowShowPrintSetupBtn = True
cr1.WindowState = crptMaximized
cr1.Action = 1




Screen.MousePointer = vbDefault

    Exit Sub
abc:
    MsgBox "" & err.DESCRIPTION



End Sub

Private Sub commandQuit_Click()
Unload Me
End Sub

Private Sub Textfirmname_GotFocus()
  ref
End Sub
Sub ref()
mode = ""

frmbill.Enabled = True



Me.Add.SetFocus

'End If
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""

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
cancel.Enabled = False

vs.Clear
grid_ini



txtbalreams.text = ""

txtbalreams.Visible = False
lblpaperbal.Visible = False

'txtFirmName.ListIndex = -1
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
