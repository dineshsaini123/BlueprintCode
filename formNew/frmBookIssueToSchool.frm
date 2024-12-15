VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBookIssueToSchool 
   Caption         =   "Book Issue To School"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   17148
   Icon            =   "frmBookIssueToSchool.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   17148
   Begin VB.TextBox txtrepName 
      Appearance      =   0  'Flat
      DataField       =   "Name"
      Height          =   330
      Left            =   6864
      MaxLength       =   100
      TabIndex        =   39
      Top             =   1260
      Width           =   2880
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4995
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   8352
      Width           =   645
   End
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   7548
      TabIndex        =   26
      Top             =   2115
      Visible         =   0   'False
      Width           =   7716
      Begin VB.TextBox txtAreacode 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   4308
         MaxLength       =   100
         TabIndex        =   38
         Top             =   765
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtArea 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   4308
         MaxLength       =   100
         TabIndex        =   36
         Top             =   405
         Width           =   1140
      End
      Begin VB.TextBox txtbcode 
         Appearance      =   0  'Flat
         DataField       =   "Name"
         Height          =   330
         Left            =   3588
         MaxLength       =   100
         TabIndex        =   35
         Top             =   405
         Width           =   690
      End
      Begin VB.CommandButton cmdprint_ 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Teacher List "
         Default         =   -1  'True
         Height          =   645
         Left            =   6696
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   264
         Width           =   924
      End
      Begin VB.ComboBox cbogp 
         Height          =   288
         ItemData        =   "frmBookIssueToSchool.frx":000C
         Left            =   2736
         List            =   "frmBookIssueToSchool.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   405
         Width           =   825
      End
      Begin VB.CommandButton cmdok 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Subject With Teacher Mobile List"
         Height          =   645
         Left            =   5484
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   264
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker txtFrom 
         Height          =   372
         Left            =   0
         TabIndex        =   28
         Top             =   408
         Width           =   1272
         _ExtentX        =   2244
         _ExtentY        =   656
         _Version        =   393216
         Format          =   541786113
         CurrentDate     =   42409
      End
      Begin MSComCtl2.DTPicker txtto 
         Height          =   372
         Left            =   1500
         TabIndex        =   29
         Top             =   408
         Width           =   1176
         _ExtentX        =   2074
         _ExtentY        =   656
         _Version        =   393216
         Format          =   541786113
         CurrentDate     =   42409
      End
      Begin VB.Label Label6 
         Caption         =   "Area"
         Height          =   288
         Left            =   4488
         TabIndex        =   37
         Top             =   180
         Width           =   552
      End
      Begin VB.Label Label5 
         Caption         =   "Category    BookCode"
         Height          =   288
         Left            =   2736
         TabIndex        =   34
         Top             =   180
         Width           =   1632
      End
      Begin VB.Label Label2 
         Caption         =   "To"
         Height          =   312
         Left            =   1272
         TabIndex        =   27
         Top             =   456
         Width           =   192
      End
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00B8E4F1&
      BorderStyle     =   0  'None
      Height          =   750
      Left            =   180
      ScaleHeight     =   756
      ScaleWidth      =   7248
      TabIndex        =   22
      Top             =   2250
      Width           =   7248
      Begin VB.CommandButton Command1_sh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "School Wise"
         Height          =   645
         Left            =   4500
         Picture         =   "frmBookIssueToSchool.frx":0010
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   45
         Width           =   1032
      End
      Begin VB.CommandButton Commandhelp 
         Caption         =   "Help"
         Height          =   495
         Left            =   -720
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton Commandedit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   645
         Left            =   1836
         Picture         =   "frmBookIssueToSchool.frx":0BF4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   45
         Width           =   840
      End
      Begin VB.CommandButton Commandsave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sa&ve"
         Height          =   645
         Left            =   924
         Picture         =   "frmBookIssueToSchool.frx":1036
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   45
         Width           =   924
      End
      Begin VB.CommandButton Commanddelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "De&lete"
         Enabled         =   0   'False
         Height          =   645
         Left            =   2664
         Picture         =   "frmBookIssueToSchool.frx":1C1A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   45
         Width           =   912
      End
      Begin VB.CommandButton Commandsearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Search"
         Height          =   645
         Left            =   3588
         Picture         =   "frmBookIssueToSchool.frx":27FE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   45
         Width           =   888
      End
      Begin VB.CommandButton CommandPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   645
         Left            =   5532
         Picture         =   "frmBookIssueToSchool.frx":33E2
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   45
         Width           =   852
      End
      Begin VB.CommandButton CommandReturn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Return"
         Height          =   645
         Left            =   6384
         Picture         =   "frmBookIssueToSchool.frx":3FC6
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   45
         Width           =   852
      End
      Begin VB.CommandButton Commandadd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   645
         Left            =   45
         Picture         =   "frmBookIssueToSchool.frx":4BAA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   45
         Width           =   876
      End
   End
   Begin VB.TextBox txtScId 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1305
      Width           =   870
   End
   Begin VB.TextBox txtstate 
      Appearance      =   0  'Flat
      DataField       =   "Name"
      Height          =   330
      Left            =   6852
      MaxLength       =   100
      TabIndex        =   5
      Top             =   1665
      Width           =   2916
   End
   Begin VB.TextBox txtcity 
      Appearance      =   0  'Flat
      DataField       =   "Name"
      Height          =   330
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   4
      Top             =   1665
      Width           =   4335
   End
   Begin VB.TextBox txtschool 
      Appearance      =   0  'Flat
      DataField       =   "Name"
      Height          =   330
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   2
      Top             =   855
      Width           =   8328
   End
   Begin VB.TextBox txtcode 
      Appearance      =   0  'Flat
      DataField       =   "Name"
      Height          =   330
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1260
      Width           =   1770
   End
   Begin VB.TextBox txtSno 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "Pub_code"
      Height          =   330
      Left            =   1440
      TabIndex        =   0
      Top             =   135
      Width           =   885
   End
   Begin MSComCtl2.DTPicker txtDate 
      Height          =   315
      Left            =   3105
      TabIndex        =   1
      Top             =   135
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   572
      _Version        =   393216
      CalendarBackColor=   16776960
      Format          =   542638081
      CurrentDate     =   38372
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   5112
      Left            =   96
      TabIndex        =   6
      Top             =   3156
      Width           =   15480
      _cx             =   27305
      _cy             =   9017
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
      BackColorBkg    =   16777215
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
      Rows            =   1000
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   450
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmBookIssueToSchool.frx":578E
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
         Left            =   15444
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   4020
         Width           =   195
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Rep.Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   276
      Index           =   4
      Left            =   5868
      TabIndex        =   40
      Top             =   1308
      Width           =   1428
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F4 Key To Delete row"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   96
      TabIndex        =   25
      Top             =   8424
      Width           =   2952
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   852
      Left            =   132
      Top             =   2208
      Width           =   7368
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For  Search School"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   21
      Top             =   585
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "State :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   3
      Left            =   5850
      TabIndex        =   18
      Top             =   1710
      Width           =   1425
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "City/Village :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   2
      Left            =   90
      TabIndex        =   17
      Top             =   1710
      Width           =   1425
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "School Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   0
      Left            =   90
      TabIndex        =   16
      Top             =   900
      Width           =   1425
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Affilation Code:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   1
      Left            =   90
      TabIndex        =   15
      Top             =   1305
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   8
      Left            =   2475
      TabIndex        =   14
      Top             =   180
      Width           =   705
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "S.No  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   12
      Left            =   90
      TabIndex        =   13
      Top             =   180
      Width           =   1425
   End
End
Attribute VB_Name = "frmBookIssueToSchool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gp As String
Private Sub cmdok_Click()


DSNNew

MainMenu.cr1.Reset
MainMenu.cr1.ReportFileName = rptPath & "/SubjectWithTechMobile.rpt"
MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass

If (UCase(cbogp.text) = "ALL" And txtbcode <> "") Then
    MainMenu.cr1.ReplaceSelectionFormula "({SubjectWithTeacherMobilenoQry.bookcode}='" & txtbcode & "' and {SubjectWithTeacherMobilenoQry.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {SubjectWithTeacherMobilenoQry.invoicedate}<=datevalue('" & Format(txtto.value, "MM/dd/yyyy") & "'))"
ElseIf (UCase(cbogp.text) = "ALL" And txtbcode = "") Then
    MainMenu.cr1.ReplaceSelectionFormula "({SubjectWithTeacherMobilenoQry.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {SubjectWithTeacherMobilenoQry.invoicedate}<=datevalue('" & Format(txtto.value, "MM/dd/yyyy") & "'))"
ElseIf (UCase(cbogp.text) <> "ALL" And txtbcode = "") Then
      MainMenu.cr1.ReplaceSelectionFormula "({SubjectWithTeacherMobilenoQry.gp}='" & cbogp & "' and {SubjectWithTeacherMobilenoQry.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {SubjectWithTeacherMobilenoQry.invoicedate}<=datevalue('" & Format(txtto.value, "MM/dd/yyyy") & "'))"
End If

MainMenu.cr1.Formulas(0) = "fdate='" & txtFrom.value & "'"
MainMenu.cr1.Formulas(1) = "tdate='" & txtto.value & "'"
MainMenu.cr1.WindowShowPrintSetupBtn = True

MainMenu.cr1.WindowShowPrintBtn = True
MainMenu.cr1.WindowShowExportBtn = True
MainMenu.cr1.WindowState = crptMaximized
MainMenu.cr1.Action = 1
Frame1.Visible = False

End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdprint__Click()

DSNNew

MainMenu.cr1.Reset
MainMenu.cr1.ReportFileName = rptPath & "/bookIssueTosampling.rpt"
MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass

Dim s_ As String

s_ = ""



If txtbcode.text <> "" Then
   s_ = "{BookIssueToSampling.bcode}='" & txtbcode.text & "'"
End If

If cbogp.text <> "ALL" Then

    If s_ = "" Then
       s_ = "{BookIssueToSampling.gp}='" & cbogp.text & "'"
    Else
       s_ = s_ & " and {BookIssueToSampling.gp}='" & cbogp.text & "'"
    End If

End If


If txtArea.text <> "" Then
If s_ = "" Then
   s_ = "{BookIssueToSampling.cityid}='" & txtAreacode.text & "'"
Else
     s_ = s_ & " and {BookIssueToSampling.cityid}='" & txtAreacode.text & "'"
End If
End If


If s_ <> "" Then
  MainMenu.cr1.ReplaceSelectionFormula "({BookIssueToSampling.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {BookIssueToSampling.invoicedate}<=datevalue('" & Format(txtto.value, "MM/dd/yyyy") & "')) and  " & s_ & " "
Else
  MainMenu.cr1.ReplaceSelectionFormula "({BookIssueToSampling.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {BookIssueToSampling.invoicedate}<=datevalue('" & Format(txtto.value, "MM/dd/yyyy") & "')) "
End If


MainMenu.cr1.Formulas(0) = "fdate='" & txtFrom.value & "'"
MainMenu.cr1.Formulas(1) = "tdate='" & txtto.value & "'"
MainMenu.cr1.WindowShowPrintSetupBtn = True

MainMenu.cr1.WindowShowPrintBtn = True
MainMenu.cr1.WindowShowExportBtn = True
MainMenu.cr1.WindowState = crptMaximized
MainMenu.cr1.Action = 1
Frame1.Visible = False

End Sub

Private Sub Command1_sh_Click()
   
   searchType = "inv11"
   popuplistFast "*", con, , , "bookissuetoschool1"

End Sub

Private Sub Command1_sh_GotFocus()

If PopUpValue1 <> "" Then

   txtschool.text = PopUpValue1
   txtSno.text = PopUpValue2
   
   txtDate.value = PopUpValue3
   
   txtScId.text = popupvalue5
   
   If RS.State = 1 Then RS.close
   RS.Open "select top 1 city,state,CBSECode from collegeView_ind where CollegeID='" & txtScId.text & "'", CON_blue
   If RS.EOF = False Then
      Commandedit.Enabled = True
      Commandsave.Enabled = False
      Commanddelete.Enabled = False
      
      
      txtcity = RS!city & ""
      txtcode = RS!cbsecode & ""
      txtstate = RS!State & ""
   End If
   
   vs.Clear
   setwidth_vs
   
   If RS.State = 1 Then RS.close
   RS.Open "select  a.bcode,b.bookname,a.qty,a.TeacherName,a.mobile," & _
   "a.email,a.TeachId,a.gp,a.repname from BookIssueToSchool as a inner join books as b " & _
   " on (a.bcode=b.bookcode) where (a.scID='" & txtScId.text & "' and a.sno=" & txtSno & ")", con
   
   For I = 1 To RS.RecordCount
   
   txtrepName.text = RS!RepName & ""
   
   vs.TextMatrix(I, 0) = I
   vs.TextMatrix(I, 1) = RS(0)
   vs.TextMatrix(I, 2) = RS(1)
   vs.TextMatrix(I, 3) = RS(2)
   
   If Not IsNull(RS!gp) Then
   vs.TextMatrix(I, 4) = RS!gp
   End If
   
   vs.TextMatrix(I, 5) = RS(3)
   vs.TextMatrix(I, 6) = RS(4)
   vs.TextMatrix(I, 7) = RS(5)
   vs.TextMatrix(I, 8) = RS(6)
   
   RS.MoveNext
   
   Next
   
   
   Total
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
End If

End Sub

Private Sub Commandadd_Click()
Screen.MousePointer = vbHourglass
  vs.Clear
  refreshFld
  
  txtschool.SetFocus
Screen.MousePointer = vbDefault
End Sub

Private Sub Commanddelete_Click()
If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
   con.Execute "delete from BookIssueToSchool where (sno=" & txtSno & " and scid='" & txtScId & "')"
   vs.Clear
   refreshFld
End If
End Sub

Private Sub Commandedit_Click()
Commandedit.Enabled = False
Commandsave.Enabled = True
Commanddelete.Enabled = True
End Sub

Private Sub CommandPrint_Click()

If (LCase(UserName) = "dc" Or LCase(UserName) = "admin" Or LCase(UserName) = "rishabh") Then
Frame1.Visible = True

End If
 

    
End Sub

Private Sub CommandReturn_Click()
Unload Me
End Sub
Sub refreshFld()
    txtSno = MaxSNo("BookIssueToSchool", "sno")
    setwidth_vs
    
    txtrepName.text = ""
    txtschool.text = ""
    txtScId.text = ""
    txtcode.text = ""
    txtstate.text = ""
    txtcity.text = ""
    txtTotal.text = ""
    
    txtArea.text = ""
    txtAreacode.text = ""
    
    
End Sub
Private Sub Commandsave_Click()


If RS.State = 1 Then RS.close
RS.Open "select * from BookIssueToSchool where (SCId='" & txtScId.text & "' and sno=" & txtSno & ")", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
   txtSno = MaxSNo("BookIssueToSchool", "sno")
End If



For I = 1 To vs.rows - 1

If (vs.TextMatrix(I, 1) <> "" And vs.TextMatrix(I, 8) <> "") Then

    If RS.State = 1 Then RS.close
    RS.Open "select * from BookIssueToSchool where (SCId='" & txtScId.text & "' and sno=" & txtSno & " and bcode='" & vs.TextMatrix(I, 1) & "' and  TeachId='" & vs.TextMatrix(I, 8) & "')", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = True Then
       'txtSno = MaxSNo("BookIssueToSchool", "sno")
       RS.AddNew
    End If
    
    RS!RepName = txtrepName.text
    RS!sno = txtSno.text
    RS!invoiceDate = txtDate.value
    RS!scid = txtScId.text
    RS!bcode = vs.TextMatrix(I, 1)
    RS!qty = vs.TextMatrix(I, 3)
    RS!gp = vs.TextMatrix(I, 4)
    
    If InStr(vs.TextMatrix(I, 5), "()") > 0 Then
       tname = Trim(Mid(vs.TextMatrix(I, 5), 1, InStr(vs.TextMatrix(I, 5), "()") - 1))
       RS!TeacherName = tname
    Else
       RS!TeacherName = vs.TextMatrix(I, 5)
    End If
    RS!mobile = vs.TextMatrix(I, 6)
    RS!email = vs.TextMatrix(I, 7)
    RS!TeachId = vs.TextMatrix(I, 8)
    RS.update
    
End If

Next


MsgBox "Data Save...", vbInformation
   
End Sub
Sub Total()
   Dim Tot As Long
   Tot = 0
   txtTotal.text = ""
   
   For J = 1 To vs.rows - 1
       If vs.TextMatrix(J, 3) <> "" Then
          Tot = Tot + IIf(vs.TextMatrix(J, 3) = "", 0, vs.TextMatrix(J, 3))
       End If
   Next
   
   txtTotal.text = Tot
   
End Sub
Private Sub Commandsearch_Click()
   searchType = "inv1"
   popuplistFast "*", con, , , "bookissuetoschool"
End Sub
Private Sub Commandsearch_GotFocus()


If PopUpValue1 <> "" Then

   txtschool.text = PopUpValue2
   txtSno.text = PopUpValue1
   
   txtDate.value = PopUpValue3
   
   txtScId.text = popupvalue5
   
   If RS.State = 1 Then RS.close
   RS.Open "select top 1 city,state,CBSECode from collegeView_ind where CollegeID='" & txtScId.text & "'", CON_blue
   If RS.EOF = False Then
      Commandedit.Enabled = True
      Commandsave.Enabled = False
      Commanddelete.Enabled = False
      
      
      txtcity = RS!city & ""
      txtcode = RS!cbsecode & ""
      txtstate = RS!State & ""
   End If
   
   vs.Clear
   setwidth_vs
   
   If RS.State = 1 Then RS.close
   RS.Open "select  a.bcode,b.bookname,a.qty,a.TeacherName,a.mobile," & _
   "a.email,a.TeachId,a.gp,a.repname from BookIssueToSchool as a inner join books as b " & _
   " on (a.bcode=b.bookcode) where (a.scID='" & txtScId.text & "' and a.sno=" & txtSno & ")", con
   
   For I = 1 To RS.RecordCount
   
   txtrepName.text = RS!RepName & ""
   
   vs.TextMatrix(I, 0) = I
   vs.TextMatrix(I, 1) = RS(0)
   vs.TextMatrix(I, 2) = RS(1)
   vs.TextMatrix(I, 3) = RS(2)
   
   If Not IsNull(RS!gp) Then
   vs.TextMatrix(I, 4) = RS!gp
   End If
   
   vs.TextMatrix(I, 5) = RS(3)
   vs.TextMatrix(I, 6) = RS(4)
   vs.TextMatrix(I, 7) = RS(5)
   vs.TextMatrix(I, 8) = RS(6)
   
   RS.MoveNext
   
   Next
   
   
   Total
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
End If



End Sub

Private Sub Form_Activate()
txtschool.SetFocus
End Sub
Private Sub Form_Load()



Me.Left = 0
Me.top = 0
Me.Width = 15700
Me.Height = 10000
txtDate.value = Format(Date, "dd/MM/yyyy")

txtFrom.value = from_date
txtto.value = to_date



refreshFld



'If (LCase(UserName) = "admin" Or LCase(UserName) = "dc") Then
   CommandPrint.Enabled = True
'Else
'   CommandPrint.Enabled = False
'End If


gp = ""

If RS.State = 1 Then RS.close
RS.Open "select groupcode from books group by groupcode", con
While RS.EOF = False

cbogp.AddItem RS(0)

If gp = "" Then
   gp = RS(0)
Else
   gp = gp & "|" & RS(0)
End If

RS.MoveNext
Wend


setwidth_vs

cbogp.AddItem "ALL"

cbogp.ListIndex = 2

BackColorFrom Me

End Sub
Sub setwidth_vs()
    
    vs.Cols = 6
    vs.FormatString = "S.No|B.Code|Book Name|>Qty.|Category|Teacher Name|Mobile|Email|TeachId"
    vs.ColWidth(0) = 500
    vs.ColWidth(1) = 1050
    vs.ColWidth(2) = 3400
    vs.ColWidth(3) = 600
    vs.ColWidth(4) = 900
    
    vs.ColWidth(5) = 2800
    vs.ColWidth(6) = 2100
    vs.ColWidth(7) = 2100
    vs.ColWidth(8) = 650
    
    vs.ColComboList(4) = gp
    
End Sub
Private Sub Form_Resize()
'panel.Left = (Me.ScaleWidth - panel.Width) / 2
'panel.Top = (Me.ScaleHeight - panel.Height) / 2
End Sub

Private Sub Text1_Change()

End Sub
Private Sub txtArea_GotFocus()

If PopUpValue1 <> "" Then

  txtArea.text = PopUpValue1
  txtAreacode.text = PopUpValue3
  
  PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""
  
End If


End Sub

Private Sub txtArea_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
 
    searchType = "party"
    popuplist_client "SELECT City,[State],[CityID] from CityView order by City", CON_blue
    
 
End If
End Sub

Private Sub txtBcode_GotFocus()

If PopUpValue1 <> "" Then

  txtbcode.text = PopUpValue1

  PopUpValue1 = ""
  PopUpValue2 = ""
End If

End Sub

Private Sub txtBcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then

   searchType = "books"
   popuplist10 "select BOOKCODE,BOOKNAME from BOOKS where " & stringyear & "  order by BOOKCODE", con

End If

End Sub

Private Sub txtcity_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtstate.SetFocus
End Sub

Private Sub txtcode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtcity.SetFocus
End Sub

Private Sub txtrepName_GotFocus()
If PopUpValue1 <> "" Then
   txtrepName.text = PopUpValue1
   txtcity.SetFocus
   PopUpValue1 = ""
End If
End Sub

Private Sub txtrepName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    popuplistModel10 "select Rep as Representative from SalesRepQry order by Rep", CON_blue
End If



End Sub

Private Sub txtschool_GotFocus()
If PopUpValue1 <> "" Then
    txtScId = PopUpValue1
    txtschool.text = PopUpValue2
    txtcity = PopUpValue3
    
    If RS.State = 1 Then RS.close
    RS.Open "select top 1 [state],cbcecode from TeacherDetails_ind where CollegeID='" & txtScId.text & "'", CON_blue
    If RS.EOF = False Then
       txtstate.text = RS(0)
       txtcode.text = RS(1)
    End If
    
    txtrepName.text = ""
    If RS.State = 1 Then RS.close
    RS.Open "select top 1 repname from collegeView where CollegeID='" & txtScId.text & "'", CON_blue
    If RS.EOF = False Then
       txtrepName.text = RS(0)
     End If
    
    
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    popupvalue4 = ""
End If

End Sub
Private Sub txtschool_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
   Screen.MousePointer = vbHourglass
   tblNo = 9
   frmSearchItem.Show
   Screen.MousePointer = vbDefault
End If

If KeyCode = 13 Then txtcode.SetFocus


End Sub



Private Sub txtSno_GotFocus()
Commandsearch_GotFocus
End Sub

Private Sub txtSno_KeyDown(KeyCode As Integer, Shift As Integer)
   
If KeyCode = 113 Then
   searchType = "inv1"
   popuplistFast "*", con, , , "bookissuetoschool"
End If


End Sub

Private Sub txtstate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then vs.SetFocus
End Sub

Private Sub vs_GotFocus()

If PopUpValue1 <> "" Then
   vs.TextMatrix(vs.RowSel, 5) = PopUpValue1
   vs.TextMatrix(vs.RowSel, 6) = PopUpValue2
   vs.TextMatrix(vs.RowSel, 7) = PopUpValue3
   vs.TextMatrix(vs.RowSel, 8) = popupvalue4
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
End If

End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
  If vs.Col = 5 Then
     If KeyCode = 13 Then
        If txtScId.text <> "" Then
           Screen.MousePointer = vbHourglass
           popuplist10 "Select Teacher,Mobile,Email,TeacherId from TeacherDetails_ind where CollegeID='" & txtScId.text & "' order by Teacher", CON_blue
           Screen.MousePointer = vbDefault
        End If
     End If
  End If
  
  
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If vs.Row >= 1 Then
           
           con.Execute "delete from BookIssueToSchool where (sno=" & txtSno & " and scid='" & txtScId & "' and BCODE='" & vs.TextMatrix(vs.RowSel, 1) & "')"
           vs.RemoveItem vs.Row
           vs.SetFocus
          End If
   End If
End If
  
  
End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   
If vs.Col = 1 Then
   If RS.State = 1 Then RS.close
   RS.Open "select bookname,bookcode,GROUPCODE,GROUPCODE_sub  from books where bookcode='" & vs.TextMatrix(vs.RowSel, 1) & "'", con
   If RS.EOF = False Then
      vs.TextMatrix(vs.RowSel, 0) = vs.Row
      vs.TextMatrix(vs.RowSel, 1) = RS(1)
      vs.TextMatrix(vs.RowSel, 2) = RS(0)
      
      If RS!groupcode = "BP" Then
      vs.TextMatrix(vs.RowSel, 4) = "K-12"
      Else
      vs.TextMatrix(vs.RowSel, 4) = RS!groupcode
      End If
      
      If Not IsNull(RS!GROUPCODE_sub) Then
         
         If Len(RS!GROUPCODE_sub) > 0 Then
            vs.TextMatrix(vs.RowSel, 4) = RS!GROUPCODE_sub
         End If
         
      End If
      
      
      sendkeys "{right}"
      sendkeys "{right}"
    End If
ElseIf (vs.Col = 3 Or vs.Col = 4 Or vs.Col <= 6) Then
    sendkeys "{right}"
ElseIf (vs.Col = 7) Then
    sendkeys "{home}"
    sendkeys "{down}"
    Total
End If

End If

End Sub

Private Sub vs_SelChange()
   If (vs.Col = 7 Or vs.Col = 5) Then
      vs.Editable = flexEDNone
   Else
      vs.Editable = flexEDKbdMouse
   End If
End Sub
