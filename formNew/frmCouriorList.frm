VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCouriorList 
   ClientHeight    =   9348
   ClientLeft      =   60
   ClientTop       =   708
   ClientWidth     =   15648
   Icon            =   "frmCouriorList.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9348
   ScaleWidth      =   15648
   Begin VB.CommandButton cmdAddPrice 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add Courier Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   13131
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   765
      Width           =   990
   End
   Begin VB.CommandButton Commandsearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   10143
      Picture         =   "frmCouriorList.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   765
      Width           =   990
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   13632
      TabIndex        =   28
      Top             =   8880
      Width           =   1140
   End
   Begin Crystal.CrystalReport cr 
      Left            =   6300
      Top             =   135
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CommandPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   9102
      Picture         =   "frmCouriorList.frx":0BF0
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   765
      Width           =   1035
   End
   Begin VB.CommandButton Commandedit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edit Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   12135
      Picture         =   "frmCouriorList.frx":17D4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   765
      Width           =   990
   End
   Begin VB.CommandButton cmdAdd_1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   7020
      Picture         =   "frmCouriorList.frx":1C16
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   765
      Width           =   1035
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   8061
      Picture         =   "frmCouriorList.frx":27FA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   765
      Width           =   1035
   End
   Begin VB.CommandButton Commandsave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sa&ve"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   11139
      Picture         =   "frmCouriorList.frx":33DE
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   765
      Width           =   990
   End
   Begin VB.CommandButton CommandReturn 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Return"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   14130
      Picture         =   "frmCouriorList.frx":3FC2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   765
      Width           =   900
   End
   Begin VB.ComboBox cboyear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmCouriorList.frx":4BA6
      Left            =   11610
      List            =   "frmCouriorList.frx":4BA8
      TabIndex        =   8
      Text            =   "cboyear"
      Top             =   180
      Width           =   3120
   End
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   360
   End
   Begin VB.ComboBox cboAName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      ItemData        =   "frmCouriorList.frx":4BAA
      Left            =   1350
      List            =   "frmCouriorList.frx":4BAC
      TabIndex        =   4
      Top             =   720
      Width           =   4920
   End
   Begin VB.ComboBox cboFirm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmCouriorList.frx":4BAE
      Left            =   1350
      List            =   "frmCouriorList.frx":4BB0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   4920
   End
   Begin MSComCtl2.DTPicker txtDate 
      Height          =   360
      Left            =   8280
      TabIndex        =   2
      Top             =   180
      Width           =   1545
      _ExtentX        =   2731
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   129105923
      CurrentDate     =   39979
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7188
      Left            =   0
      TabIndex        =   7
      Top             =   1668
      Width           =   15540
      _cx             =   27411
      _cy             =   12679
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   11162880
      BackColorFixed  =   7917545
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   8388608
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   520
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCouriorList.frx":4BB2
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
      Begin VB.Frame frmSearch 
         Caption         =   "Search ...."
         Height          =   1872
         Left            =   10035
         TabIndex        =   30
         Top             =   45
         Visible         =   0   'False
         Width           =   4515
         Begin VB.TextBox txtSvalue 
            Height          =   330
            Left            =   1620
            TabIndex        =   34
            Top             =   675
            Width           =   2805
         End
         Begin VB.ComboBox cbofld 
            Height          =   315
            ItemData        =   "frmCouriorList.frx":4CAE
            Left            =   1620
            List            =   "frmCouriorList.frx":4CBB
            TabIndex        =   33
            Top             =   315
            Width           =   2805
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Exit"
            Height          =   516
            Left            =   2835
            TabIndex        =   36
            Top             =   1215
            Width           =   1050
         End
         Begin VB.CommandButton cmdView_ 
            Caption         =   "&Search View"
            Height          =   510
            Left            =   1620
            TabIndex        =   35
            Top             =   1215
            Width           =   1140
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Search Field  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   90
            TabIndex        =   32
            Top             =   315
            Width           =   1680
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Serach Value :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   90
            TabIndex        =   31
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.Frame frmPrint 
         Height          =   1905
         Left            =   1035
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   2985
         Begin VB.CommandButton cmdPtint_ 
            Caption         =   "&Print"
            Height          =   555
            Left            =   1215
            TabIndex        =   27
            Top             =   1215
            Width           =   870
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Exit"
            Height          =   555
            Left            =   2115
            TabIndex        =   26
            Top             =   1215
            Width           =   825
         End
         Begin MSComCtl2.DTPicker fdate 
            Height          =   360
            Left            =   1395
            TabIndex        =   22
            Top             =   315
            Width           =   1545
            _ExtentX        =   2731
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   128843779
            CurrentDate     =   39979
         End
         Begin MSComCtl2.DTPicker tdate 
            Height          =   360
            Left            =   1395
            TabIndex        =   24
            Top             =   720
            Width           =   1545
            _ExtentX        =   2731
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   128843779
            CurrentDate     =   39979
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "To Date :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   90
            TabIndex        =   25
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "From Date :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   90
            TabIndex        =   23
            Top             =   315
            Width           =   1365
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1905
         Left            =   4005
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   6000
         Begin MSComDlg.CommonDialog cd 
            Left            =   3690
            Top             =   1215
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5535
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   270
            Width           =   405
         End
         Begin VB.TextBox txtPath 
            Height          =   330
            Left            =   45
            TabIndex        =   18
            Top             =   270
            Width           =   5505
         End
         Begin VB.CommandButton cmdExit_ 
            Caption         =   "Exit"
            Height          =   555
            Left            =   5130
            TabIndex        =   17
            Top             =   1170
            Width           =   825
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   555
            Left            =   4230
            TabIndex        =   16
            Top             =   1170
            Width           =   870
         End
      End
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      Height          =   288
      Left            =   13020
      TabIndex        =   29
      Top             =   8928
      Width           =   648
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Billing Month :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9900
      TabIndex        =   9
      Top             =   180
      Width           =   1725
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Agency"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      TabIndex        =   5
      Top             =   720
      Width           =   1410
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ent.Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7020
      TabIndex        =   3
      Top             =   180
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Firm Name "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   180
      Width           =   1365
   End
End
Attribute VB_Name = "frmCouriorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim searchType As String
Dim place_ As String
Dim con_chitra As ADODB.Connection

Private Sub cboFirm_LostFocus()

Dim db_ As String
Dim sqluser, sqlpass As String



If (cbofirm.text = "RAJLUXMI PUBLICATIONS") Then
    sqluser = "bpdatabase"
    sqlpass = "dinesh@123"
    db_ = "Database=RLData_" & Right(databaseNew, 4)
    serverNameNew_ = "192.168.0.140\BPSQLServer"
    
ElseIf (cbofirm.text = "BLUEPRINT EDUCATION") Then
    sqluser = "bpdatabase"
    sqlpass = "dinesh@123"
    db_ = "Database=chitradata_" & Right(databaseNew, 4)
    serverNameNew_ = "192.168.0.140\BPSQLServer"
    
ElseIf (cbofirm.text = "CHITRA PRAKASHAN (I) PVT.LTD.") Then
    sqluser = "chitradatabase"
    sqlpass = "java.123"
    db_ = "Database=chitraDNet_" & Right(databaseNew, 4)
    serverNameNew_ = "WIN-FI4EQR95VL3\CHITRASQLserver"

End If




Set con_chitra = New ADODB.Connection
    

    



con_chitra.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverNameNew_ & "; " & db_ & "; UID=" & sqluser & "; PWD=" & sqlpass
    
DoEvents
DoEvents
    
con_chitra.CursorLocation = adUseClient
If con_chitra.State = 1 Then con_chitra.close
con_chitra.Open


End Sub

Private Sub cmdAdd_1_Click()
 Commandsave.Enabled = True
 vs.Clear
 vs.rows = 2
 setWidth
 cmdAdd_1.Enabled = True
 
End Sub

Private Sub cmdAddPrice_Click()
frmPriceMasterCourier.Show
End Sub

Private Sub cmdExit__Click()
Frame1.Visible = False
End Sub
Private Sub cmdPtint__Click()

CR.Reset
CR.ReportFileName = rptPath & "/couriorChargesList.rpt"
CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
If (cbofirm.text <> "" And cboAName.text <> "") Then
   CR.ReplaceSelectionFormula "({CourierCharge.firmname}='" & cbofirm.text & "' and {CourierCharge.aname}='" & cboAName.text & "') and ({CourierCharge.dates}>=datevalue('" & Format(fdate.value, "MM/dd/yy") & "') and {CourierCharge.dates}<=datevalue('" & Format(tdate.value, "MM/dd/yy") & "'))"
End If
ss_ = fdate.value & " TO " & tdate.value

CR.Formulas(0) = "GTOTAL='" & "BILLING AMOUNT (" & ss_ & ") : " & "'"

CR.WindowShowPrintSetupBtn = True
CR.WindowState = crptMaximized
CR.Action = 1

End Sub

Private Sub cmdSave_Click()

Screen.MousePointer = vbHourglass

Dim con_ch As New ADODB.Connection

con_ch.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + txtpath

con_ch.Open

con.Execute "delete from CouriorMaster"

con.Execute "insert into CouriorMaster(Name,FirmName) SELECT distinct Party,'BLUEPRINT EDUCATION' as Type_ FROM SLEDGER where gledger='SUNDRY DEBTORS'"
con.Execute "insert into CouriorMaster(Name,FirmName) SELECT distinct AgentName,'BLUEPRINT EDUCATION' as Type_ FROM INVOICEA where AgentName not in (select name from CouriorMaster)"

If rs1.State = 1 Then rs1.close
rs1.Open "SELECT distinct Party,'CHITRA PRAKASHAN (I) PVT.LTD.' as Type_ FROM SLEDGER where gledger='SUNDRY DEBTORS'", con_ch, adOpenDynamic, adLockOptimistic
While rs1.EOF = False
    con.Execute "insert into CouriorMaster(Name,FirmName) values('" & rs1(0) & "','" & rs1(1) & "')"
    rs1.MoveNext
Wend

If rs1.State = 1 Then rs1.close
rs1.Open "SELECT distinct AGENTNAME FROM AGENTMASTER", con_ch, adOpenDynamic, adLockOptimistic
While rs1.EOF = False
    con.Execute "insert into CouriorMaster(Name,FirmName) values('" & rs1(0) & "','CHITRA PRAKASHAN (I) PVT.LTD.')"
    rs1.MoveNext
Wend

con.Execute "insert into CouriorMaster(Name,FirmName) SELECT distinct NameOfParty,FirmName FROM CourierCharge where FirmName='CHITRA PRAKASHAN (I) PVT.LTD.' and NameOfParty  not in (select name from CouriorMaster where FirmName='CHITRA PRAKASHAN (I) PVT.LTD.')"
con.Execute "insert into CouriorMaster(Name,FirmName) SELECT distinct NameOfParty,FirmName FROM CourierCharge where FirmName='BLUEPRINT EDUCATION' and NameOfParty  not in (select name from CouriorMaster where FirmName='BLUEPRINT EDUCATION')"

Screen.MousePointer = vbDefault

MsgBox "updated record...", vbInformation

End Sub

Private Sub cmdView__Click()
searchType = "2"
searchData
End Sub

Private Sub cmdView_Click()
searchType = "1"
searchData

'cmdAdd_1.Enabled = False
End Sub

Private Sub Command1_Click()
cd.ShowOpen
txtpath = cd.filename
cmdSave.Enabled = True
End Sub

Private Sub Command13_Click()
HeadTbl = "aname"
frmMasters.Show 1
End Sub
Sub setWidth()
    
    vs.FormatString = "SNo|Description of Contents||Name Of Party/Recient Address|Delivery Station/City|State|ConsignmentNo|Courier Service|Weight|Unit|Rate|Freight Charge"
    vs.ColWidth(0) = 500
    vs.ColWidth(1) = 1800
    vs.ColWidth(2) = 0
    vs.ColWidth(3) = 3800
    
    vs.ColWidth(4) = 1900
    vs.ColWidth(5) = 800
    
    vs.ColWidth(6) = 1500   '5 change
    vs.ColWidth(7) = 1500
    
    vs.ColWidth(8) = 800
    vs.ColWidth(9) = 700
    vs.ColWidth(10) = 800
    vs.ColWidth(11) = 1000
    
    vs.WordWrap = True
    
    
    vs.Cell(flexcpFontSize, 0, 0) = 9
    vs.Cell(flexcpFontSize, 0, 1) = 9
    vs.Cell(flexcpFontSize, 0, 2) = 9
    vs.Cell(flexcpFontSize, 0, 3) = 9
    vs.Cell(flexcpFontSize, 0, 4) = 9
    vs.Cell(flexcpFontSize, 0, 5) = 9
    vs.Cell(flexcpFontSize, 0, 6) = 9
    vs.Cell(flexcpFontSize, 0, 7) = 9
    vs.Cell(flexcpFontSize, 0, 8) = 9
    vs.Cell(flexcpFontSize, 0, 9) = 9
    vs.Cell(flexcpFontSize, 0, 10) = 9
    
    
End Sub
Sub fillcombo()
  
If RS.State = 1 Then RS.close
cboAName.Clear
RS.Open "select  name from MasterTbl where category='aname'", con
While RS.EOF = False
  cboAName.AddItem RS(0)
  RS.MoveNext
Wend
   

'============================

If RS.State = 1 Then RS.close
cboyear.Clear
RS.Open "select  monthname, year_ from yearmonth_ order by id", con
While RS.EOF = False
  cboyear.AddItem RS!MonthName & "-" & RS!year_
  RS.MoveNext
Wend

   
   
End Sub

Private Sub Command2_Click()
frmPrint.Visible = False
End Sub

Private Sub Command4_Click()
frmSearch.Visible = False
End Sub

Private Sub Commandedit_Click()
Frame1.Visible = True
cmdSave.Enabled = False
End Sub

Private Sub CommandPrint_Click()
frmPrint.Visible = True
End Sub

Private Sub CommandReturn_Click()
Unload Me
End Sub
Private Sub Commandsave_Click()


If cbofirm.text = "" Then
   MsgBox "Select Firm Name...", vbCritical
   cbofirm.SetFocus
   Exit Sub
End If

If cboAName.text = "" Then
   MsgBox "Select Agent Name...", vbCritical
   cboAName.SetFocus
   Exit Sub
End If

If cboyear.text = "" Then
   MsgBox "Select Month Name...", vbCritical
   cboyear.SetFocus
   Exit Sub
End If

Screen.MousePointer = vbHourglass

On Error GoTo aa:

Set RS = New ADODB.Recordset
For K = 1 To vs.rows - 1


If vs.TextMatrix(K, 0) <> "" Then
    
    
    If RS.State = 1 Then RS.close
    RS.Open "select * from CourierCharge where (convert(datetime,Dates,103)=convert(datetime,'" & txtDate.value & "',103) and firmname='" & cbofirm.text & "' and sn='" & vs.TextMatrix(K, 0) & "' and aname='" & cboAName.text & "')", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = True Then
       RS.AddNew
    End If
       
    RS!dates = txtDate.value
    RS!firmname = cbofirm.text
    RS!aname = cboAName.text
    RS!sn = vs.TextMatrix(K, 0)
    RS!Month_ = cboyear.text
    RS!DESCRIPTION = Trim(vs.TextMatrix(K, 1))
    RS!RefNo = vs.TextMatrix(K, 2)
    RS!NameOfParty = vs.TextMatrix(K, 3)
    RS!station = Trim(vs.TextMatrix(K, 4))
    RS!states = Trim(vs.TextMatrix(K, 5))
    RS!docno = vs.TextMatrix(K, 6)
    RS!cname = Trim(vs.TextMatrix(K, 7))
    RS!weight = IIf(vs.TextMatrix(K, 8) = "", 0, vs.TextMatrix(K, 8))
    RS!unit = vs.TextMatrix(K, 9)
    RS!rate = IIf(vs.TextMatrix(K, 10) = "", "", vs.TextMatrix(K, 10))
    RS!FCharges = IIf(vs.TextMatrix(K, 11) = "", 0, vs.TextMatrix(K, 11))
    RS.update

End If


Next

MsgBox "Data Saved...", vbQuestion + vbYesNo

Screen.MousePointer = vbDefault

Exit Sub
aa:
Screen.MousePointer = vbDefault
MsgBox "" & err.DESCRIPTION
   
End Sub
Function checkDocNo(docno As String) As Boolean

Set rs1 = New ADODB.Recordset
Set rs1 = con.Execute("exec searchList '" & "docno" & "'")
rs1.MoveFirst
rs1.Find "docno='" & docno & "'"
'rs1.Open "select count(*) from CourierCharge where (DocNo='" & docno & "')"
If rs1.EOF = False Then
   MsgBox docno & " : Doc.No Alreay Exist in this record. " & vbCrLf & "SN.  : " & rs1!sn & vbCrLf & "Date : " & rs1!dates & vbCrLf & "FirmName : " & rs1!firmname & vbCrLf & "Month : " & rs1!Month_ & vbCrLf & "Name Of Recipeint : " & rs1!NameOfParty & vbCrLf & "Name Of AgentName : " & rs1!aname
   checkDocNo = True
Else
   checkDocNo = False
End If

End Function
Sub searchData()

Screen.MousePointer = vbHourglass

On Error GoTo aa:

Dim str_ As String
Dim frt As Double
Set RS = New ADODB.Recordset

vs.Clear
vs.rows = 2
frt = 0
str_ = ""

If searchType = "1" Then
    str_ = "(convert(datetime,Dates,103)=convert(datetime,'" & txtDate.value & "',103))"
    If cbofirm.text <> "" Then str_ = str_ & " and firmname='" & cbofirm.text & "'"
    'If cboFirm.Text <> "" Then str_ = str_ & " and firmname='" & cboFirm.Text & "'"
    If cboAName.text <> "" Then str_ = str_ & " and aname='" & cboAName.text & "'"
    If cboyear.text <> "" Then str_ = str_ & " and month_='" & cboyear.text & "'"
Else
    If cbofirm.text <> "" Then str_ = "firmname='" & cbofirm.text & "'"
    If cboAName.text <> "" Then str_ = str_ & " and aname='" & cboAName.text & "'"
    
    str_ = str_ & " and " & cbofld.text & " like '" & txtSvalue & "%'"
    
End If




If RS.State = 1 Then RS.close
RS.Open "select * from CourierCharge where " & str_, con, adOpenDynamic, adLockOptimistic



For K = 1 To RS.RecordCount
If RS.EOF = False Then
    
    Commandsave.Enabled = True
    
    vs.rows = vs.rows + 1
    txtDate.value = RS!dates
    cbofirm.text = RS!firmname
    cboAName.text = RS!aname
    vs.TextMatrix(K, 0) = RS!sn
    cboyear.text = RS!Month_
    vs.TextMatrix(K, 1) = RS!DESCRIPTION
    vs.TextMatrix(K, 2) = RS!RefNo
    vs.TextMatrix(K, 3) = RS!NameOfParty
    vs.TextMatrix(K, 4) = RS!station
    vs.TextMatrix(K, 5) = RS!states & ""
    vs.TextMatrix(K, 6) = RS!docno
    vs.TextMatrix(K, 7) = RS!cname
    vs.TextMatrix(K, 8) = RS!weight
    vs.TextMatrix(K, 9) = RS!unit & ""
    vs.TextMatrix(K, 10) = RS!rate & ""
    vs.TextMatrix(K, 11) = RS!FCharges
    
    If (Not IsNull(RS!FCharges) Or RS!FCharges <> "") Then
       frt = frt + RS!FCharges
    End If
    
    '''--------------
    If (vs.TextMatrix(K, 1) <> "" And vs.TextMatrix(K, 3) <> "" And vs.TextMatrix(K, 6) <> "" And vs.TextMatrix(K, 7) <> "" And Len(Trim(vs.TextMatrix(K, 8))) >= 1 And vs.TextMatrix(K, 9) <> "" And vs.TextMatrix(K, 10) <> "") Then
     For I = 0 To 11
     vs.Cell(flexcpBackColor, K, I) = vbGreen
     Next
    End If
     

End If

RS.MoveNext

Next

txtTotal.text = 0

txtTotal.text = frt

setWidth

Screen.MousePointer = vbDefault

Exit Sub
aa:
Screen.MousePointer = vbDefault
MsgBox "" & err.DESCRIPTION

End Sub

Private Sub Commandsearch_Click()
frmSearch.Visible = True
End Sub

Private Sub Form_Load()

Me.top = 10
Me.Left = 10
Me.Height = 9840
Me.Width = 15480

BackColorFrom Me

cbofirm.AddItem "BLUEPRINT EDUCATION"
cbofirm.AddItem "CHITRA PRAKASHAN (I) PVT.LTD."
cbofirm.AddItem "RAJLUXMI PUBLICATIONS"

txtDate.value = Format(Date, "dd/MM/yyyy")

'ssss = MonthName(Month(Date)) & "-" & Year(Date)
fdate.value = Format(Date, "dd/MM/yyyy")
tdate.value = Format(Date, "dd/MM/yyyy")


fillcombo
setWidth

place_ = ""

If RS.State = 1 Then RS.close
RS.Open "select  name from MasterTbl where category='cplace' order by name", con
While RS.EOF = False
    If place_ = "" Then
       place_ = RS(0)
    Else
       place_ = place_ & "|" & RS(0)
    End If
    RS.MoveNext
Wend

vs.ColComboList(5) = place_

cboyear.text = MonthName(Month(Date)) & "-" & Year(Date)

End Sub

Private Sub vs_GotFocus()

 
 
  If vs.Col = 3 Then
 
    If PopUpValue1 <> "" Then
       vs.TextMatrix(vs.RowSel, 3) = PopUpValue1
    End If
    
    If PopUpValue1 <> "" Then
       vs.TextMatrix(vs.RowSel, 4) = PopUpValue2
    End If

    
    If PopUpValue1 <> "" Then
       vs.TextMatrix(vs.RowSel, 5) = PopUpValue3
    End If

    
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""

 
 ElseIf vs.Col = 7 Then
 
    If PopUpValue1 <> "" Then
       vs.TextMatrix(vs.RowSel, 7) = PopUpValue1
    End If
    
    'SendKeys "{right}"
    PopUpValue1 = ""
 ElseIf vs.Col = 3 Then
    If PopUpValue1 <> "" Then
     s1 = InStr(PopUpValue1, ",")
     If s1 > 0 Then
        vs.TextMatrix(vs.RowSel, 3) = Mid(PopUpValue1, 1, s1 - 1)
        vs.TextMatrix(vs.RowSel, 4) = Mid(PopUpValue1, s1 + 1)
     Else
        vs.TextMatrix(vs.RowSel, 3) = PopUpValue1
     End If
    End If
    'SendKeys "{right}"
    PopUpValue1 = ""
 
 ElseIf vs.Col = 4 Then
    
    If PopUpValue1 <> "" Then
     vs.TextMatrix(vs.RowSel, 4) = PopUpValue1
    End If
    
    'SendKeys "{right}"
    PopUpValue1 = ""

 End If
End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 112 Then
  
  
   If vs.Col = 3 Then
    
     popuplistFast "SELECT distinct NameOfParty,Station,states FROM CourierCharge order by NameOfParty", con, , , "courier"

  
  ElseIf vs.Col = 7 Then
     'value = "select Transportname from [transportmaster] order by Transportname"
     value = "select  name from MasterTbl where category='aname'"
     popuplistModel10 value, con
  ElseIf vs.Col = 4 Then
     value = "select Placeofsupply from TransportDet group by Placeofsupply order by Placeofsupply"
     popuplistModel10 value, con
  ElseIf vs.Col = 3 Then
     
     If cbofirm.text <> "" Then
        'value = "select Party from SLEDGER where gledger ='SUNDRY DEBTORS' order by Party"
        searchType = "cmaster"
       If cbofirm.text = "CHITRA PRAKASHAN (I) PVT.LTD." Then
        value = "select name from CouriorMaster order by name"
        popuplistFast value, con, , , "cmaster_ch"
       Else
        value = "select name from CouriorMaster order by name"
        popuplistFast value, con, , , "cmaster_bp"
       
       End If
     End If
     
     
  End If
  
  ElseIf KeyCode = 115 Then
  
    str_ = "(convert(datetime,Dates,103)=convert(datetime,'" & txtDate.value & "',103))"
    If cbofirm.text <> "" Then
    str_ = str_ & " and firmname='" & cbofirm.text & "'"
    End If

    If cbofirm.text <> "" Then
    str_ = str_ & " and firmname='" & cbofirm.text & "'"
    End If

    If cboAName.text <> "" Then
    str_ = str_ & " and aname='" & cboAName.text & "'"
    End If

    If cboyear.text <> "" Then
    str_ = str_ & " and month_='" & cboyear.text & "'"
    End If

    If MsgBox("Want to Delete ?", vbQuestion + vbYesNo) = vbYes Then
    
    If vs.TextMatrix(vs.RowSel, 0) <> "" Then
        con.Execute "delete from CourierCharge where sn='" & vs.TextMatrix(vs.RowSel, 0) & "' and " & str_
        vs.RemoveItem (vs.RowSel)
    End If
    
    End If
  
  'ElseIf KeyCode = 113 Then
   
  End If
  
  
  
End Sub
Function returnState(st1 As String) As String

Dim s4 As String

If (st1 = "ANDHRA PRADESH") Then
 s4 = "AP"
ElseIf (st1 = "BIHAR") Then
 s4 = "BH"
ElseIf (st1 = "DELHI") Then
s4 = "DLH"
ElseIf (st1 = "HARYANA") Then
s4 = "HR"
ElseIf (st1 = "MADHYA PRADESH") Then
s4 = "MP"
ElseIf (st1 = "PUNJAB") Then
s4 = "PJB"
ElseIf (st1 = "RAJASTHAN") Then
s4 = "RJ"
ElseIf (st1 = "UTTAR PRADESH") Then
s4 = "UP"
ElseIf (st1 = "UTTARAKHAND") Then
 s4 = "UK"
ElseIf (st1 = "WEST BENGAL") Then

 s4 = "WB"
ElseIf (st1 = "HIMACHAL PRADESH") Then
 s4 = "HP"
End If

returnState = s4
End Function

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
  
  On Error Resume Next
  
  If KeyCode = 13 Then
     
     
    If vs.Col = 1 Then
    
    
    Dim ss10, vtype, vno
    
    ss10 = InStr(vs.TextMatrix(vs.RowSel, 1), "-")
    
    
    If (ss10 > 0) Then
    
    
    vtype = UCase(Mid(vs.TextMatrix(vs.RowSel, 1), 1, ss10 - 1))
    vno = UCase(Mid(vs.TextMatrix(vs.RowSel, 1), ss10 + 1))
    
    vs.TextMatrix(vs.RowSel, 1) = UCase(vs.TextMatrix(vs.RowSel, 1))
    
    
        If (cbofirm.text = "RAJLUXMI PUBLICATIONS" Or cbofirm.text = "BLUEPRINT EDUCATION") Then
            If (vtype = "I") Then
                If rs1.State = 1 Then rs1.close
                rs1.Open "select party,city,states from invoiceaQry where invoiceno=" & vno & "", con_chitra
                
                If (rs1.EOF = False) Then
                
                vs.TextMatrix(vs.RowSel, 3) = rs1(0)
                vs.TextMatrix(vs.RowSel, 4) = rs1(1)
                hh = returnState(rs1(2))
                vs.TextMatrix(vs.RowSel, 5) = hh
                
                End If
                
            ElseIf (vtype = "CI") Then
            
                If rs1.State = 1 Then rs1.close
                rs1.Open "select party,city,states from CREDITAQry where invoiceno=" & vno & "", con_chitra
                
                If (rs1.EOF = False) Then
                    vs.TextMatrix(vs.RowSel, 3) = rs1(0)
                    vs.TextMatrix(vs.RowSel, 4) = rs1(1)
                    hh = returnState(rs1(2))
                    vs.TextMatrix(vs.RowSel, 5) = hh
                End If
    
           
            ElseIf (vtype = "CN") Then
            
                If rs1.State = 1 Then rs1.close
                rs1.Open "SELECT [DESCFORINVOICE],[ADDRESS3],states FROM SLEDGER where SUBLEDGER in (select SUBLEDGER from CreditNotReg where cnn=" & vno & ")", con_chitra
                If (rs1.EOF = False) Then
                    vs.TextMatrix(vs.RowSel, 3) = rs1(0)
                    vs.TextMatrix(vs.RowSel, 4) = rs1(1)
                    hh = returnState(rs1(2))
                    vs.TextMatrix(vs.RowSel, 5) = hh
                
                End If
    
            
            
            ElseIf (vtype = "DN") Then
            
                If rs1.State = 1 Then rs1.close
                rs1.Open "SELECT [DESCFORINVOICE],[ADDRESS3],states FROM SLEDGER where SUBLEDGER in (select SUBLEDGER from debitRegister where dnn=" & vno & ")", con_chitra
                If (rs1.EOF = False) Then
                    vs.TextMatrix(vs.RowSel, 3) = rs1(0)
                    vs.TextMatrix(vs.RowSel, 4) = rs1(1)
                    hh = returnState(rs1(2))
                    vs.TextMatrix(vs.RowSel, 5) = hh
                End If
                
            ElseIf (vtype = "SP") Then
            
                If rs1.State = 1 Then rs1.close
                rs1.Open "SELECT AgentName,city,states,Shipto,Shipto_City,Shipto_States FROM INVOICEA_sp where invoiceno=" & vno & "", con_chitra
                If (rs1.EOF = False) Then
                   
                   If IsNull(rs1!Shipto) Then
                   
                     If RS.State = 1 Then RS.close
                     RS.Open "SELECT DESCFORINVOICE FROM SLEDGER where SUBLEDGER ='" & rs1!agentname & "'", con_chitra
                     If RS.EOF = False Then
                        vs.TextMatrix(vs.RowSel, 3) = RS!DESCFORINVOICE
                     Else
                        vs.TextMatrix(vs.RowSel, 3) = rs1(0)
                     End If
                   
                     
                     vs.TextMatrix(vs.RowSel, 4) = rs1(1)
                     hh = returnState(rs1(2))
                     vs.TextMatrix(vs.RowSel, 5) = hh

                   Else
                     vs.TextMatrix(vs.RowSel, 3) = rs1!Shipto
                     vs.TextMatrix(vs.RowSel, 4) = rs1!Shipto_City
                     hh = returnState(rs1!Shipto_States)
                     vs.TextMatrix(vs.RowSel, 5) = hh
                   End If
                   
                End If
   
    
            
            End If
            
       ElseIf (cbofirm.text = "CHITRA PRAKASHAN (I) PVT.LTD.") Then
       
       
            If (vtype = "I") Then
                If rs1.State = 1 Then rs1.close
                rs1.Open "select DESCFORINVOICE as party,city,states from SaleARegister where invoiceno=" & vno & "", con_chitra
                
                If (rs1.EOF = False) Then
                
                vs.TextMatrix(vs.RowSel, 3) = rs1(0)
                vs.TextMatrix(vs.RowSel, 4) = rs1(1)
                hh = returnState(rs1(2))
                vs.TextMatrix(vs.RowSel, 5) = hh
                
                End If
                
            ElseIf (vtype = "CI") Then
            
                If rs1.State = 1 Then rs1.close
                rs1.Open "select DESCFORINVOICE as party,city,states from CreditARegister where invoiceno=" & vno & "", con_chitra
                
                If (rs1.EOF = False) Then
                    vs.TextMatrix(vs.RowSel, 3) = rs1(0)
                    vs.TextMatrix(vs.RowSel, 4) = rs1(1)
                    hh = returnState(rs1(2))
                    vs.TextMatrix(vs.RowSel, 5) = hh
                End If
    
           
            ElseIf (vtype = "CN") Then
            
                If rs1.State = 1 Then rs1.close
                rs1.Open "select Subledger,city,states from CreditNoteARegister where invoiceno=" & vno & "", con_chitra
                If (rs1.EOF = False) Then
                    vs.TextMatrix(vs.RowSel, 3) = rs1(0)
                    vs.TextMatrix(vs.RowSel, 4) = rs1(1)
                    hh = returnState(rs1(2))
                    vs.TextMatrix(vs.RowSel, 5) = hh
                
                End If
    
            
            
            ElseIf (vtype = "DN") Then
            
                If rs1.State = 1 Then rs1.close
                rs1.Open "select Subledger,city,states from DebitNoteARegister where invoiceno=" & vno & "", con_chitra
                If (rs1.EOF = False) Then
                    vs.TextMatrix(vs.RowSel, 3) = rs1(0)
                    vs.TextMatrix(vs.RowSel, 4) = rs1(1)
                    hh = returnState(rs1(2))
                    vs.TextMatrix(vs.RowSel, 5) = hh
                End If
    
           
            ElseIf (vtype = "SP") Then
            
                If rs1.State = 1 Then rs1.close
                rs1.Open "select AgentName from Con_BookTransToAgnReg_A where invoiceno=" & vno & "", con_chitra
                If (rs1.EOF = False) Then
                    vs.TextMatrix(vs.RowSel, 3) = rs1(0)
                    'vs.TextMatrix(vs.RowSel, 4) = rs1(1)
                    'hh = returnState(rs1(2))
                    'vs.TextMatrix(vs.RowSel, 5) = hh
                End If

            
            End If
       
       
       
       
      
       End If
   
   End If

    
 
    
    
    ''Set rs1 = New ADODB.Recordset
    ''con_chitra
    
     
    End If
    
     
     
     
    If vs.Col = 5 Then
       If vs.TextMatrix(vs.RowSel, 5) <> "" Then
          Set RS = New ADODB.Recordset
          RS.Open "SELECT SN,Dates,FirmName,Month_,NameOfParty,aname FROM CourierCharge where DocNo='" & vs.TextMatrix(vs.RowSel, 5) & "'", con
          If RS.EOF = False Then
             MsgBox vs.TextMatrix(vs.RowSel, 5) & " : Doc.No Alreay Exist in this record. " & vbCrLf & "SN.  : " & RS!sn & vbCrLf & "Date : " & RS!dates & vbCrLf & "FirmName : " & RS!firmname & vbCrLf & "Month : " & RS!Month_ & vbCrLf & "Name Of Recipeint : " & RS!NameOfParty & vbCrLf & "Name Of AgentName : " & RS!aname
             vs.SetFocus
             
             GoTo aa:
          End If
          
          
       End If
    End If
     
     rate_ = 0
     
     If vs.Col <= 9 Then
     
        If (vs.Col = 9) Then
           If rs1.State = 1 Then rs1.close
           If (vs.TextMatrix(vs.RowSel, 5) <> "" And vs.TextMatrix(vs.RowSel, 7) <> "") Then
           rs1.Open "select ChargePerMin,ChargePerKG from CourierPriceMaster " & _
           " where (PlaceOfSupp='" & Trim(vs.TextMatrix(vs.RowSel, 5)) & "' and CourierMaster='" & vs.TextMatrix(vs.RowSel, 7) & "' and AgencyName='" & cboAName.text & "')", con
           If rs1.EOF = False Then
               If vs.TextMatrix(vs.RowSel, 9) = "Min" Then
                 vs.TextMatrix(vs.RowSel, 10) = rs1(0)
                 rate_ = rs1(0)
                 vs.TextMatrix(vs.RowSel, 11) = rate_
              Else
                 vs.TextMatrix(vs.RowSel, 10) = rs1(1)
                 rate_ = rs1(1)
                 
                 
                 wt = 0
                 wt1 = 0
                 
                 wt = IIf(vs.TextMatrix(vs.RowSel, 8) = "", 0, vs.TextMatrix(vs.RowSel, 8))
                 wt = wt - Int(wt)
                 If (wt > 0) Then
                 
                   wt1 = IIf(vs.TextMatrix(vs.RowSel, 8) = "", 0, vs.TextMatrix(vs.RowSel, 8))
                   wt = wt1 - wt
                   
                   wt = wt + 1
                 Else
                   wt = IIf(vs.TextMatrix(vs.RowSel, 8) = "", 0, vs.TextMatrix(vs.RowSel, 8))
                 End If
                 
                 
                 vs.TextMatrix(vs.RowSel, 11) = (rate_ * wt)
              End If
              
              'vs.TextMatrix(vs.RowSel, 10) = (rate_ * vs.TextMatrix(vs.RowSel, 8))
           End If
           End If
        ElseIf vs.Col = 9 Then
        
           
        
        End If
     
        sendkeys "{right}"
        vs.TextMatrix(vs.RowSel, 0) = vs.Row
     Else
     
     vs.rows = vs.rows + 1
     If vs.Row > 1 Then
        'vs.TextMatrix(vs.RowSel + 1, 5) = vs.TextMatrix(vs.RowSel, 5)
     End If
     
     
     sendkeys "{home}"
     sendkeys "{down}"
     End If
     
  End If
  
aa:
End Sub
Private Sub vs_SelChange()
''If (vs.Col = 3 Or vs.Col = 4 Or vs.Col = 6) Then
''   vs.Editable = flexEDNone
''Else
''   vs.Editable = flexEDKbdMouse
''End If
End Sub
