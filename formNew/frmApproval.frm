VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmApproval 
   Caption         =   "Approval Form"
   ClientHeight    =   9732
   ClientLeft      =   60
   ClientTop       =   396
   ClientWidth     =   15108
   Icon            =   "frmApproval.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9732
   ScaleWidth      =   15108
   Begin VB.Frame Frame2 
      Height          =   552
      Left            =   4140
      TabIndex        =   55
      Top             =   1296
      Visible         =   0   'False
      Width           =   3324
      Begin VB.CommandButton cmdexcel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Excel"
         Height          =   396
         Left            =   2556
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   108
         Width           =   696
      End
      Begin VB.ComboBox cbogp 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         ItemData        =   "frmApproval.frx":000C
         Left            =   36
         List            =   "frmApproval.frx":000E
         TabIndex        =   56
         Top             =   144
         Width           =   2496
      End
   End
   Begin VB.CheckBox Check1_gpofSchool 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Group Of School"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   4176
      TabIndex        =   54
      Top             =   972
      Width           =   2136
   End
   Begin VB.CommandButton cmdSchWisePartywise 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Party && School Wise App. Details"
      Height          =   720
      Left            =   8892
      Picture         =   "frmApproval.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   8424
      Width           =   1392
   End
   Begin VB.CommandButton Check2_AppDet1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print Total Approval List"
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   1440
      Width           =   2520
   End
   Begin VSFlex7Ctl.VSFlexGrid VS_AppDet 
      Height          =   2730
      Left            =   90
      TabIndex        =   51
      Top             =   2835
      Visible         =   0   'False
      Width           =   14790
      _cx             =   26088
      _cy             =   4815
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   12640511
      ForeColor       =   11162880
      BackColorFixed  =   12640511
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   -2147483647
      BackColorBkg    =   16777215
      BackColorAlternate=   12640511
      GridColor       =   -2147483633
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmApproval.frx":0BF4
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
   Begin VB.CheckBox Check2_AppDet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Approval List"
      Height          =   345
      Left            =   996
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   1452
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton cmdFindSchoolNameChange 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Find Changed SchoolName In Inv."
      Height          =   735
      Left            =   10296
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   8415
      Width           =   1104
   End
   Begin VB.TextBox txtAuthorised 
      Enabled         =   0   'False
      Height          =   285
      Left            =   13995
      TabIndex        =   47
      Top             =   684
      Width           =   825
   End
   Begin VB.TextBox txtFAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   13488
      TabIndex        =   44
      Top             =   9000
      Width           =   1080
   End
   Begin VB.CommandButton cmdModify 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Modify"
      Enabled         =   0   'False
      Height          =   720
      Left            =   3204
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8415
      Width           =   924
   End
   Begin VB.CheckBox Check1_noApp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "No Approval"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   2616
      TabIndex        =   42
      Top             =   960
      Width           =   1344
   End
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Party Remarks"
      Height          =   780
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   1035
      Width           =   1476
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   15192
      TabIndex        =   37
      Top             =   8910
      Width           =   105
   End
   Begin VB.TextBox txtNet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   13488
      TabIndex        =   35
      Top             =   8640
      Width           =   1080
   End
   Begin VB.TextBox txtGross 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   13488
      TabIndex        =   34
      Top             =   8280
      Width           =   1080
   End
   Begin VB.CheckBox Check1_manually 
      BackColor       =   &H8000000E&
      Caption         =   "Add App. No Manually"
      Height          =   255
      Left            =   8550
      TabIndex        =   31
      Top             =   180
      Width           =   2010
   End
   Begin VB.CommandButton cmdPendingApp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pending &Approval"
      Height          =   720
      Left            =   7968
      Picture         =   "frmApproval.frx":0CF4
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8415
      Width           =   924
   End
   Begin Crystal.CrystalReport cr 
      Left            =   945
      Top             =   9270
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CommandPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   720
      Left            =   7008
      Picture         =   "frmApproval.frx":18D8
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8415
      Width           =   960
   End
   Begin VB.CommandButton Commandsearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "S&earch"
      Height          =   720
      Left            =   6036
      Picture         =   "frmApproval.frx":24BC
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8415
      Width           =   960
   End
   Begin VB.CommandButton Commandedit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   720
      Left            =   5076
      Picture         =   "frmApproval.frx":30A0
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8415
      Width           =   960
   End
   Begin VB.ComboBox cboColName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmApproval.frx":34E2
      Left            =   10800
      List            =   "frmApproval.frx":34F2
      TabIndex        =   5
      Top             =   1080
      Width           =   1815
   End
   Begin VB.ComboBox cboValue 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmApproval.frx":351C
      Left            =   10800
      List            =   "frmApproval.frx":352C
      TabIndex        =   6
      Top             =   1485
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddSer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&ADD Series Name For Upto 5 %"
      Height          =   465
      Left            =   10710
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   45
      Width           =   2445
   End
   Begin VB.CheckBox Check1_upto5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Up to 5 (%)"
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   1128
   End
   Begin VB.CommandButton Commanddelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "De&lete"
      Enabled         =   0   'False
      Height          =   720
      Left            =   4116
      Picture         =   "frmApproval.frx":3545
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8415
      Width           =   960
   End
   Begin VB.ComboBox cboyesno 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      ItemData        =   "frmApproval.frx":4129
      Left            =   10800
      List            =   "frmApproval.frx":412B
      TabIndex        =   4
      Top             =   600
      Width           =   1815
   End
   Begin VB.CheckBox Check1_schoolAll 
      BackColor       =   &H8000000E&
      Caption         =   "Select All School"
      Height          =   255
      Left            =   8235
      TabIndex        =   14
      Top             =   1170
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.CommandButton cmdAdd_1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1284
      Picture         =   "frmApproval.frx":412D
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8415
      Width           =   960
   End
   Begin VB.CommandButton cmdSave_2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   2244
      Picture         =   "frmApproval.frx":4D11
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8415
      Width           =   960
   End
   Begin VB.CommandButton cmdExit_12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   756
      Left            =   11400
      Picture         =   "frmApproval.frx":58F5
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8415
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   1395
      TabIndex        =   13
      Top             =   45
      Width           =   2964
      Begin VB.OptionButton Option2_Party 
         Caption         =   "Party"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1710
         TabIndex        =   1
         Top             =   180
         Width           =   1410
      End
      Begin VB.OptionButton Option1_school 
         Caption         =   "School"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   0
         Top             =   135
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.TextBox txtscid 
      Enabled         =   0   'False
      Height          =   315
      Left            =   7464
      TabIndex        =   11
      Top             =   600
      Width           =   1488
   End
   Begin VB.TextBox txtSchoolName 
      Height          =   315
      Left            =   1395
      MaxLength       =   150
      TabIndex        =   2
      Top             =   600
      Width           =   6024
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3720
      Left            =   48
      TabIndex        =   7
      Top             =   1932
      Width           =   14880
      _cx             =   26247
      _cy             =   6562
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmApproval.frx":64D9
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
      Begin VB.Frame frmCheck 
         Height          =   1050
         Left            =   45
         TabIndex        =   39
         Top             =   -45
         Visible         =   0   'False
         Width           =   10590
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "E&xit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   9945
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   180
            Width           =   570
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Exit"
            Height          =   510
            Left            =   13815
            TabIndex        =   41
            Top             =   135
            Width           =   465
         End
         Begin VSFlex7Ctl.VSFlexGrid vs2 
            Height          =   2040
            Left            =   45
            TabIndex        =   40
            Top             =   135
            Width           =   9870
            _cx             =   17410
            _cy             =   3598
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
            BackColorFixed  =   -2147483633
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
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   400
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
      End
   End
   Begin MSMask.MaskEdBox txtAppDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   7470
      TabIndex        =   24
      Top             =   135
      Width           =   1050
      _ExtentX        =   1842
      _ExtentY        =   572
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtAppNo 
      Height          =   315
      Left            =   5535
      TabIndex        =   25
      Top             =   135
      Width           =   1035
      _ExtentX        =   1842
      _ExtentY        =   572
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VSFlex7Ctl.VSFlexGrid vs1 
      Height          =   2505
      Left            =   45
      TabIndex        =   36
      Top             =   5715
      Width           =   14880
      _cx             =   26247
      _cy             =   4419
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmApproval.frx":65E1
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
   Begin VSFlex7Ctl.VSFlexGrid vs_gp 
      Height          =   168
      Left            =   6696
      TabIndex        =   58
      Top             =   1044
      Visible         =   0   'False
      Width           =   696
      _cx             =   1228
      _cy             =   296
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
      BackColorFixed  =   -2147483633
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   400
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Authorised :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   372
      Left            =   12732
      TabIndex        =   48
      Top             =   684
      Width           =   1272
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Final Amt. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   8
      Left            =   12456
      TabIndex        =   45
      Top             =   9000
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Amt :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   7
      Left            =   12480
      TabIndex        =   33
      Top             =   8328
      Width           =   1068
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amt. :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   6
      Left            =   12480
      TabIndex        =   32
      Top             =   8688
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "App. Date :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   6615
      TabIndex        =   27
      Top             =   135
      Width           =   1110
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "App. No :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   4770
      TabIndex        =   26
      Top             =   180
      Width           =   1110
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Press Enter Button..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   12690
      TabIndex        =   22
      Top             =   1530
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Column Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   9135
      TabIndex        =   21
      Top             =   1125
      Width           =   1740
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Column Value :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   9135
      TabIndex        =   20
      Top             =   1530
      Width           =   1740
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F4 Key To Delete rows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1530
      TabIndex        =   19
      Top             =   9450
      Width           =   2955
   End
   Begin VB.Label lblrow 
      Height          =   240
      Left            =   24
      TabIndex        =   16
      Top             =   8412
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fill Series Wise/All :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   0
      Left            =   9132
      TabIndex        =   15
      Top             =   648
      Width           =   1740
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "School Name / Party Name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   468
      Index           =   0
      Left            =   132
      TabIndex        =   12
      Top             =   600
      Width           =   1476
   End
End
Attribute VB_Name = "frmApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dt_str As String
Dim fdate1, tdate2 As String
Dim Edit As Boolean
Dim party_ As String
Dim GROSS, net As Double
Dim search_New As Boolean
Dim CON_next As ADODB.Connection

Private Sub cboValue_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   
   
For K = 4 To vs.Cols - 1
       If vs.TextMatrix(0, K) = cboColName.text Then
          If cboyesno.text = "ALL" Then
            For J = 1 To vs.rows - 1
                  vs.TextMatrix(J, K) = cboValue.text
            Next
          Else
            For J = 1 To vs.rows - 1
                 If vs.TextMatrix(J, 2) = cboyesno.text Then
                     vs.TextMatrix(J, K) = cboValue.text
                 End If
            Next
          End If
       End If
Next
   
   
   
End If

End Sub

Private Sub Check1_gpofSchool_Click()

If Check1_gpofSchool.value = 1 Then
    Frame2.Visible = True
Else
    Frame2.Visible = False
End If

End Sub
Sub fillGP_school()

Dim rss As New ADODB.Recordset
Set rss = New ADODB.Recordset

rss.Open "SELECT Name From  MasterTbl where Category='groupofschool'", CON_blue
While rss.EOF = False
    cbogp.AddItem rss(0)
rss.MoveNext
Wend

End Sub
Private Sub Check1_manually_Click()
If Check1_manually.value = 1 Then
   txtappno.Enabled = True
   txtAppDate.Enabled = True
Else
   txtappno.Enabled = False
   txtAppDate.Enabled = False
End If

End Sub

Private Sub Check1_upto5_Click()
   

   
If (search_New = False) Then
   
If Check1_upto5.value = 1 Then
   
 For J = 1 To vs.rows - 1
 If vs.TextMatrix(J, 2) <> "" Then
   If rs1.State = 1 Then rs1.close
   rs1.Open "select  name from  MasterTbl where (name='" & vs.TextMatrix(J, 2) & "' and category='SerName')", con
   If rs1.EOF = False Then
      vs.TextMatrix(J, 4) = 5
   End If
 End If
Next

Else

 For J = 1 To vs.rows - 1
 If vs.TextMatrix(J, 1) <> "" Then
      vs.TextMatrix(J, 4) = ""
 End If
Next

End If

End If
   
   
End Sub
Sub appDateAppDet()
  
Dim ff As New ADODB.Recordset
Dim rssave As New ADODB.Recordset

Dim tdis, dis, finalAmt
tdis = 0
dis = 0
finalAmt = 0

If rssave.State = 1 Then rssave.close
rssave.Open "select * from ApprovalDet", con, adOpenDynamic, adLockOptimistic



If rs1.State = 1 Then rs1.close
rs1.Open "select SerName,appno,ID as scid,(AppPer+adj+discount+promo) as tdis,Net_Gross from AppForm group by SerName,appno,id,AppPer,adj,discount,promo,net_gross", con, adOpenDynamic, adLockOptimistic
While rs1.EOF = False

If ff.State = 1 Then ff.close
ff.Open "select appno from ApprovalDet where appno='" & rs1!appno & "'", con, adOpenDynamic, adLockOptimistic
If ff.EOF = True Then

Set RS = New ADODB.Recordset
RS.Open "select invoiceNo,InvoiceDate,subledger,Scid,SerName,Fyear,sum(grossAmt),sum(NetAmt),appNo from useForApprovalQry where (SerName='" & rs1!sername & "' and Scid='" & rs1!scid & "' and " & dt_str & ") group by invoiceNo,InvoiceDate,subledger,Scid,SerName,Fyear,AppNo order by AppNo", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then

        rssave.AddNew
        rssave!invoiceNo = RS!invoiceNo
        rssave!invoiceDate = RS!invoiceDate
        rssave!subledger = RS!subledger
        rssave!scid = RS!scid
        rssave!sername = RS!sername
        rssave!gamount = RS(6)
        rssave!netamount = RS(7)
        rssave!fyear = RS!fyear
        rssave!appno = rs1!appno
        
        rssave!tdis = rs1(3)
        
        If rs1!Net_Gross = "Gross" Then
           dis_ = Round(Val(RS(6) * Val(tdis) / 100), 0)
           rssave!TdisAmt = RS(6) - dis_
        ElseIf rs1!Net_Gross = "Net" Then
           dis_ = Round(Val(RS(7) * Val(tdis) / 100), 0)
           rssave!TdisAmt = RS(7) - dis_
        End If
        
        rssave!addData = "add"
        rssave.update
   
End If
   
End If
   
 rs1.MoveNext
Wend
   
End Sub

Private Sub Check2_AppDet_Click()


''Dim rs_f As ADODB.Recordset
''Set rs_f = New ADODB.Recordset
''
''Screen.MousePointer = vbHourglass
''
''If Option2_Party.value = True Then
''    rs_f.Open "SELECT appNo,scid,scname,SUBLEDGER,Promo,Adj,Discount,AppPer,(Promo+Adj+Discount+AppPer) as TDis,Net_Gross,SerName " & _
''    "from PartyRemarksQryNew where substring(SUBLEDGER,1,5)='" & txtScId.Text & "'  order by appNo,scid", con
''Else
''    rs_f.Open "SELECT appNo,scid,scname,SUBLEDGER,Promo,Adj,Discount,AppPer,(Promo+Adj+Discount+AppPer) as TDis,Net_Gross,SerName " & _
''    "from PartyRemarksQryNew where scid='" & txtScId.Text & "'  order by appNo,scid", con
''
''End If
''
''Set VS_AppDet.DataSource = rs_f
''
''VS_AppDet.FormatString = "AppNo|ScId|ScName|Party|Promo|Adj|Discount|AppPer|TDis|Net_Gross|SerName"
''
''VS_AppDet.ColWidth(0) = 700
''VS_AppDet.ColWidth(1) = 800
''VS_AppDet.ColWidth(2) = 3000
''VS_AppDet.ColWidth(3) = 3500
''VS_AppDet.ColWidth(4) = 700
''VS_AppDet.ColWidth(5) = 700
''VS_AppDet.ColWidth(6) = 700
''VS_AppDet.ColWidth(7) = 700
''VS_AppDet.ColWidth(8) = 700
''VS_AppDet.ColWidth(9) = 1000
''VS_AppDet.ColWidth(10) = 1000
''
''VS_AppDet.Editable = flexEDNone
''
''
''Screen.MousePointer = vbDefault
'''======================================
''
''If Check2_AppDet.value = 1 Then
''   VS_AppDet.Visible = True
''Else
''   VS_AppDet.Visible = False
''End If


End Sub

Private Sub Check2_AppDet1_Click()

Dim pterm As String
pterm = ""

If rs1.State = 1 Then rs1.close
rs1.Open "select distinct id,code,School_Party from AppForm where (code='" & txtScId.text & "' or id='" & txtScId.text & "')"
If rs1.EOF = False Then
   If rs1!School_Party = "School" Then
      code_ = rs1!Code
   Else
      code_ = rs1!id
   End If

    If RS.State = 1 Then RS.close
    RS.Open "SELECT PartyRemarks FROM SLEDGER where code = '" & code_ & "'", con
    If RS.EOF = False Then
       con.Execute "update AppForm set pterms='" & RS(0) & "' where (id='" & code_ & "' or code='" & code_ & "')"
    End If

End If



con.Execute "update a set a.repname=b.agentname from AppForm as a " & _
"inner join agentQry_update_toAppFrom as b on (a.SerName=b.SerName and a.code=b.scid and a.fyear=b.fyear)"

con.Execute "update a set a.repname=b.agentname from AppForm as a " & _
"inner join agentQry_update_toAppFrom as b on (a.SerName=b.SerName and a.id=b.scid and a.fyear=b.fyear)"



Set RS = New ADODB.Recordset
RS.Open "select fromDate,toDate,NotCreated,DataBase from turnOverDis where Current_Next='next'", CCON
If RS.EOF = False Then
  If RS!NotCreated = "y" Then
     con.Execute "update a set a.repname=b.agentname from AppForm as a " & _
     "inner join [" + RS!Database + "].[dbo].[agentQry_update_toAppFrom] as b on (a.SerName=b.SerName and a.code=b.scid and a.fyear=b.fyear)"
     
     con.Execute "update a set a.repname=b.agentname from AppForm as a " & _
     "inner join [" + RS!Database + "].[dbo].[agentQry_update_toAppFrom] as b on (a.SerName=b.SerName and a.id=b.scid and a.fyear=b.fyear)"

  End If
End If



DSNNew

CR.Reset
CR.ReportFileName = rptPath & "/ApprovalDet.rpt"

'If Option2_Party.value = True Then
 CR.ReplaceSelectionFormula "({AppForm.id}='" & txtScId.text & "' OR {AppForm.CODE}='" & txtScId.text & "')"
'Else
   'cr.ReplaceSelectionFormula "{AppForm.id}='" & txtScId.Text & "'"
'   Exit Sub
'End If

CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
'cr.Formulas(0) = "pterms_='" & pterm & "'"
CR.Formulas(0) = "fyear='" & session & "'"
CR.WindowShowPrintSetupBtn = True
CR.WindowShowPrintBtn = True
CR.WindowShowExportBtn = True
CR.WindowState = crptMaximized
CR.Action = 1


End Sub

Private Sub cmdAdd_1_Click()
      
'appDateAppDet

search_New = False
      
cmdPendingApp.Enabled = True
CommandPrint.Enabled = True
cmdFindSchoolNameChange.Enabled = True
cmdAdd_1.Enabled = True
Commandsearch.Enabled = True
         
txtAppDate.text = Format(Date, "dd/MM/yyyy")
Edit = False
vs.Clear
setWidth
txtSchoolName.text = ""
txtAuthorised = ""
txtScId.text = ""
Check1_upto5.value = 0
maxId
txtappno.Enabled = False
vs.Enabled = True
Commandedit.Enabled = False
Commanddelete.Enabled = False
cmdSave_2.Enabled = True
cmdModify.Enabled = False

txtFAmt.text = 0
txtNet.text = 0
txtGross.text = 0
txtSchoolName.SetFocus
Check1_noApp.value = 0
vs1.Clear
vs1.Enabled = True
      
      
End Sub
Sub maxId()
    If rs1.State = 1 Then rs1.close
    rs1.Open "select max(appno) from AppForm", con
    If IsNull(rs1(0)) Then
       txtappno = 1
    Else
       txtappno.text = rs1(0) + 1
    End If
    
End Sub
Private Sub cmdAddSer_Click()
HeadTbl = "SerName"
frmMasters.Show 1

End Sub
Private Sub cmdCheck_Click()
party_ = "view"
printPRemarks
End Sub
Sub printPRemarks()

Dim f As New ADODB.Recordset
Set f = New ADODB.Recordset
st10 = ""

con.Execute "delete from AppPrintTmp"

If (txtScId.text <> "" And txtSchoolName.text <> "") Then
If Option1_school.value = True Then
   str1 = "SELECT distinct a.SUBLEDGER,b.PartyRemarks  FROM ApprovalDet as a inner join sledger as b on (a.SUBLEDGER =b.SUBLEDGER) where a.Scid ='" & txtScId.text & "'"
Else
   party_ = txtScId.text & " " & txtSchoolName.text
   str1 = "SELECT distinct a.SUBLEDGER,b.PartyRemarks  FROM ApprovalDet as a inner join sledger as b on (a.SUBLEDGER =b.SUBLEDGER) where a.SUBLEDGER ='" & party_ & "'"
End If


If f.State = 1 Then f.close
f.Open str1, con

While f.EOF = False
   
If f(1) <> "NA" Then
   con.Execute "insert into AppPrintTmp(party,Remarks) values('" & f(0) & "','" & f(1) & "')"
End If
f.MoveNext

Wend


 

If party_ = "view" Then
 Set vs2.DataSource = f
 frmCheck.Visible = True
 Command3.Enabled = True
 frmCheck.Enabled = True
 vs.Enabled = True
End If


End If

End Sub
Sub findChangedSchoolName()

Dim rss As New ADODB.Recordset
Dim f As New ADODB.Recordset
Set f = New ADODB.Recordset
st10 = ""

Set rss = New ADODB.Recordset



str1 = "SELECT INVOICENO,INVOICEDATE,ScName,ScID,AppNo FROM useForApprovalQry where app_add='y'"

If f.State = 1 Then f.close
f.Open str1, con
While f.EOF = False
   
   
   

f.MoveNext
Wend



End Sub

Private Sub cmdExcel_Click()
Dim rs_ As ADODB.Recordset



Set rs_ = New ADODB.Recordset
rs_.Open "SELECT CollegeID,School,Add1 +Add2 as Adress,City,District,[State] FROM collegeView where GroupOfSchool='" & cbogp.text & "'", CON_blue
Set vs_gp.DataSource = rs_


Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim str_ As String




If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r As Long
Dim Q1, q2, J As Double



row_ = 1
col_ = 1

xl.Columns("A:H").ColumnWidth = 12
J = 2


For I = 0 To vs_gp.rows - 1
    For J = 0 To vs_gp.Cols - 1
      
        xlSheet.Cells(row_, col_).value = vs_gp.TextMatrix(I, J)
        
        If (col_ = 2) Then
            If rs1.State = 1 Then rs1.close
            rs1.Open "SELECT AppNO FROM AppForm where (id='" & vs_gp.TextMatrix(I, J) & "' or code='" & vs_gp.TextMatrix(I, J) & "')", con
            If rs1.EOF = False Then
                xlSheet.Cells(row_, 1).value = rs1(0)
            End If
        End If
        
        col_ = col_ + 1
    Next
    row_ = row_ + 1
    col_ = 1
Next

MsgBox "Task Completed....", vbInformation

End Sub

Private Sub cmdExit_12_Click()
  Unload Me
End Sub
Private Sub cmdModify_Click()

If RS.State = 1 Then RS.close
RS.Open "select fromDate,toDate,NotCreated from turnOverDis where fyear='" & session & "'", CCON
If RS.EOF = False Then
  If RS!NotCreated = "y" Then
   fdate1 = RS!fromdate
   tdate2 = RS!todate
  End If
End If

'====================================================================

For I = 1 To vs.rows - 1
If vs.TextMatrix(I, 0) <> "" Then
If vs.TextMatrix(I, 11) = "0" Then
  If RS.State = 1 Then RS.close
  RS.Open "select * from AppForm where (Appno='" & txtappno.text & "' and SerName ='" & vs.TextMatrix(I, 2) & "' and id='" & txtScId.text & "' and  School_PartyName='" & txtSchoolName.text & "' and code='" & vs.TextMatrix(I, 0) & "')", con, adOpenDynamic, adLockOptimistic
  If RS.EOF = True Then
    RS.AddNew
  End If
  
    If Check1_noApp.value = 1 Then
       RS!noapp = "y"
    Else
       RS!noapp = "n"
    End If
    
    RS!id = txtScId.text
    RS!School_PartyName = txtSchoolName.text
    
    RS!appno = txtappno.text
    RS!appdate = txtAppDate.text
    'RS!bAuthorized = 0
    
    RS!GrossAmt = Val(txtGross.text)
    RS!netamt = Val(txtNet.text)
    If Option1_school.value = True Then
    RS!School_Party = "School"
    Else
    RS!School_Party = "Party"
    End If
    
    RS!Code = vs.TextMatrix(I, 0)
    RS!pname = vs.TextMatrix(I, 1)
    RS!sername = vs.TextMatrix(I, 2)
    RS!discount = vs.TextMatrix(I, 3)
    
    'If (vs.TextMatrix(I, 4) <> "" And vs.TextMatrix(I, 4) <> "0") Then
    'RS!AppPer = vs.TextMatrix(I, 4)
    'Else
    'RS!AppPer = 0
    'End If
    
    
    RS!appper = IIf(vs.TextMatrix(I, 4) = "", 0, vs.TextMatrix(I, 4))
    RS!adj = IIf(vs.TextMatrix(I, 5) = "", 0, vs.TextMatrix(I, 5))
    RS!Promo = IIf(vs.TextMatrix(I, 6) = "", 0, vs.TextMatrix(I, 6))
    RS!Net_Gross = IIf(vs.TextMatrix(I, 7) = "", "", vs.TextMatrix(I, 7))
    
    RS!tod = vs.TextMatrix(I, 8)
    RS!cd = vs.TextMatrix(I, 9)
    RS!remarks = vs.TextMatrix(I, 10)
    RS!fyear = vs.TextMatrix(I, 12)
    RS!userid = UId
    
    RS!updatedBy = UserName
    
    
    
    RS.update
End If
End If
Next

'======================================================================================================================
con.Execute "update AppForm set GrossAmt=" & txtGross.text & ",NetAmt=" & txtNet.text & ",FinalAmt=" & txtFAmt & " where appno=" & txtappno.text & ""
'======================================================================================================================

Dim rs_1 As New ADODB.Recordset

For I = 1 To vs1.rows - 1
If vs1.TextMatrix(I, 0) <> "" Then

Set RS = New ADODB.Recordset
RS.Open "select * from ApprovalDet where (appno='" & txtappno.text & "' and invoiceNo = '" & vs1.TextMatrix(I, 0) & "' and SerName='" & vs1.TextMatrix(I, 4) & "')", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
    RS.AddNew
Else
    If (vs1.TextMatrix(I, 8) = txtappno.text) Then
        RS!gamount = vs1.TextMatrix(I, 5)
        RS!netamount = vs1.TextMatrix(I, 6)
        RS!TdisAmt = vs1.TextMatrix(I, 10)
        RS.update
    End If
End If
    
If (vs1.TextMatrix(I, 8) = "" Or vs1.TextMatrix(I, 8) = txtappno.text) Then
    RS!invoiceNo = vs1.TextMatrix(I, 0)
    RS!invoiceDate = vs1.TextMatrix(I, 1)
    RS!subledger = vs1.TextMatrix(I, 2)
    RS!scid = vs1.TextMatrix(I, 3)
    RS!sername = vs1.TextMatrix(I, 4)
    RS!gamount = vs1.TextMatrix(I, 5)
    RS!netamount = vs1.TextMatrix(I, 6)
    RS!fyear = vs1.TextMatrix(I, 7)
    RS!appno = Trim(txtappno.text)
    RS!tdis = IIf(vs1.TextMatrix(I, 9) = "", 0, vs1.TextMatrix(I, 9))
    RS!TdisAmt = IIf(vs1.TextMatrix(I, 10) = "", 0, vs1.TextMatrix(I, 10))
    RS.update
End If

End If
Next

'===================================================================
   
If Option1_school.value = True Then
   
    For K = 1 To vs1.rows - 1
    If vs1.TextMatrix(K, 8) = "" Then
       
       If rs1.State = 1 Then rs1.close
       rs1.Open "select bookcode,fyear from updateInvoiceB_fromAppTbl where id='" & txtScId.text & "' and sername='" & vs1.TextMatrix(K, 4) & "'", con
       While rs1.EOF = False
       
       If vs1.TextMatrix(K, 7) = session Then
         con.Execute "update invoiceB set appno='" & txtappno.text & "',app_add='y' where (invoiceNo='" & vs1.TextMatrix(K, 0) & "' and bookcode='" & rs1!Bookcode & "')"
         con.Execute " update invoicea set App_Add='y',appno='" & txtappno & "'  where INVOICENO = " & vs1.TextMatrix(K, 0) & ""
       Else
         CON_next.Execute "update invoiceB set appno='" & txtappno.text & "',app_add='y' where (invoiceNo='" & vs1.TextMatrix(K, 0) & "' and bookcode='" & rs1!Bookcode & "')"
         CON_next.Execute " update invoicea set App_Add='y',appno='" & txtappno & "'  where INVOICENO = " & vs1.TextMatrix(K, 0) & ""
       
       End If
         
         rs1.MoveNext
       Wend
       
    End If
    Next
   
Else

    For K = 1 To vs1.rows - 1
    If vs1.TextMatrix(K, 8) = "" Then
       If rs1.State = 1 Then rs1.close
       rs1.Open "select bookcode from updateInvoiceB_fromAppTbl where id='" & txtScId.text & "' and sername='" & vs1.TextMatrix(K, 4) & "'", con
       While rs1.EOF = False
       
           If vs1.TextMatrix(K, 7) = session Then
            con.Execute "update invoiceB set appno='" & txtappno & "',app_add='y' where (invoiceNo='" & vs1.TextMatrix(K, 0) & "' and bookcode='" & rs1!Bookcode & "')"
            con.Execute " update invoicea set App_Add='y',appno='" & txtappno & "'  where INVOICENO = " & vs1.TextMatrix(K, 0) & ""
           Else
             CON_next.Execute "update invoiceB set appno='" & txtappno.text & "',app_add='y' where (invoiceNo='" & vs1.TextMatrix(K, 0) & "' and bookcode='" & rs1!Bookcode & "')"
             CON_next.Execute " update invoicea set App_Add='y',appno='" & txtappno & "'  where INVOICENO = " & vs1.TextMatrix(K, 0) & ""
           
           End If
       
       rs1.MoveNext
       Wend
       
    End If
    Next

 
End If


MsgBox " Updated.....", vbInformation
Commandedit.Enabled = True
cmdSave_2.Enabled = False
cmdModify.Enabled = False
Commanddelete.Enabled = False

End Sub

Private Sub cmdPendingApp_Click()
frmPendingApp.Show 1
End Sub
Private Sub cmdSave_2_Click()

If RS.State = 1 Then RS.close
RS.Open "select fromDate,toDate,NotCreated from turnOverDis where fyear='" & session & "'", CCON
If RS.EOF = False Then
  If RS!NotCreated = "y" Then
   fdate1 = RS!fromdate
   tdate2 = RS!todate
  End If
End If

If rs1.State = 1 Then rs1.close
rs1.Open "select top 1 * from AppForm where appno=" & txtappno.text & "", con, adOpenStatic, adLockReadOnly
If rs1.EOF = False Then
    If rs1!bAuthorized = True Then
        MsgBox "You can'nt change, this Approval No Locked !!", vbExclamation, "Alert"
        Exit Sub
    End If
End If




If Edit = False Then
If Check1_manually.value = 0 Then
   maxId
End If
End If

If vs.rows = 1 Then
   MsgBox "Data not found ...", vbCritical
   Exit Sub
End If



If RS.State = 1 Then RS.close
RS.Open "select * from AppForm where appno='" & txtappno.text & "'", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
con.Execute "delete from AppForm where appno='" & txtappno.text & "'"
End If

For I = 1 To vs.rows - 1

If vs.TextMatrix(I, 2) <> "" Then

    RS.AddNew
    If Check1_noApp.value = 1 Then
       RS!noapp = "y"
    Else
       RS!noapp = "n"
    End If
    
    RS!id = txtScId.text
    RS!School_PartyName = txtSchoolName.text
    
    RS!appno = txtappno.text
    RS!appdate = txtAppDate.text
    RS!bAuthorized = 0
    
    RS!GrossAmt = Val(txtGross.text)
    RS!netamt = Val(txtNet.text)
    RS!finalAmt = Val(txtFAmt.text)
    
    
    If Option1_school.value = True Then
    RS!School_Party = "School"
    Else
    RS!School_Party = "Party"
    End If
    
    
    
    RS!Code = vs.TextMatrix(I, 0)
    RS!pname = vs.TextMatrix(I, 1)
    RS!sername = vs.TextMatrix(I, 2)
    RS!discount = vs.TextMatrix(I, 3)
    
    If Val(vs.TextMatrix(I, 4)) > 0 Then
    RS!appper = vs.TextMatrix(I, 4)
    Else
    RS!appper = 0
    End If
    
    RS!adj = IIf(vs.TextMatrix(I, 5) = "", 0, vs.TextMatrix(I, 5))
    RS!Promo = IIf(vs.TextMatrix(I, 6) = "", 0, vs.TextMatrix(I, 6))
    RS!Net_Gross = IIf(vs.TextMatrix(I, 7) = "", 0, vs.TextMatrix(I, 7))
    RS!tod = vs.TextMatrix(I, 8)
    RS!cd = vs.TextMatrix(I, 9)
    RS!remarks = vs.TextMatrix(I, 10)
    RS!fyear = vs.TextMatrix(I, 12)
    RS!userid = UId
    
    RS!updatedBy = UserName
    
    RS.update

End If
Next

'======================================================================
Dim rs_1 As New ADODB.Recordset

'If RS.State = 1 Then RS.close
'RS.Open "select * from ApprovalDet where appno='" & txtAppNo.Text & "'", con, adOpenDynamic, adLockOptimistic
'If RS.EOF = False Then
   'con.Execute "delete from ApprovalDet where appno='" & txtAppNo.Text & "'"
'End If

For I = 1 To vs1.rows - 1
If vs1.TextMatrix(I, 0) <> "" Then
    
Set RS = New ADODB.Recordset
RS.Open "select * from ApprovalDet where (invoiceNo ='" & vs1.TextMatrix(I, 0) & "' and appno='" & txtappno.text & "' and SerName='" & vs1.TextMatrix(I, 4) & "')", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
   If vs1.TextMatrix(I, 8) = "" Then
        RS.AddNew
        RS!invoiceNo = vs1.TextMatrix(I, 0)
        RS!invoiceDate = vs1.TextMatrix(I, 1)
        RS!subledger = vs1.TextMatrix(I, 2)
        RS!scid = vs1.TextMatrix(I, 3)
        RS!sername = vs1.TextMatrix(I, 4)
        RS!gamount = vs1.TextMatrix(I, 5)
        RS!netamount = vs1.TextMatrix(I, 6)
        RS!fyear = vs1.TextMatrix(I, 7)
        RS!appno = Trim(txtappno.text)
        RS!tdis = IIf(vs1.TextMatrix(I, 9) = "", 0, vs1.TextMatrix(I, 9))
        RS!TdisAmt = IIf(vs1.TextMatrix(I, 10) = "", 0, vs1.TextMatrix(I, 10))
        RS.update
    End If

Else
        RS!invoiceNo = vs1.TextMatrix(I, 0)
        RS!invoiceDate = vs1.TextMatrix(I, 1)
        RS!subledger = vs1.TextMatrix(I, 2)
        RS!scid = vs1.TextMatrix(I, 3)
        RS!sername = vs1.TextMatrix(I, 4)
        RS!gamount = vs1.TextMatrix(I, 5)
        RS!netamount = vs1.TextMatrix(I, 6)
        RS!fyear = vs1.TextMatrix(I, 7)
        RS!appno = Trim(txtappno.text)
        RS!tdis = IIf(vs1.TextMatrix(I, 9) = "", 0, vs1.TextMatrix(I, 9))
        RS!TdisAmt = IIf(vs1.TextMatrix(I, 10) = "", 0, vs1.TextMatrix(I, 10))
        RS.update
End If
        
    
End If
Next


'===================================================================
If Option1_school.value = True Then
    
    For K = 1 To vs1.rows - 1
    If (vs1.TextMatrix(K, 0) <> "" And vs1.TextMatrix(K, 8) = "") Then
       If rs1.State = 1 Then rs1.close
       rs1.Open "select bookcode,fyear from updateInvoiceB_fromAppTbl where id='" & txtScId.text & "' and sername='" & vs1.TextMatrix(K, 4) & "'", con
       While rs1.EOF = False
         
       If rs1!fyear = session Then
         con.Execute "update invoiceB set appno='" & txtappno.text & "',app_add='y' where (invoiceNo='" & vs1.TextMatrix(K, 0) & "' and bookcode='" & rs1!Bookcode & "')"
         con.Execute " update invoicea set App_Add='y',appno='" & txtappno & "'  where INVOICENO = " & vs1.TextMatrix(K, 0) & ""
       Else
         CON_next.Execute "update invoiceB set appno='" & txtappno.text & "',app_add='y' where (invoiceNo='" & vs1.TextMatrix(K, 0) & "' and bookcode='" & rs1!Bookcode & "')"
         CON_next.Execute " update invoicea set App_Add='y',appno='" & txtappno & "'  where INVOICENO = " & vs1.TextMatrix(K, 0) & ""
       End If
         rs1.MoveNext
       Wend
       
       
    End If
    Next
  
Else

    For K = 1 To vs1.rows - 1
    If (vs1.TextMatrix(K, 0) <> "" And vs1.TextMatrix(K, 8) = "") Then
       
       If rs1.State = 1 Then rs1.close
       rs1.Open "select bookcode,fyear from updateInvoiceB_fromAppTbl where id='" & txtScId.text & "' and sername='" & vs1.TextMatrix(K, 4) & "'", con
       While rs1.EOF = False
       
       If rs1!fyear = session Then
         con.Execute "update invoiceB set appno='" & txtappno & "',app_add='y' where (invoiceNo='" & vs1.TextMatrix(K, 0) & "' and bookcode='" & rs1!Bookcode & "')"
         con.Execute " update invoicea set App_Add='y',appno='" & txtappno & "'  where INVOICENO = " & vs1.TextMatrix(K, 0) & ""
       Else
         CON_next.Execute "update invoiceB set appno='" & txtappno & "',app_add='y' where (invoiceNo='" & vs1.TextMatrix(K, 0) & "' and bookcode='" & rs1!Bookcode & "')"
         CON_next.Execute " update invoicea set App_Add='y',appno='" & txtappno & "'  where INVOICENO = " & vs1.TextMatrix(K, 0) & ""

       End If
       rs1.MoveNext
       Wend
       
    End If
    Next

 
End If


'===================================================================

'If Option1_school.value = True Then
'   con.Execute "exec updateInvoiceB_withAppForm '" & fdate1 & "','" & tdate2 & "','School'"
'Else
'   con.Execute "exec updateInvoiceB_withAppForm '" & fdate1 & "','" & tdate2 & "','Party'"
'End If

search_appformdetNew

MsgBox " Saved.....", vbInformation
Commandedit.Enabled = True
cmdSave_2.Enabled = False
cmdModify.Enabled = False
Commanddelete.Enabled = False

End Sub
Sub SearchDataNew()

Dim rss As New ADODB.Recordset


Dim k1 As Integer
k1 = 1


GROSS = 0
net = 0

If Edit = False Then
   vs.rows = 1
End If

cboyesno.Clear


If RS.State = 1 Then RS.close
If rs1.State = 1 Then rs1.close

If Option2_Party.value = True Then
RS.Open "select scid,scname as subledger,sername,discount,sum(grossAmt),sum(NetAmt),fyear from useForApprovalQry where substring([SUBLEDGER],1,5)='" & txtScId.text & "' and ((Appno is null OR Appno='') or (App_Add='n' OR App_Add='')) and " & dt_str & " group by scid,scname,sername,discount,fyear order by scname", con, adOpenDynamic, adLockOptimistic
rs1.Open "select distinct sername from useForApprovalQry where substring([SUBLEDGER],1,5)='" & txtScId.text & "' and ((Appno is null OR Appno='') or (App_Add='n' OR App_Add='')) and " & dt_str & "", con, adOpenDynamic, adLockOptimistic

    
    While rs1.EOF = False
       If Not IsNull(rs1(0)) Then
          cboyesno.AddItem rs1(0)
       End If
          rs1.MoveNext
    Wend


Else

    
    RS.Open "select subledger,sername,discount,sum(grossAmt),sum(NetAmt),fyear from useForApprovalQry where (((Appno is null OR Appno='') or (App_Add='n' OR App_Add='')) and Scid='" & txtScId.text & "' and " & dt_str & ") group by subledger,sername,discount,fyear order by SUBLEDGER", con, adOpenDynamic, adLockOptimistic
    rs1.Open "select distinct sername from useForApprovalQry where (((Appno is null OR Appno='') or (App_Add='n' OR App_Add='')) AND Scid='" & txtScId.text & "'  and " & dt_str & ")", con, adOpenDynamic, adLockOptimistic
    While rs1.EOF = False
         If Not IsNull(rs1(0)) Then
          cboyesno.AddItem rs1(0)
         End If
          rs1.MoveNext
    Wend


End If


If Edit = True Then
   For I = 1 To vs.rows - 1
   If vs.TextMatrix(I, 0) <> "" Then
      k1 = k1 + 1
   End If
   Next
End If


If RS.EOF = False Then

For I = 1 To RS.RecordCount

If RS.EOF = False Then
   If RS!sername <> "" Then
    vs.rows = vs.rows + 1
    If Option2_Party.value = True Then
       If RS!scid <> "" Then
          vs.TextMatrix(k1, 0) = RS!scid & ""
       Else
          vs.TextMatrix(k1, 0) = "-"
       End If
       
       If RS!subledger <> "" Then
          vs.TextMatrix(k1, 1) = RS!subledger & ""
       Else
          vs.TextMatrix(k1, 1) = "-"
       End If
       
       If vs.TextMatrix(k1, 1) = "" Then
       
          vs.TextMatrix(k1, 0) = Trim(Mid(RS!subledger, 1, 6))
          vs.TextMatrix(k1, 1) = Trim(Mid(RS!subledger, 6))
          
       End If
       
       
    Else
       vs.TextMatrix(k1, 0) = Trim(Mid(RS!subledger, 1, 6))
       vs.TextMatrix(k1, 1) = Trim(Mid(RS!subledger, 6))
   
    End If
    vs.TextMatrix(k1, 2) = RS!sername & ""
    vs.TextMatrix(k1, 3) = RS!discount & ""
    vs.TextMatrix(k1, 12) = RS!fyear & ""
    k1 = k1 + 1
    
    GROSS = GROSS + RS(3)
    net = net + RS(4)
    
    End If
    RS.MoveNext
End If



Next

End If

'======================
 search_appformdetNew
'======================

cboyesno.AddItem "ALL"
setWidth


txtGross.text = GROSS
txtNet.text = net

txtGross.text = Round(GROSS, 0)
txtNet.text = Round(net, 0)

lblrow.Caption = "Total : " & vs.rows - 1

End Sub
Sub search_appformdet()

Dim rss As New ADODB.Recordset
Dim dis_, finalAmt As Double

GROSS = 0
net = 0

dis_ = 0
finalAmt = 0

vs1.Clear
vs1.rows = 1
vs1.Cols = 11

If rss.State = 1 Then rss.close
If Option2_Party.value = True Then
  rss.Open "select invoiceNo,InvoiceDate,subledger,Scid,SerName,Fyear,sum(grossAmt),sum(NetAmt),appNo from useForApprovalQry where (substring([SUBLEDGER],1,5)='" & txtScId.text & "' and " & dt_str & ") and (appno=" & txtappno.text & " or appno='')  group by invoiceNo,InvoiceDate,subledger,Scid,SerName,Fyear,appno order by AppNo", con, adOpenDynamic, adLockOptimistic
Else
  rss.Open "select invoiceNo,InvoiceDate,subledger,Scid,SerName,Fyear,sum(grossAmt),sum(NetAmt),appNo from useForApprovalQry where (Scid='" & txtScId.text & "' and " & dt_str & " ) and (appno=" & txtappno.text & " or appno='')  group by invoiceNo,InvoiceDate,subledger,Scid,SerName,Fyear,AppNo order by AppNo", con, adOpenDynamic, adLockOptimistic
End If

For I = 1 To rss.RecordCount

vs1.rows = vs1.rows + 1




vs1.TextMatrix(I, 0) = rss!invoiceNo
vs1.TextMatrix(I, 1) = rss!invoiceDate
vs1.TextMatrix(I, 2) = rss!subledger
vs1.TextMatrix(I, 3) = rss!scid
vs1.TextMatrix(I, 4) = rss!sername
vs1.TextMatrix(I, 5) = rss(6)
vs1.TextMatrix(I, 6) = rss(7)
vs1.TextMatrix(I, 7) = rss!fyear
vs1.TextMatrix(I, 8) = rss!appno & ""

If rss!appno = txtappno.text Then

    If Not IsNull(rss(6)) Then
       GROSS = GROSS + rss(6)
    End If
    If Not IsNull(rss(7)) Then
       net = net + rss(7)
    End If
    
    For a1 = 1 To vs.rows - 1
    If vs.TextMatrix(a1, 0) <> "" Then
    
    If Option1_school.value = True Then
      cc_ = Left(vs1.TextMatrix(I, 2), 5)
      If (rss!sername = vs.TextMatrix(a1, 2) And cc_ = vs.TextMatrix(a1, 0)) Then
         vs1.TextMatrix(I, 9) = (IIf(vs.TextMatrix(a1, 3) = "", 0, Val(vs.TextMatrix(a1, 3))) + IIf(vs.TextMatrix(a1, 4) = "", 0, Val(vs.TextMatrix(a1, 4))) + IIf(vs.TextMatrix(a1, 5) = "", 0, Val(vs.TextMatrix(a1, 5))) + IIf(vs.TextMatrix(a1, 6) = "", 0, Val(vs.TextMatrix(a1, 6))))
         GoTo aaa:
      End If
      
    Else
      cc_ = vs1.TextMatrix(I, 3)
      If vs.TextMatrix(a1, 0) = "-" Then
        If (rss!sername = vs.TextMatrix(a1, 2)) Then
           vs1.TextMatrix(I, 9) = (IIf(vs.TextMatrix(a1, 3) = "", 0, Val(vs.TextMatrix(a1, 3))) + IIf(vs.TextMatrix(a1, 4) = "", 0, Val(vs.TextMatrix(a1, 4))) + IIf(vs.TextMatrix(a1, 5) = "", 0, Val(vs.TextMatrix(a1, 5))) + IIf(vs.TextMatrix(a1, 6) = "", 0, Val(vs.TextMatrix(a1, 6))))
           GoTo aaa:
        End If
      ElseIf (rss!sername = vs.TextMatrix(a1, 2) And cc_ = vs.TextMatrix(a1, 0)) Then
         If (rss!sername = vs.TextMatrix(a1, 2)) Then
           vs1.TextMatrix(I, 9) = (IIf(vs.TextMatrix(a1, 3) = "", 0, Val(vs.TextMatrix(a1, 3))) + IIf(vs.TextMatrix(a1, 4) = "", 0, Val(vs.TextMatrix(a1, 4))) + IIf(vs.TextMatrix(a1, 5) = "", 0, Val(vs.TextMatrix(a1, 5))) + IIf(vs.TextMatrix(a1, 6) = "", 0, Val(vs.TextMatrix(a1, 6))))
           GoTo aaa:
        End If
      End If
    
    
    End If
    
      
    End If
    Next
    
    
aaa:
    
    dis_ = Round(Val(vs1.TextMatrix(I, 5) * Val(vs1.TextMatrix(I, 9)) / 100), 0)
    vs1.TextMatrix(I, 10) = Round(Val(vs1.TextMatrix(I, 5)) - dis_, 0)
    finalAmt = finalAmt + IIf(vs1.TextMatrix(I, 10) = "", 0, Val(vs1.TextMatrix(I, 10)))

End If



If rss!appno <> "" Then
   vs1.Cell(flexcpBackColor, I, 0) = vbGreen
Else
   vs1.Cell(flexcpBackColor, I, 0) = vbWhite
End If


rss.MoveNext

Next

'=========================================================
vs1.FormatString = "inv.No|Inv.Date|Subledger|Scid|SerName|GrossAmt|NetAmt|Fyear|AppNo|TDisAmt|FinalAmt"
vs1.ColWidth(0) = 1000
vs1.ColWidth(1) = 1100
vs1.ColWidth(2) = 3100
vs1.ColWidth(3) = 1100
vs1.ColWidth(4) = 1500
vs1.ColWidth(5) = 1100
vs1.ColWidth(6) = 1100
vs1.ColWidth(7) = 1000
vs1.ColWidth(8) = 1000
vs1.ColWidth(9) = 1000
'=========================================================

txtGross.text = GROSS
txtNet.text = net

txtGross.text = Round(GROSS, 0)
txtNet.text = Round(net, 0)
txtFAmt.text = Round(finalAmt, 0)


End Sub
Sub search_appformdetNew()

Dim rss As New ADODB.Recordset
Dim dis_, finalAmt As Double

GROSS = 0
net = 0

dis_ = 0
finalAmt = 0

vs1.Clear
vs1.rows = 1
vs1.Cols = 11

If rss.State = 1 Then rss.close
If Option2_Party.value = True Then
  rss.Open "select invoiceNo,InvoiceDate,subledger,Scid,SerName,Fyear,sum(grossAmt),sum(NetAmt),appNo from useForApprovalQry where (substring([SUBLEDGER],1,5)='" & txtScId.text & "' and " & dt_str & ")   group by invoiceNo,InvoiceDate,subledger,Scid,SerName,Fyear,appno order by AppNo", con, adOpenDynamic, adLockOptimistic
Else
  rss.Open "select invoiceNo,InvoiceDate,subledger,Scid,SerName,Fyear,sum(grossAmt),sum(NetAmt),appNo from useForApprovalQry where (Scid='" & txtScId.text & "' and " & dt_str & " )   group by invoiceNo,InvoiceDate,subledger,Scid,SerName,Fyear,AppNo order by AppNo", con, adOpenDynamic, adLockOptimistic
End If

For I = 1 To rss.RecordCount

vs1.rows = vs1.rows + 1




vs1.TextMatrix(I, 0) = rss!invoiceNo
vs1.TextMatrix(I, 1) = rss!invoiceDate
vs1.TextMatrix(I, 2) = rss!subledger
vs1.TextMatrix(I, 3) = rss!scid
vs1.TextMatrix(I, 4) = rss!sername
vs1.TextMatrix(I, 5) = rss(6)
vs1.TextMatrix(I, 6) = rss(7)
vs1.TextMatrix(I, 7) = rss!fyear
vs1.TextMatrix(I, 8) = rss!appno & ""

If rss!appno = txtappno.text Then

    If Not IsNull(rss(6)) Then
       GROSS = GROSS + rss(6)
    End If
    If Not IsNull(rss(7)) Then
       net = net + rss(7)
    End If
    
    For a1 = 1 To vs.rows - 1
    If vs.TextMatrix(a1, 0) <> "" Then
    
    If Option1_school.value = True Then
      cc_ = Left(vs1.TextMatrix(I, 2), 5)
      If (rss!sername = vs.TextMatrix(a1, 2) And cc_ = vs.TextMatrix(a1, 0)) Then
         vs1.TextMatrix(I, 9) = (IIf(vs.TextMatrix(a1, 3) = "", 0, Val(vs.TextMatrix(a1, 3))) + IIf(vs.TextMatrix(a1, 4) = "", 0, Val(vs.TextMatrix(a1, 4))) + IIf(vs.TextMatrix(a1, 5) = "", 0, Val(vs.TextMatrix(a1, 5))) + IIf(vs.TextMatrix(a1, 6) = "", 0, Val(vs.TextMatrix(a1, 6))))
         GoTo aaa:
      End If
      
    Else
      cc_ = vs1.TextMatrix(I, 3)
      If vs.TextMatrix(a1, 0) = "-" Then
        If (rss!sername = vs.TextMatrix(a1, 2)) Then
           vs1.TextMatrix(I, 9) = (IIf(vs.TextMatrix(a1, 3) = "", 0, Val(vs.TextMatrix(a1, 3))) + IIf(vs.TextMatrix(a1, 4) = "", 0, Val(vs.TextMatrix(a1, 4))) + IIf(vs.TextMatrix(a1, 5) = "", 0, Val(vs.TextMatrix(a1, 5))) + IIf(vs.TextMatrix(a1, 6) = "", 0, Val(vs.TextMatrix(a1, 6))))
           GoTo aaa:
        End If
      ElseIf (rss!sername = vs.TextMatrix(a1, 2) And cc_ = vs.TextMatrix(a1, 0)) Then
         If (rss!sername = vs.TextMatrix(a1, 2)) Then
           vs1.TextMatrix(I, 9) = (IIf(vs.TextMatrix(a1, 3) = "", 0, Val(vs.TextMatrix(a1, 3))) + IIf(vs.TextMatrix(a1, 4) = "", 0, Val(vs.TextMatrix(a1, 4))) + IIf(vs.TextMatrix(a1, 5) = "", 0, Val(vs.TextMatrix(a1, 5))) + IIf(vs.TextMatrix(a1, 6) = "", 0, Val(vs.TextMatrix(a1, 6))))
           GoTo aaa:
        End If
      End If
    
    
    End If
    
      
    End If
    Next
    
    
aaa:
    
    dis_ = Round(Val(vs1.TextMatrix(I, 5) * Val(vs1.TextMatrix(I, 9)) / 100), 0)
    vs1.TextMatrix(I, 10) = Round(Val(vs1.TextMatrix(I, 5)) - dis_, 0)
    finalAmt = finalAmt + IIf(vs1.TextMatrix(I, 10) = "", 0, Val(vs1.TextMatrix(I, 10)))

End If



If rss!appno <> "" Then
   vs1.Cell(flexcpBackColor, I, 0) = vbGreen
Else
   vs1.Cell(flexcpBackColor, I, 0) = vbWhite
End If


rss.MoveNext

Next

'=========================================================
vs1.FormatString = "inv.No|Inv.Date|Subledger|Scid|SerName|GrossAmt|NetAmt|Fyear|AppNo|TDisAmt|FinalAmt"
vs1.ColWidth(0) = 1000
vs1.ColWidth(1) = 1100
vs1.ColWidth(2) = 3100
vs1.ColWidth(3) = 1100
vs1.ColWidth(4) = 1500
vs1.ColWidth(5) = 1100
vs1.ColWidth(6) = 1100
vs1.ColWidth(7) = 1000
vs1.ColWidth(8) = 1000
vs1.ColWidth(9) = 1000
'=========================================================

txtGross.text = GROSS
txtNet.text = net

txtGross.text = Round(GROSS, 0)
txtNet.text = Round(net, 0)
txtFAmt.text = Round(finalAmt, 0)


End Sub
Sub searchData()

Dim rss As New ADODB.Recordset
Dim k1 As Integer
k1 = 1

Dim GROSS, net As Double
GROSS = 0
net = 0
txtNet.text = 0
txtGross.text = 0
txtNet.text = 0

If Edit = False Then
   vs.rows = 1
End If

cboyesno.Clear

If RS.State = 1 Then RS.close
If rs1.State = 1 Then rs1.close

If Option2_Party.value = True Then

RS.Open "select scid,scname,sername,discount,sum(grossAmt),sum(NetAmt) from useForApprovalQry where substring([SUBLEDGER],1,5)='" & txtScId.text & "' and " & dt_str & " group by scid,scname,sername,discount order by scname", con, adOpenDynamic, adLockOptimistic
rs1.Open "select distinct sername from useForApprovalQry where substring([SUBLEDGER],1,5)='" & txtScId.text & "'", con, adOpenDynamic, adLockOptimistic

    While rs1.EOF = False
       If Not IsNull(rs1(0)) Then
          cboyesno.AddItem rs1(0)
       End If
          rs1.MoveNext
    Wend

Else
    
    RS.Open "select subledger,sername,discount,sum(grossAmt),sum(NetAmt) from useForApprovalQry where (Scid='" & txtScId.text & "' and " & dt_str & ") group by subledger,sername,discount order by SUBLEDGER", con, adOpenDynamic, adLockOptimistic
    rs1.Open "select distinct sername from useForApprovalQry where Scid='" & txtScId.text & "'", con, adOpenDynamic, adLockOptimistic
    While rs1.EOF = False
         If Not IsNull(rs1(0)) Then
          cboyesno.AddItem rs1(0)
         End If
          rs1.MoveNext
    Wend


End If


If RS.EOF = False Then

For I = 1 To RS.RecordCount

If RS.EOF = False Then

If Option2_Party.value = True Then
   If rss.State = 1 Then rss.close
   rss.Open "select * from AppForm where (id='" & txtScId.text & "' and sername='" & RS!sername & "')", con, adOpenDynamic, adLockOptimistic
Else
   If rss.State = 1 Then rss.close
   rss.Open "select * from AppForm where (code='" & txtScId.text & "' and sername='" & RS!sername & "')", con, adOpenDynamic, adLockOptimistic
End If
    
If rss.EOF = False Then
   GROSS = GROSS + RS(3)
   net = net + RS(4)
   
End If
    
If rss.EOF = True Then
    
    If RS!sername <> "" Then
    
    If Edit = True Then
       k1 = I
       Edit = False
    End If
    
    vs.rows = vs.rows + 1
    If Option2_Party.value = True Then
       vs.TextMatrix(k1, 0) = RS!scid & ""
       vs.TextMatrix(k1, 1) = RS!scname & ""
    Else
       vs.TextMatrix(k1, 0) = Trim(Mid(RS!subledger, 1, 6))
       vs.TextMatrix(k1, 1) = Trim(Mid(RS!subledger, 6))
    
    End If
    vs.TextMatrix(k1, 2) = RS!sername & ""
    vs.TextMatrix(k1, 3) = RS!discount & ""
    k1 = k1 + 1
    
    GROSS = GROSS + RS(3)
    net = net + RS(4)
    
    End If
    
End If
    
    RS.MoveNext
End If



Next

End If

cboyesno.AddItem "ALL"

setWidth

txtGross.text = Round(GROSS, 0)
txtNet.text = Round(net, 0)

lblrow.Caption = "Total : " & vs.rows - 1

End Sub

Sub SearchDataEdit()

Dim party_ As String

If RS.State = 1 Then RS.close
'RS.Open "select distinct scid,scname,sername,discount from tmtuseForApprovalQry where substring([SUBLEDGER],1,5)='" & txtscid.Text & "' and sername not in(select distinct SerName from AppForm where (id='" & txtscid.Text & "' or code='" & txtscid.Text & "')) order by scname", con, adOpenDynamic, adLockOptimistic
RS.Open "select scid,scname,sername,discount,sum(grossAmt),sum(NetAmt) from useForApprovalQry where substring([SUBLEDGER],1,5)='" & txtScId.text & "' and sername not in(select distinct SerName from AppForm where (id='" & txtScId.text & "' or code='" & txtScId.text & "')) group by scid,scname,sername,discount order by scname", con, adOpenDynamic, adLockOptimistic
'RS.Open "select subledger,sername,discount,sum(grossAmt),sum(NetAmt) from useForApprovalQry where (Scid='" & txtscid.Text & "' and " & dt_str & ") group by subledger,sername,discount order by SUBLEDGER", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
For I = 1 To RS.RecordCount

            
     If RS!sername <> "" Then
            
            If Option2_Party.value = True Then
               party_ = "Party"
               con.Execute "insert into AppForm(code,PName,School_Party,sername,discount,id,School_PartyName,appper,adj,promo,UserId) values('" & RS!scid & "','" & RS!scname & "','" & party_ & "','" & RS!sername & "','" & RS!discount & "','" & txtScId & "','" & txtSchoolName & "',0,0,0," & UId & ")"
            Else
              party_ = "School"
              con.Execute "insert into AppForm(code,pname,School_Party,sername,discount) values('" & Trim(Mid(RS!subledger, 1, 6)) & "','" & Trim(Mid(RS!subledger, 6)) & "','" & party_ & "','" & RS!sername & "','" & RS!discount & "')"
            End If
            
      End If
            

RS.MoveNext

Next

End If



End Sub
Sub Saved_SearchData()

vs.rows = 2
Check1_upto5.value = 0

cboyesno.Clear
If rs1.State = 1 Then rs1.close
rs1.Open "select distinct SerName from AppForm where (appno=" & txtappno.text & ")", con, adOpenDynamic, adLockOptimistic
While rs1.EOF = False
     
    cboyesno.AddItem rs1(0)
    rs1.MoveNext
Wend
cboyesno.AddItem "ALL"



Set RS = New ADODB.Recordset
RS.Open "select ID,School_PartyName,School_Party,TOD,code,PName,discount,SerName,AppPer,remarks,adj,promo,Net_Gross,appno,appdate,cd,grossAmt,netAmt,noapp,BAuthorized,fyear from AppForm where (appno=" & txtappno.text & ") order by PName,SerName", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then

search_New = True


If RS!bAuthorized = True Then
   txtAuthorised.text = "Yes"
Else
   txtAuthorised.text = "No"
End If
   
For I = 1 To RS.RecordCount

If RS.EOF = False Then

   
   
   vs1.Enabled = False
   vs.Enabled = False
   Commanddelete.Enabled = False
   cmdSave_2.Enabled = False
   cmdModify.Enabled = True
   Commandedit.Enabled = True
   
   If RS!noapp = "y" Then
      Check1_noApp.value = 1
    Else
      Check1_noApp.value = 0
    End If

   
   txtGross.text = RS!GrossAmt & ""
   txtNet.text = RS!netamt & ""
   
   txtScId.text = RS!id
   txtSchoolName.text = RS!School_PartyName
   txtappno = RS!appno & ""
   If IsDate(RS!appdate) Then
   txtAppDate = RS!appdate
   End If
   
   
   If RS!School_Party = "School" Then
      Option1_school.value = True
   Else
      Option2_Party.value = True
   End If

    vs.rows = vs.rows + 1
    
    vs.TextMatrix(I, 0) = RS!Code
    vs.TextMatrix(I, 1) = RS!pname
    vs.TextMatrix(I, 2) = RS!sername
    vs.TextMatrix(I, 3) = RS!discount
    vs.TextMatrix(I, 4) = RS!appper & ""
    
    If RS!appper > 0 Then
       Check1_upto5.value = 1
    End If

    
    vs.TextMatrix(I, 5) = RS!adj & ""
    vs.TextMatrix(I, 6) = RS!Promo & ""
    vs.TextMatrix(I, 7) = RS!Net_Gross & ""
    vs.TextMatrix(I, 8) = RS!tod & ""
    
    vs.TextMatrix(I, 9) = RS!cd & ""
    vs.TextMatrix(I, 10) = RS!remarks & ""
    If RS!bAuthorized = True Then
       vs.TextMatrix(I, 11) = 1
    Else
       vs.TextMatrix(I, 11) = 0
    End If
    
    vs.TextMatrix(I, 12) = RS!fyear & ""
    RS.MoveNext

End If
Next

End If


'Check1_upto5.value = True

'If RS!appper > 0 Then
'   Check1_upto5.value = 1
'End If


search_appformdet

setWidth

lblrow.Caption = "Total : " & vs.rows - 1

End Sub
Sub Saved_SearchData_New()

'---------------------------------------------------------------------------


vs.rows = 2
Check1_upto5.value = 0

cboyesno.Clear
If rs1.State = 1 Then rs1.close
rs1.Open "select distinct SerName from AppForm where (appno=" & txtappno.text & ")", con_LAST, adOpenDynamic, adLockOptimistic
While rs1.EOF = False
     
    cboyesno.AddItem rs1(0)
    rs1.MoveNext
Wend
cboyesno.AddItem "ALL"



Set RS = New ADODB.Recordset
RS.Open "select ID,School_PartyName,School_Party,TOD,code,PName,discount,SerName,AppPer,remarks,adj,promo,Net_Gross,appno,appdate,cd,grossAmt,netAmt,noapp,BAuthorized,fyear from AppForm where (appno=" & txtappno.text & ") order by PName,SerName", con_LAST, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then


If RS!bAuthorized = True Then
   txtAuthorised.text = "Yes"
Else
   txtAuthorised.text = "No"
End If
   
For I = 1 To RS.RecordCount

If RS.EOF = False Then

   
   
   vs1.Enabled = False
   vs.Enabled = False
   Commanddelete.Enabled = False
   cmdSave_2.Enabled = False
   cmdModify.Enabled = False
   Commandedit.Enabled = False
   cmdPendingApp.Enabled = False
   CommandPrint.Enabled = False
   cmdFindSchoolNameChange.Enabled = False
   cmdAdd_1.Enabled = False
   Commandsearch.Enabled = False
   
   
   If RS!noapp = "y" Then
      Check1_noApp.value = 1
    Else
      Check1_noApp.value = 0
    End If

   
   txtGross.text = RS!GrossAmt & ""
   txtNet.text = RS!netamt & ""
   
   txtScId.text = RS!id
   txtSchoolName.text = RS!School_PartyName
   txtappno = RS!appno & ""
   If IsDate(RS!appdate) Then
   txtAppDate = RS!appdate
   End If
   
   
   If RS!School_Party = "School" Then
      Option1_school.value = True
   Else
      Option2_Party.value = True
   End If

    vs.rows = vs.rows + 1
    
    vs.TextMatrix(I, 0) = RS!Code
    vs.TextMatrix(I, 1) = RS!pname
    vs.TextMatrix(I, 2) = RS!sername
    vs.TextMatrix(I, 3) = RS!discount
    vs.TextMatrix(I, 4) = RS!appper & ""
    
    If RS!appper > 0 Then
       Check1_upto5.value = 1
    End If
    
    
    vs.TextMatrix(I, 5) = RS!adj & ""
    vs.TextMatrix(I, 6) = RS!Promo & ""
    vs.TextMatrix(I, 7) = RS!Net_Gross & ""
    vs.TextMatrix(I, 8) = RS!tod & ""
    
    vs.TextMatrix(I, 9) = RS!cd & ""
    vs.TextMatrix(I, 10) = RS!remarks & ""
    If RS!bAuthorized = True Then
    vs.TextMatrix(I, 11) = 1
    Else
    vs.TextMatrix(I, 11) = 0
    End If
    vs.TextMatrix(I, 12) = RS!fyear & ""
    RS.MoveNext

End If
Next

End If

'=======================================================================

Dim rss As New ADODB.Recordset
Dim dis_, finalAmt As Double

GROSS = 0
net = 0

dis_ = 0
finalAmt = 0

vs1.Clear
vs1.rows = 1
vs1.Cols = 11

If IsEmpty(appno) Then
   appno = txtappno.text
End If

If rss.State = 1 Then rss.close
If Option2_Party.value = True Then
  rss.Open "select invoiceNo,InvoiceDate,subledger,Scid,SerName,Fyear,sum(grossAmt),sum(NetAmt),appNo from useForApprovalQry where (substring(SUBLEDGER,1,5)='" & txtScId.text & "' and appno=" & appno & ") group by invoiceNo,InvoiceDate,subledger,Scid,SerName,Fyear,appno order by AppNo", con_LAST, adOpenDynamic, adLockOptimistic
Else
  rss.Open "select invoiceNo,InvoiceDate,subledger,Scid,SerName,Fyear,sum(grossAmt),sum(NetAmt),appNo from useForApprovalQry where (Scid='" & txtScId.text & "' and appno=" & txtappno & ") group by invoiceNo,InvoiceDate,subledger,Scid,SerName,Fyear,AppNo order by AppNo", con_LAST, adOpenDynamic, adLockOptimistic
End If

For I = 1 To rss.RecordCount

vs1.rows = vs1.rows + 1




vs1.TextMatrix(I, 0) = rss!invoiceNo
vs1.TextMatrix(I, 1) = rss!invoiceDate
vs1.TextMatrix(I, 2) = rss!subledger
vs1.TextMatrix(I, 3) = rss!scid
vs1.TextMatrix(I, 4) = rss!sername
vs1.TextMatrix(I, 5) = rss(6)
vs1.TextMatrix(I, 6) = rss(7)
vs1.TextMatrix(I, 7) = rss!fyear
vs1.TextMatrix(I, 8) = rss!appno & ""
'vs1.TextMatrix(I, 9) = RS!Tdis & ""
'vs1.TextMatrix(I, 10) = RS!TdisAmt & ""

If rss!appno = txtappno.text Then

    If Not IsNull(rss(6)) Then
       GROSS = GROSS + rss(6)
    End If
    If Not IsNull(rss(7)) Then
       net = net + rss(7)
    End If
    
    For a1 = 1 To vs.rows - 1
    If vs.TextMatrix(a1, 0) <> "" Then
      If (rss!sername = vs.TextMatrix(a1, 2) And rss!sername = vs.TextMatrix(a1, 2)) Then
         vs1.TextMatrix(I, 9) = (IIf(vs.TextMatrix(a1, 3) = "", 0, Val(vs.TextMatrix(a1, 3))) + IIf(vs.TextMatrix(a1, 4) = "", 0, Val(vs.TextMatrix(a1, 4))) + IIf(vs.TextMatrix(a1, 5) = "", 0, Val(vs.TextMatrix(a1, 5))) + IIf(vs.TextMatrix(a1, 6) = "", 0, Val(vs.TextMatrix(a1, 6))))
         GoTo aaa:
      End If
      End If
    Next
aaa:
    
    dis_ = Round(Val(vs1.TextMatrix(I, 5) * Val(vs1.TextMatrix(I, 9)) / 100), 0)
    vs1.TextMatrix(I, 10) = Round(Val(vs1.TextMatrix(I, 5)) - dis_, 0)
    finalAmt = finalAmt + IIf(vs1.TextMatrix(I, 10) = "", 0, Val(vs1.TextMatrix(I, 10)))

End If


If rss!appno <> "" Then
   vs1.Cell(flexcpBackColor, I, 0) = vbGreen
Else
   vs1.Cell(flexcpBackColor, I, 0) = vbWhite
End If


rss.MoveNext

Next

'=========================================================
vs1.FormatString = "inv.No|Inv.Date|Subledger|Scid|SerName|GrossAmt|NetAmt|Fyear|AppNo|TDisAmt|FinalAmt"
vs1.ColWidth(0) = 1000
vs1.ColWidth(1) = 1100
vs1.ColWidth(2) = 3100
vs1.ColWidth(3) = 1100
vs1.ColWidth(4) = 1500
vs1.ColWidth(5) = 1100
vs1.ColWidth(6) = 1100
vs1.ColWidth(7) = 1000
vs1.ColWidth(8) = 1000
vs1.ColWidth(9) = 1000
'=========================================================

txtGross.text = GROSS
txtNet.text = net

txtGross.text = Round(GROSS, 0)
txtNet.text = Round(net, 0)
txtFAmt.text = Round(finalAmt, 0)


'=======================================================================

setWidth

lblrow.Caption = "Total : " & vs.rows - 1

End Sub

Private Sub cmdSchWisePartywise_Click()
    
    If (next_dbase = current_dbase) Then
        db10 = ""
        con.Execute "exec Sp_SchoolWisePartwiseApproval '" & db10 & "'"
    Else
        db10 = next_dbase & ".[dbo]"
        con.Execute "exec Sp_SchoolWisePartwiseApproval '" & db10 & "'"
    End If
    
    DSNNew
    
    
    If MsgBox("Want to Print ?", vbQuestion + vbYesNo) = vbYes Then
    
   
    CR.Reset
    CR.ReportFileName = rptPath & "/PartySchoolWiseAppDet.rpt"
    CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    CR.WindowShowPrintSetupBtn = True
    CR.WindowShowPrintBtn = True
    CR.WindowShowExportBtn = True
    CR.WindowState = crptMaximized
    CR.Action = 1
    
    End If
    
End Sub
Private Sub Command1_Click()

st_ = "SELECT  dbo.ApprovalDet.INVOICENO, dbo.ApprovalDet.SUBLEDGER, dbo.ApprovalDet.SerName, dbo.BOOKS.BOOKCODE,dbo.ApprovalDet.appno " & _
    "FROM dbo.ApprovalDet  INNER JOIN dbo.BOOKS ON dbo.ApprovalDet.SerName = dbo.BOOKS.SerName where dbo.ApprovalDet.appno='393'"
    
If rs1.State = 1 Then rs1.close
rs1.Open st_, con
While rs1.EOF = False
  con.Execute "update invoiceb set app_add='y',appno='" & rs1(4) & "' where (invoiceno=" & rs1(0) & " and subledger='" & rs1(1) & "' and bookcode='" & rs1(3) & "')"
  'con.Execute "update invoiceb set appno='" & rs1(4) & "' where (invoiceno=" & rs1(0) & " and subledger='" & rs1(1) & "' and bookcode='" & rs1(3) & "')"
  rs1.MoveNext
Wend


con.Execute "update a set a.App_Add=b.App_Add from invoicea as a " & _
"inner join invoiceb as b on (a.INVOICENO =5674) "
 
    
con.Execute "update a set a.appNo= b.AppNo  from invoicea as a " & _
"inner join invoiceb as b on (a.INVOICENO =5674)"

    
MsgBox "update..."
    
End Sub

Private Sub Command2_Click()
frmCheck.Visible = False
End Sub

Private Sub Command3_Click()
frmCheck.Visible = False
End Sub

Private Sub Commanddelete_Click()

If RS.State = 1 Then RS.close
RS.Open "select fromDate,toDate,NotCreated from turnOverDis where fyear='" & session & "'", CCON
If RS.EOF = False Then
  If RS!NotCreated = "y" Then
   fdate1 = RS!fromdate
   tdate2 = RS!todate
  End If
End If

If rs1.State = 1 Then rs1.close
rs1.Open "select top 1 * from AppForm where appno=" & txtappno.text & "", con, adOpenStatic, adLockReadOnly
If rs1.EOF = False Then
    If rs1!bAuthorized = True Then
        MsgBox "You can'nt change, this Approval No Locked !!", vbExclamation, "Alert"
        Exit Sub
    End If
   
End If




If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
   
   
If Option1_school.value = True Then
   
   con.Execute "update a set a.App_Add='n',a.AppNo= b.appno from invoiceb as a " & _
   "inner join updateInvoiceB_fromAppTbl as b on (a.bookcode=b.bookcode and a.subledger=b.subledger1) " & _
   "INNER JOIN invoicea as inva on (a.INVOICENO=inva.INVOICENO and inva.ScID = b.id and substring(inva.SUBLEDGER,1,5)=b.code) " & _
   " where b.appno='" & txtappno.text & "'"
   
Else
  con.Execute "update a set a.AppNo= b.appno,a.App_Add='n' from invoiceb as a " & _
  "inner join updateInvoiceB_fromAppTbl as b on (a.bookcode=b.bookcode and a.subledger=b.subledger) " & _
  "INNER JOIN invoicea as inva on (a.INVOICENO=inva.INVOICENO and inva.ScID = b.code and substring(inva.SUBLEDGER,1,5)=b.ID) " & _
  " where b.appno='" & txtappno.text & "'"
End If



con.Execute " update invoicea set App_Add='n',appno=''  where appno='" & txtappno.text & "'"
con.Execute " update invoiceb set App_Add='n',appno=''  where appno='" & txtappno.text & "'"
''CON_next.Execute "update invoicea set App_Add='n',appno=''  where appno='" & txtAppNo.text & "'"
''CON_next.Execute "update invoiceb set App_Add='n',appno=''  where appno='" & txtAppNo.text & "'"


con.Execute "delete from AppForm where AppNO='" & txtappno.text & "'"
con.Execute "delete from ApprovalDet where AppNO='" & txtappno.text & "'"

 

''For I = 1 To vs1.Rows - 1
''If vs1.TextMatrix(I, 0) = "n" Then
''
''  con.Execute "update invoicea set app_add='n' where appno='" & txtAppNo.Text & "'"
''  con.Execute "update invoiceb set app_add='n' where appno='" & txtAppNo.Text & "'"
''
''End If
''Next



'===================================================================
If Option1_school.value = True Then
    
    For K = 1 To vs1.rows - 1
    If (vs1.TextMatrix(K, 0) <> "" And vs1.TextMatrix(K, 8) = "") Then
       If rs1.State = 1 Then rs1.close
       rs1.Open "select bookcode,fyear from updateInvoiceB_fromAppTbl where id='" & txtScId.text & "' and sername='" & vs1.TextMatrix(K, 4) & "'", con
       While rs1.EOF = False
         
       If rs1!fyear = session Then
         con.Execute "update invoiceB set appno='',app_add='n' where (invoiceNo='" & vs1.TextMatrix(K, 0) & "' and bookcode='" & rs1!Bookcode & "')"
         con.Execute " update invoicea set App_Add='n',appno=''  where INVOICENO = " & vs1.TextMatrix(K, 0) & ""
         
       Else
         CON_next.Execute "update invoiceB set appno='',app_add='n' where (invoiceNo='" & vs1.TextMatrix(K, 0) & "' and bookcode='" & rs1!Bookcode & "')"
         CON_next.Execute " update invoicea set App_Add='n',appno=''  where INVOICENO = " & vs1.TextMatrix(K, 0) & ""
         
       End If
         rs1.MoveNext
       Wend
       
       
    End If
    Next
  
Else

    For K = 1 To vs1.rows - 1
    If (vs1.TextMatrix(K, 0) <> "" And vs1.TextMatrix(K, 8) = "") Then
       
       If rs1.State = 1 Then rs1.close
       rs1.Open "select bookcode,fyear from updateInvoiceB_fromAppTbl where id='" & txtScId.text & "' and sername='" & vs1.TextMatrix(K, 4) & "'", con
       While rs1.EOF = False
       
       If rs1!fyear = session Then
         con.Execute "update invoiceB set appno='',app_add='n' where (invoiceNo='" & vs1.TextMatrix(K, 0) & "' and bookcode='" & rs1!Bookcode & "')"
         con.Execute " update invoicea set App_Add='n',appno=''  where INVOICENO = " & vs1.TextMatrix(K, 0) & ""
       Else
         CON_next.Execute "update invoiceB set appno='',app_add='n' where (invoiceNo='" & vs1.TextMatrix(K, 0) & "' and bookcode='" & rs1!Bookcode & "')"
         CON_next.Execute " update invoicea set App_Add='n',appno=''  where INVOICENO = " & vs1.TextMatrix(K, 0) & ""

       End If
       rs1.MoveNext
       Wend
       
    End If
    Next

 
End If





'=============
   
cmdAdd_1_Click

End If

End Sub

Private Sub Commandedit_Click()
   

If txtAuthorised.text = "Yes" Then

    Set RS = New ADODB.Recordset
    'RS.Open "select sername from AppForm where appno=" & txtAppNo & " group by sername", con
    For I = 1 To vs1.rows - 1
    'While RS.EOF = False
    If rs1.State = 1 Then rs1.close
    rs1.Open "select top 1 sername from AppForm where (appno=" & txtappno & " and sername='" & vs1.TextMatrix(I, 4) & "')", con
    If rs1.EOF = True Then
      If (vs1.TextMatrix(I, 8) = "" Or vs1.TextMatrix(I, 8) = txtappno) Then
       MsgBox "First you have to Un Authorised this App. No...", vbCritical
       Exit Sub
      End If
    End If
    'RS.MoveNext
    Next
    'Wend

End If
   
   


If MsgBox("Want to Edit ?", vbQuestion + vbYesNo) = vbYes Then
   'SearchDataEdit
   Saved_SearchData
   'con.Execute "delete from AppForm where AppNO='" & txtAppNo.Text & "'"
   Edit = True
   SearchDataNew
   Commanddelete.Enabled = True
   cmdModify.Enabled = True
   cmdSave_2.Enabled = True
   cmdModify.Enabled = True
   Commandedit.Enabled = False
   vs.Enabled = True
   vs1.Enabled = True
End If

End Sub
Private Sub CommandPrint_Click()

'If MsgBox("Want to Print ?", vbQuestion + vbYesNo) = vbYes Then
    
    If txtappno.text <> "" Then
       
       If txtNet.text = "" Then txtNet.text = 0
       If txtGross.text = "" Then txtGross.text = 0
       If txtFAmt.text = "" Then txtFAmt.text = 0
       
       Dim gross_, net_, final_ As Double
       gross_ = 0
       net_ = 0
       final_ = 0
       
       For k1 = 1 To vs1.rows - 1
       If Val(vs1.TextMatrix(k1, 8)) = Val(txtappno.text) Then
       
''       Set RS = New ADODB.Recordset
''       RS.Open "select GAMOUNT,NETAMOUNT,TdisAmt from ApprovalDet where appno=" & Val(txtAppNo.Text) & "", con
''       While RS.EOF = False
              
              gross_ = gross_ + Val(vs1.TextMatrix(k1, 5))
              '''gross_ = gross_ + Val(RS!GAMOUNT)
              net_ = net_ + Val(vs1.TextMatrix(k1, 6))
              '''net_ = net_ + Val(RS!netamount)
              final_ = final_ + Val(vs1.TextMatrix(k1, 10))
              '''final_ = final_ + IIf(IsNull(RS!TdisAmt), 0, RS!TdisAmt)
              '''RS.MoveNext
       '''Wend
       
       End If
       Next
       
       
       txtNet.text = net_
       txtGross.text = gross_
       txtFAmt.text = final_
       
       
       
       con.Execute "update AppForm set GrossAmt=" & txtGross.text & ",NetAmt=" & txtNet.text & ",FinalAmt=" & txtFAmt & " where appno=" & txtappno.text & ""
    
    End If
    
    
    Dim aa
    aa = Mid(session, 6)
    
    If Val(aa) >= 23 Then
    
        If (Option1_school.value = True) Then
           con.Execute "exec UpDateAppTDis '" & txtScId.text & "','School'"
        Else
           con.Execute "exec UpDateAppTDis '" & txtScId.text & "','Party'"
        End If
    
    End If
    
    
    DSNNew
    
    party_ = "print"
    printPRemarks
    
    Dim y_ As String

    

    CR.Reset
    CR.ReportFileName = rptPath & "/Approval.rpt"
    CR.ReplaceSelectionFormula "{AppForm.appno}=" & txtappno & ""
    CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
     
    If Check1_upto5.value = 1 Then
       y_ = "Yes"
    Else
       y_ = "No"
    End If
    CR.Formulas(0) = "upto5='" & y_ & "'"
    CR.Formulas(1) = "fyear='" & session & "'"
    
    CR.WindowShowPrintSetupBtn = True
    CR.WindowShowPrintBtn = True
    CR.WindowShowExportBtn = True
    CR.WindowState = crptMaximized
    CR.Action = 1
    
'    If sss = 13 Then
'       MsgBox "s"
'    End If

'End If

End Sub

Private Sub Commandsearch_Click()

'searchType = "party"
'value = "SELECT distinct AppNO,School_PartyName,School_Party,AppDATE FROM AppForm order by AppNO"
'popuplist_client value, con

search_ = "f2"
searchType = "inv"
popuplistFast "SELECT distinct AppNO,AppDATE,School_PartyName,School_Party FROM AppForm order by AppNO", con, , , "approval"
    

'set_focus = True

End Sub
Private Sub Commandsearch_GotFocus()

  If PopUpValue1 <> "" Then
     
     txtappno.text = PopUpValue1
     txtAppDate.text = PopUpValue2
     
     Saved_SearchData
     
     
     PopUpValue1 = ""
     PopUpValue2 = ""
     PopUpValue3 = ""
     popupvalue4 = ""
  
  End If


End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Unload Me
End If
End Sub
Private Sub Form_Load()

search_ = False

Me.top = 50
Me.Left = 50
Me.Height = 10395
Me.Width = 15100

fillGP_school

Set RS = New ADODB.Recordset
RS.Open "select fromDate,toDate,NotCreated from turnOverDis where fyear='" & session & "'", CCON
If RS.EOF = False Then
  If RS!NotCreated = "y" Then
   dt_str = "(INVOICEDATE >= convert(smalldatetime,'" & RS!fromdate & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!todate & "',103))"
   fdate1 = RS!fromdate
   tdate2 = RS!todate
  End If
End If


Set RS = New ADODB.Recordset
RS.Open "select fromDate,toDate,NotCreated,DataBase from turnOverDis where Current_Next='next'", CCON
If RS.EOF = False Then
  If RS!NotCreated = "y" Then
    tdate2 = RS!todate
    dt_str = "(INVOICEDATE >= convert(smalldatetime,'" & fdate1 & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!todate & "',103))"
  
  
    Set CON_next = New ADODB.Connection
    If LCase(server_) = "server" Then
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & RS!Database & "; UID=" & sql_user & "; PWD=" & sql_pass
       CON_next.Open
       
    Else
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=HP\SQL2008; DATABASE=" & RS!Database & "; UID=; PWD=;"
       CON_next.Open
    End If
  End If
End If


txtAppDate.text = Format(Date, "dd/MM/yyyy")


setWidth
maxId
con.Execute "exec UseApproval " & UId & ""


'=====================
If inviceNo <> "" Then
   txtappno.text = inviceNo
   
   If RS.State = 1 Then RS.close
   RS.Open "select top 1 appno ApprovalDet from ApprovalDet where invoiceno=" & PopUpValue6 & "", con
   If RS.EOF = False Then
      Saved_SearchData
   Else
      Saved_SearchData_New
   End If
   
   PopUpValue6 = ""
   inviceNo = ""
End If


End Sub
Sub setWidth()


vs.FormatString = "Code|Party/SchoolName|SerName|B.Rate|5 (%)| Adj(%)|Promotion(%)|Net/Gross|TOD|CD|Remarks||Fyear"
vs.ColWidth(0) = 700
vs.ColWidth(1) = 3500
vs.ColWidth(2) = 1600
vs.ColWidth(3) = 1000
vs.ColWidth(4) = 1000   'up to 5%
vs.ColWidth(5) = 1000   'Adj %
vs.ColWidth(6) = 1000   'promotion
vs.ColWidth(7) = 1000
vs.ColWidth(8) = 800
vs.ColWidth(9) = 800
vs.ColWidth(10) = 1400
vs.ColWidth(11) = 0
vs.ColWidth(12) = 600

End Sub

Private Sub txtSchoolName_GotFocus()

On Error GoTo aa1:

If Option1_school.value = True Then

        If PopUpValue1 <> "" Then
            txtSchoolName.text = PopUpValue1
            txtScId.text = PopUpValue2
            If txtSchoolName.text <> "" Then
              
             SearchDataNew
              
            End If

        End If
    
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""

   
Else

   If PopUpValue1 <> "" Then
        txtSchoolName.text = PopUpValue1   '& ", " & PopUpValue2
        txtScId.text = PopUpValue2
        If txtSchoolName.text <> "" Then
           SearchDataNew
        End If

   End If
   
   PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""


End If

Exit Sub

aa1:

MsgBox "" & err.DESCRIPTION
End Sub

Private Sub txtSchoolName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    
If Option1_school.value = True Then

    
    If Check1_schoolAll.value = 0 Then
        
        searchType = "party"
        
        'value = "SELECT des as ScName,Billtype as ScID FROM tempLedger_net group by des,Billtype"
        value = "SELECT ScName,ScID FROM useForApprovalQry where " & dt_str & " group by ScName,ScID order by ScName"
        popuplist_client value, con
        set_focus = True
     
     Else
        
        Screen.MousePointer = vbHourglass
        searchType = "party"
        value = "SELECT  School,City,CollegeID as ScID FROM collegeView_ind order by School"
        popuplist_client value, CON_blue
        set_focus = True
        Screen.MousePointer = vbDefault
        
    End If
    
Else

  Screen.MousePointer = vbHourglass

   searchType = "party"
   'value = "SELECT  DESCFORINVOICE as PartyName,address3 as City,Code FROM SLEDGER where len(DESCFORINVOICE)>0 order by DESCFORINVOICE"
   value = "SELECT  distinct substring(SUBLEDGER,7,150) as  PartyName,PCode as Code  FROM useForApprovalQry order by substring(SUBLEDGER,7,150)"
   
   popuplist_client value, con
   set_focus = True
   
   Screen.MousePointer = vbDefault
    
End If
    
    
End If


End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
     
     If KeyCode = 115 Then
        
     If vs.TextMatrix(vs.RowSel, 11) <> "" Then
        If vs.TextMatrix(vs.RowSel, 11) = 0 Then
           vs.RemoveItem (vs.RowSel)
        Else
           MsgBox "You can'nt change, this Approval No Locked !!", vbExclamation, "Alert"
           Exit Sub
        End If
     End If
     
     End If
     
     
     If KeyCode = 13 Then
        
        sc_ = ""
        
        If (vs.Col = 4 Or vs.Col = 5 Or vs.Col = 6 Or vs.Col = 7) Then
           sendkeys "{down}"
        ElseIf vs.Col = 8 Then
           sc_ = vs.TextMatrix(vs.RowSel, 1)
           sendkeys "{down}"
           
           For I = 1 To vs.rows - 1
               If sc_ = vs.TextMatrix(I, 1) Then
                  vs.TextMatrix(I, 8) = vs.TextMatrix(vs.RowSel, 8)
               End If
           Next
        End If
        
        
        
     End If
     
End Sub
Private Sub vs_SelChange()
vs.TextMatrix(vs.RowSel, 11) = IIf(vs.TextMatrix(vs.RowSel, 11) = "", 0, vs.TextMatrix(vs.RowSel, 11))
If vs.TextMatrix(vs.RowSel, 11) <> "" Then
 If (vs.TextMatrix(vs.RowSel, 11) = 0) Then
    vs.Editable = flexEDKbdMouse
 Else
    vs.Editable = flexEDNone
 End If
End If
End Sub

Private Sub vs1_DblClick()
If MsgBox("Want to Update ?", vbQuestion + vbYesNo) = vbYes Then
   
   If Option1_school.value = True Then
    If vs1.TextMatrix(vs1.RowSel, 0) <> "" Then
       If rs1.State = 1 Then rs1.close
       'rs1.Open "select bookcode from updateInvoiceB_fromAppTbl where (SUBLEDGER1='" & vs1.TextMatrix(vs1.RowSel, 2) & "' and id='" & vs1.TextMatrix(vs1.RowSel, 3) & "' and sername='" & vs1.TextMatrix(vs1.RowSel, 4) & "')", con
       rs1.Open "SELECT   BOOKCODE, SUBLEDGER," & _
       " useForApprovalQry.ScID, useForApprovalQry.SerName FROM  useForApprovalQry " & _
       " LEFT OUTER JOIN BOOKS ON useForApprovalQry.SerName = BOOKS.SerName " & _
       " where (useForApprovalQry.invoiceno=" & vs1.TextMatrix(vs1.RowSel, 0) & " and useForApprovalQry.SUBLEDGER='" & vs1.TextMatrix(vs1.RowSel, 2) & "' and useForApprovalQry.scid='" & vs1.TextMatrix(vs1.RowSel, 3) & "' and useForApprovalQry.sername='" & vs1.TextMatrix(vs1.RowSel, 4) & "')", con
       For k1 = 1 To rs1.RecordCount
         
         If Val(Right(session, 2)) = Val(Right(vs1.TextMatrix(vs1.RowSel, 7), 2)) Then
            con.Execute "update invoiceB set appno='',app_add='n' where (invoiceno='" & vs1.TextMatrix(vs1.RowSel, 0) & "' and bookcode='" & rs1!Bookcode & "')"
         Else
            CON_next.Execute "update invoiceB set appno='',app_add='n' where (invoiceno='" & vs1.TextMatrix(vs1.RowSel, 0) & "' and bookcode='" & rs1!Bookcode & "')"
         End If
         
         rs1.MoveNext
       Next
       
       '''con.Execute "update invoiceb set App_Add='n',appno=''  where (SUBLEDGER='" & vs1.TextMatrix(vs1.RowSel, 2) & "' AND invoiceNo='" & vs1.TextMatrix(vs1.RowSel, 0) & "')"
       If Val(Right(session, 2)) = Val(Right(vs1.TextMatrix(vs1.RowSel, 7), 2)) Then
          con.Execute "update invoicea set App_Add='n',appno=''  where (SUBLEDGER='" & vs1.TextMatrix(vs1.RowSel, 2) & "' AND invoiceNo='" & vs1.TextMatrix(vs1.RowSel, 0) & "')"
       Else
          CON_next.Execute "update invoicea set App_Add='n',appno=''  where (SUBLEDGER='" & vs1.TextMatrix(vs1.RowSel, 2) & "' AND invoiceNo='" & vs1.TextMatrix(vs1.RowSel, 0) & "')"
       End If
       vs1.TextMatrix(vs1.RowSel, 8) = ""
    End If
    
    


Else

    If vs1.TextMatrix(vs1.RowSel, 0) <> "" Then
       If rs1.State = 1 Then rs1.close
       'rs1.Open "select bookcode from updateInvoiceB_fromAppTbl where appno='" & txtAppNo.Text & "' and code='" & txtscid.Text & "' and sername='" & vs1.TextMatrix(vs1.RowSel, 4) & "'", con
        rs1.Open "SELECT   BOOKCODE, SUBLEDGER," & _
       " useForApprovalQry.ScID, useForApprovalQry.SerName FROM  useForApprovalQry " & _
       " LEFT OUTER JOIN BOOKS ON useForApprovalQry.SerName = BOOKS.SerName " & _
       " where (useForApprovalQry.invoiceno=" & vs1.TextMatrix(vs1.RowSel, 0) & " and useForApprovalQry.SUBLEDGER='" & vs1.TextMatrix(vs1.RowSel, 2) & "' and useForApprovalQry.scid='" & vs1.TextMatrix(vs1.RowSel, 3) & "' and useForApprovalQry.sername='" & vs1.TextMatrix(vs1.RowSel, 4) & "')", con
    
       For k1 = 1 To rs1.RecordCount
         If Val(Right(session, 2)) = Val(Right(vs1.TextMatrix(vs1.RowSel, 7), 2)) Then
            con.Execute "update invoiceB set appno='',app_add='n' where (invoiceno='" & vs1.TextMatrix(vs1.RowSel, 0) & "' and bookcode='" & rs1!Bookcode & "')"
         Else
            CON_next.Execute "update invoiceB set appno='',app_add='n' where (invoiceno='" & vs1.TextMatrix(vs1.RowSel, 0) & "' and bookcode='" & rs1!Bookcode & "')"
         End If
         rs1.MoveNext
       Next
       
       If Val(Right(session, 2)) = Val(Right(vs1.TextMatrix(vs1.RowSel, 7), 2)) Then
          con.Execute "update invoicea set App_Add='n',appno=''  where (SUBLEDGER='" & vs1.TextMatrix(vs1.RowSel, 2) & "' AND invoiceNo='" & vs1.TextMatrix(vs1.RowSel, 0) & "')"
       Else
          CON_next.Execute "update invoicea set App_Add='n',appno=''  where (SUBLEDGER='" & vs1.TextMatrix(vs1.RowSel, 2) & "' AND invoiceNo='" & vs1.TextMatrix(vs1.RowSel, 0) & "')"
       End If
       
       vs1.TextMatrix(vs1.RowSel, 8) = ""
    End If
    
 
End If
End If

'===================================================================
SearchDataNew
      
End Sub

Private Sub vs1_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 115 Then
   
If vs1.TextMatrix(vs1.RowSel, 8) = "" Then

   con.Execute "update INVOICEA set App_Add='n',appno='' where invoiceno='" & vs1.TextMatrix(vs1.RowSel, 0) & "'"
   con.Execute "update INVOICEB set App_Add='n',appno='' where invoiceno='" & vs1.TextMatrix(vs1.RowSel, 0) & "'"
   vs1.RemoveItem (vs1.RowSel)
Else
   MsgBox "You can'nt change, this Approval No Locked !!", vbExclamation, "Alert"
   Exit Sub
End If
 
   
End If
     
GROSS = 0
net = 0

For I = 1 To vs1.rows - 1

If vs1.TextMatrix(I, 1) <> "" Then
   If vs1.TextMatrix(I, 8) = txtappno.text Then
        GROSS = GROSS + vs1.TextMatrix(I, 5)
        net = net + vs1.TextMatrix(I, 6)
   End If
End If

Next


txtGross.text = Round(GROSS, 0)
txtNet.text = Round(net, 0)


     
End Sub
Private Sub vs1_SelChange()
 If vs1.Col = 9 Then
    vs1.Editable = flexEDKbdMouse
 Else
    vs1.Editable = flexEDNone
 End If

 
End Sub
