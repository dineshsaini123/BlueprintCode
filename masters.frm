VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form master 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9375
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "masters.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin VB.Frame gledger 
      Height          =   3615
      Left            =   420
      TabIndex        =   100
      Top             =   120
      Width           =   7935
      Begin VB.CheckBox GMASTERSL 
         Alignment       =   1  'Right Justify
         Caption         =   "Contains Sub Ledgers  "
         Height          =   255
         Left            =   360
         TabIndex        =   108
         Top             =   2370
         Width           =   2955
      End
      Begin VB.CheckBox GMASTERPL 
         Alignment       =   1  'Right Justify
         Caption         =   "To be Included in P&&L"
         Height          =   255
         Left            =   360
         TabIndex        =   107
         Top             =   1680
         Width           =   2955
      End
      Begin VB.CheckBox GMASTERBS 
         Alignment       =   1  'Right Justify
         Caption         =   "To be included in B\S"
         Height          =   255
         Left            =   360
         TabIndex        =   106
         Top             =   2040
         Width           =   2955
      End
      Begin VB.ComboBox ComboSPECIALCATEGORY 
         Height          =   315
         Left            =   3120
         TabIndex        =   105
         Top             =   800
         Width           =   1545
      End
      Begin VB.TextBox Textglgeneralledgerdiscription 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3120
         MaxLength       =   39
         TabIndex        =   104
         Top             =   1320
         Width           =   2565
      End
      Begin VB.TextBox Textglyearopeningbalance 
         Height          =   285
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   103
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox Textfindgl 
         Height          =   345
         Left            =   3120
         TabIndex        =   102
         Top             =   180
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox Cashbankbook 
         Alignment       =   1  'Right Justify
         Caption         =   "Cash / Bank A/C "
         Height          =   255
         Left            =   360
         TabIndex        =   101
         Top             =   3120
         Width           =   2955
      End
      Begin VB.Label Label9 
         Caption         =   "Year Opening Balance"
         Height          =   585
         Left            =   420
         TabIndex        =   113
         Top             =   2790
         Width           =   2985
      End
      Begin VB.Label Label1 
         Caption         =   "Specify Category"
         Height          =   225
         Left            =   405
         TabIndex        =   112
         Top             =   840
         Width           =   2940
      End
      Begin VB.Label Label2 
         Height          =   585
         Left            =   600
         TabIndex        =   111
         Top             =   1770
         Width           =   2865
      End
      Begin VB.Label Label4 
         Caption         =   "General Ledger Description"
         Height          =   255
         Left            =   405
         TabIndex        =   110
         Top             =   1320
         Width           =   2955
      End
      Begin VB.Label Label7 
         Height          =   585
         Left            =   600
         TabIndex        =   109
         Top             =   2070
         Width           =   2865
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid VS 
      Height          =   4575
      Left            =   120
      TabIndex        =   99
      Top             =   4680
      Width           =   9195
      _cx             =   16219
      _cy             =   8070
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   600
      Left            =   10590
      TabIndex        =   32
      Top             =   3240
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   1058
      _Version        =   393216
      Tabs            =   6
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "&General Ledger "
      TabPicture(0)   =   "masters.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Su&B Ledger "
      TabPicture(1)   =   "masters.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "sledger"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Invoice End Part "
      TabPicture(2)   =   "masters.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "invnoteend"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "&Credit Note End Part "
      TabPicture(3)   =   "masters.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "crenoteend"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "&Discount Category"
      TabPicture(4)   =   "masters.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "discount"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Counter Sale End Part"
      TabPicture(5)   =   "masters.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cashend"
      Tab(5).ControlCount=   1
      Begin VB.Frame sledger 
         Height          =   4245
         Left            =   330
         TabIndex        =   33
         Top             =   1080
         Width           =   8325
         Begin VB.TextBox txtphoneno 
            Height          =   285
            Left            =   3150
            MaxLength       =   49
            TabIndex        =   97
            Top             =   2730
            Width           =   3135
         End
         Begin VB.TextBox Textsldiscriptionforinvoice 
            Height          =   285
            Left            =   3150
            MaxLength       =   39
            TabIndex        =   95
            Top             =   1410
            Width           =   4275
         End
         Begin VB.ComboBox Combosldistrictcode 
            Height          =   315
            Left            =   3150
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   720
            Width           =   4035
         End
         Begin VB.ComboBox CBODISTCODE 
            Height          =   315
            Left            =   2310
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2130
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.TextBox Textsladdress3 
            Height          =   285
            Left            =   3150
            MaxLength       =   49
            TabIndex        =   8
            Top             =   2400
            Width           =   3135
         End
         Begin VB.TextBox Textsladdress2 
            Height          =   285
            Left            =   3150
            MaxLength       =   49
            TabIndex        =   7
            Top             =   2070
            Width           =   3135
         End
         Begin VB.TextBox Textslfindgl 
            Height          =   285
            Left            =   420
            TabIndex        =   53
            Top             =   -60
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.TextBox TextFINDSUBLEADGER 
            Height          =   285
            Left            =   4170
            TabIndex        =   34
            Top             =   -60
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.TextBox Textslyearopeningbalance 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3150
            MaxLength       =   15
            TabIndex        =   9
            Top             =   3060
            Width           =   3135
         End
         Begin VB.TextBox Textsladdress1 
            Height          =   285
            Left            =   3150
            MaxLength       =   49
            TabIndex        =   6
            Top             =   1740
            Width           =   3135
         End
         Begin VB.TextBox Textslsubledgerdiscription 
            Height          =   285
            Left            =   3810
            MaxLength       =   49
            TabIndex        =   5
            Top             =   1080
            Width           =   3615
         End
         Begin VB.ComboBox Combosldistrictcode1 
            Height          =   315
            Left            =   6870
            TabIndex        =   12
            Top             =   3090
            Visible         =   0   'False
            Width           =   3165
         End
         Begin VB.ComboBox Combosldiscountcategory 
            Height          =   315
            Left            =   3150
            Style           =   1  'Simple Combo
            TabIndex        =   11
            Top             =   3375
            Width           =   3165
         End
         Begin VB.ComboBox Comboslgenledgerdiscription 
            Height          =   315
            Left            =   3150
            Sorted          =   -1  'True
            TabIndex        =   2
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label35 
            Caption         =   "Phone No :"
            Height          =   195
            Left            =   570
            TabIndex        =   98
            Top             =   2790
            Width           =   1935
         End
         Begin VB.Label Label16 
            Caption         =   "(Description for Invoice)"
            Height          =   315
            Left            =   630
            TabIndex        =   96
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label txtdistcode 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3150
            TabIndex        =   94
            Top             =   1080
            Width           =   645
         End
         Begin VB.Label Label20 
            Caption         =   "District Name "
            Height          =   255
            Left            =   600
            TabIndex        =   93
            Top             =   690
            Width           =   2655
         End
         Begin VB.Label TXTCUSTCODE 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2340
            TabIndex        =   92
            Top             =   1770
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label Label34 
            Caption         =   "TIN Number"
            Height          =   255
            Left            =   600
            TabIndex        =   91
            Top             =   3450
            Width           =   2535
         End
         Begin VB.Label Label3 
            Caption         =   "Year Opening Balance"
            Height          =   255
            Left            =   570
            TabIndex        =   39
            Top             =   3120
            Width           =   1845
         End
         Begin VB.Label Label6 
            Caption         =   "General Ledger Description"
            Height          =   255
            Left            =   600
            TabIndex        =   38
            Top             =   390
            Width           =   2595
         End
         Begin VB.Label Label13 
            Caption         =   "Address"
            Height          =   375
            Left            =   630
            TabIndex        =   37
            Top             =   1710
            Width           =   2055
         End
         Begin VB.Label Label15 
            Caption         =   "Sub. Ledger Discription"
            Height          =   375
            Left            =   600
            TabIndex        =   36
            Top             =   1110
            Width           =   2295
         End
         Begin VB.Label Label19 
            Caption         =   "Discount Category Code"
            Height          =   255
            Left            =   600
            TabIndex        =   35
            Top             =   3090
            Visible         =   0   'False
            Width           =   2535
         End
      End
      Begin VB.Frame cashend 
         Height          =   3765
         Left            =   -74670
         TabIndex        =   73
         Top             =   810
         Width           =   7965
         Begin VB.TextBox cashTextInvePrintOrder 
            Height          =   315
            Left            =   3390
            MaxLength       =   6
            TabIndex        =   69
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox cashTextcnep20chartext 
            Height          =   345
            Left            =   3360
            TabIndex        =   79
            Top             =   2490
            Width           =   3135
         End
         Begin VB.ComboBox cashCombocnepcontragenledgerdesc 
            Height          =   315
            Left            =   3390
            Sorted          =   -1  'True
            TabIndex        =   74
            Top             =   750
            Width           =   3165
         End
         Begin VB.ComboBox cashCombocnepcontrasubledgerdesc 
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   75
            Top             =   1200
            Width           =   3165
         End
         Begin VB.ComboBox cashCombocnepgenledgerdesc 
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   77
            Top             =   1620
            Width           =   3165
         End
         Begin VB.ComboBox cashCombocnepsubledgerdesc 
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   78
            Top             =   2040
            Width           =   3165
         End
         Begin VB.ComboBox cashCombocnepdrorcr 
            Height          =   315
            ItemData        =   "masters.frx":00B4
            Left            =   3360
            List            =   "masters.frx":00BE
            Sorted          =   -1  'True
            TabIndex        =   81
            Top             =   3390
            Width           =   1365
         End
         Begin VB.TextBox cashTextcneprate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   3360
            TabIndex        =   80
            Top             =   2940
            Width           =   3135
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Print Order No"
            Height          =   195
            Left            =   510
            TabIndex        =   89
            Top             =   420
            Width           =   1005
         End
         Begin VB.Label Label28 
            Caption         =   "Rate % (if any)"
            Height          =   255
            Left            =   480
            TabIndex        =   87
            Top             =   2940
            Width           =   2655
         End
         Begin VB.Label Label26 
            Caption         =   "Debit/Credit"
            Height          =   255
            Left            =   480
            TabIndex        =   86
            Top             =   3420
            Width           =   3015
         End
         Begin VB.Label Label25 
            Caption         =   "Gen.ledger Desc."
            Height          =   375
            Left            =   480
            TabIndex        =   85
            Top             =   1620
            Width           =   2295
         End
         Begin VB.Label Label24 
            Caption         =   "Contra Sub. Ledger Desc."
            Height          =   375
            Left            =   450
            TabIndex        =   84
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label Label23 
            Caption         =   "Sub.ledger Desc."
            Height          =   375
            Left            =   480
            TabIndex        =   83
            Top             =   2100
            Width           =   2295
         End
         Begin VB.Label Label22 
            Caption         =   "Contra Gen.l Ledger Desc."
            Height          =   255
            Left            =   450
            TabIndex        =   82
            Top             =   750
            Width           =   2955
         End
         Begin VB.Label Label5 
            Caption         =   "20 char.Text-->"
            Height          =   225
            Left            =   480
            TabIndex        =   76
            Top             =   2550
            Width           =   2985
         End
      End
      Begin VB.Frame invnoteend 
         Height          =   4395
         Left            =   -74445
         TabIndex        =   57
         Top             =   900
         Width           =   7605
         Begin VB.TextBox TextInvePrintOrder 
            Height          =   315
            Left            =   3360
            MaxLength       =   6
            TabIndex        =   55
            Top             =   300
            Width           =   1335
         End
         Begin VB.TextBox Textinvep20chartext 
            Height          =   285
            Left            =   3390
            TabIndex        =   65
            Top             =   2490
            Width           =   3135
         End
         Begin VB.ComboBox Comboinvepcontragenledgerdesc 
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   58
            Top             =   720
            Width           =   3165
         End
         Begin VB.ComboBox Comboinvepcontrasubledgerdesc 
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   59
            Top             =   1080
            Width           =   3165
         End
         Begin VB.ComboBox Comboinvepgenledgerdesc 
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   61
            Top             =   1560
            Width           =   3165
         End
         Begin VB.ComboBox Comboinvepsubledgerdesc 
            Height          =   315
            Left            =   3390
            Sorted          =   -1  'True
            TabIndex        =   63
            Top             =   2010
            Width           =   3165
         End
         Begin VB.ComboBox Comboinvepdrorcr 
            Height          =   315
            ItemData        =   "masters.frx":00D1
            Left            =   3360
            List            =   "masters.frx":00DB
            Sorted          =   -1  'True
            TabIndex        =   70
            Top             =   3330
            Width           =   1365
         End
         Begin VB.TextBox Textinveprate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3360
            TabIndex        =   67
            Top             =   2910
            Width           =   3135
         End
         Begin VB.Label Label29 
            Caption         =   "Print Order"
            Height          =   255
            Left            =   480
            TabIndex        =   88
            Top             =   300
            Width           =   1245
         End
         Begin VB.Label Label21 
            Caption         =   "Rate % (if any)"
            Height          =   255
            Left            =   480
            TabIndex        =   72
            Top             =   2910
            Width           =   2655
         End
         Begin VB.Label Label18 
            Caption         =   "Debit/Credit"
            Height          =   255
            Left            =   510
            TabIndex        =   71
            Top             =   3330
            Width           =   3015
         End
         Begin VB.Label Label17 
            Caption         =   "Gen.ledger Desc."
            Height          =   375
            Left            =   480
            TabIndex        =   68
            Top             =   1590
            Width           =   2295
         End
         Begin VB.Label Label12 
            Caption         =   "Contra Sub. Ledger Desc."
            Height          =   375
            Left            =   450
            TabIndex        =   66
            Top             =   1110
            Width           =   2295
         End
         Begin VB.Label Label11 
            Caption         =   "Sub.ledger Desc."
            Height          =   375
            Left            =   480
            TabIndex        =   64
            Top             =   2040
            Width           =   2295
         End
         Begin VB.Label Label10 
            Caption         =   "Contra Gen.l Ledger Desc."
            Height          =   255
            Left            =   450
            TabIndex        =   62
            Top             =   690
            Width           =   2955
         End
         Begin VB.Label Label8 
            Caption         =   "20 char.Text-->"
            Height          =   225
            Left            =   480
            TabIndex        =   60
            Top             =   2520
            Width           =   2985
         End
      End
      Begin VB.Frame discount 
         Height          =   3495
         Left            =   -73710
         TabIndex        =   48
         Top             =   1080
         Width           =   6885
         Begin VB.TextBox textfinddiscgroupcode 
            Height          =   285
            Left            =   3360
            TabIndex        =   56
            Top             =   210
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.TextBox Textfinddiscountcategory 
            Height          =   285
            Left            =   90
            TabIndex        =   54
            Top             =   210
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.ComboBox Combobgroupcode 
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   21
            Top             =   930
            Width           =   3165
         End
         Begin VB.ComboBox Combobgroupname 
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   22
            Top             =   1380
            Width           =   3165
         End
         Begin VB.TextBox Textdcdiscountrate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   3360
            TabIndex        =   23
            Top             =   1890
            Width           =   3135
         End
         Begin VB.TextBox Textdcdiscountcategorycode 
            Height          =   285
            Left            =   3360
            MaxLength       =   7
            TabIndex        =   20
            Top             =   570
            Width           =   3135
         End
         Begin VB.Label Label33 
            Caption         =   "Group Name"
            Height          =   255
            Left            =   570
            TabIndex        =   52
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label Label32 
            Caption         =   "Group Code"
            Height          =   255
            Left            =   570
            TabIndex        =   51
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label Label27 
            Caption         =   "Discount category code"
            Height          =   255
            Left            =   570
            TabIndex        =   50
            Top             =   540
            Width           =   2385
         End
         Begin VB.Label Label14 
            Caption         =   "Discount Rate"
            Height          =   255
            Left            =   540
            TabIndex        =   49
            Top             =   1980
            Width           =   2625
         End
      End
      Begin VB.Frame crenoteend 
         Height          =   3855
         Left            =   -74400
         TabIndex        =   40
         Top             =   990
         Width           =   7935
         Begin VB.TextBox CneTextInvePrintOrder 
            Height          =   315
            Left            =   3390
            MaxLength       =   6
            TabIndex        =   10
            Top             =   300
            Width           =   1335
         End
         Begin VB.TextBox Textcneprate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   3360
            TabIndex        =   18
            Top             =   2940
            Width           =   3135
         End
         Begin VB.ComboBox Combocnepdrorcr 
            Height          =   315
            ItemData        =   "masters.frx":00EE
            Left            =   3360
            List            =   "masters.frx":00F8
            Sorted          =   -1  'True
            TabIndex        =   19
            Top             =   3360
            Width           =   1365
         End
         Begin VB.ComboBox Combocnepsubledgerdesc 
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   2040
            Width           =   3165
         End
         Begin VB.ComboBox Combocnepgenledgerdesc 
            Height          =   315
            Left            =   3360
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   1620
            Width           =   3165
         End
         Begin VB.ComboBox Combocnepcontrasubledgerdesc 
            Height          =   315
            Left            =   3390
            Sorted          =   -1  'True
            TabIndex        =   14
            Top             =   1170
            Width           =   3135
         End
         Begin VB.ComboBox Combocnepcontragenledgerdesc 
            Height          =   315
            Left            =   3390
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   720
            Width           =   3105
         End
         Begin VB.TextBox Textcnep20chartext 
            Height          =   285
            Left            =   3360
            TabIndex        =   17
            Top             =   2490
            Width           =   3135
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Print Order No"
            Height          =   195
            Left            =   420
            TabIndex        =   90
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label45 
            Caption         =   "20 char.Text-->"
            Height          =   225
            Left            =   480
            TabIndex        =   47
            Top             =   2550
            Width           =   2985
         End
         Begin VB.Label Label47 
            Caption         =   "Contra Gen.l Ledger Desc."
            Height          =   255
            Left            =   450
            TabIndex        =   46
            Top             =   810
            Width           =   2955
         End
         Begin VB.Label Label49 
            Caption         =   "Sub.ledger Desc."
            Height          =   375
            Left            =   480
            TabIndex        =   45
            Top             =   2100
            Width           =   2295
         End
         Begin VB.Label Label50 
            Caption         =   "Contra Sub. Ledger Desc."
            Height          =   375
            Left            =   450
            TabIndex        =   44
            Top             =   1230
            Width           =   2295
         End
         Begin VB.Label Label51 
            Caption         =   "Gen.ledger Desc."
            Height          =   375
            Left            =   480
            TabIndex        =   43
            Top             =   1650
            Width           =   2295
         End
         Begin VB.Label Label53 
            Caption         =   "Debit/Credit"
            Height          =   255
            Left            =   480
            TabIndex        =   42
            Top             =   3390
            Width           =   3015
         End
         Begin VB.Label Label54 
            Caption         =   "Rate % (if any)"
            Height          =   255
            Left            =   480
            TabIndex        =   41
            Top             =   2940
            Width           =   2655
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   675
      Left            =   390
      ScaleHeight     =   615
      ScaleWidth      =   7845
      TabIndex        =   31
      Top             =   3840
      Width           =   7905
      Begin VB.CommandButton CommandmasterReturn 
         Caption         =   "&Return"
         Height          =   525
         Left            =   6705
         TabIndex        =   29
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton CommandmasterPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Height          =   525
         Left            =   5640
         TabIndex        =   28
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton Commandmastersearch 
         Caption         =   "&Search"
         Enabled         =   0   'False
         Height          =   525
         Left            =   5835
         TabIndex        =   27
         Top             =   1170
         Width           =   975
      End
      Begin VB.CommandButton Commandmasterdelete 
         Caption         =   "De&lete"
         Enabled         =   0   'False
         Height          =   525
         Left            =   4605
         TabIndex        =   26
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton Commandmasterabandon 
         Caption         =   "Aba&ndon"
         Enabled         =   0   'False
         Height          =   525
         Left            =   3480
         TabIndex        =   25
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton Commandmastersave 
         Caption         =   "Sa&ve"
         Enabled         =   0   'False
         Height          =   525
         Left            =   2415
         TabIndex        =   24
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton Commandmasteredit 
         Caption         =   "&Edit"
         Enabled         =   0   'False
         Height          =   525
         Left            =   1305
         TabIndex        =   1
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton Commandmasteradd 
         Caption         =   "&Add"
         Height          =   525
         Left            =   180
         TabIndex        =   0
         Top             =   30
         Width           =   975
      End
      Begin VB.CommandButton Commandmasterhelp 
         Caption         =   "Help"
         Height          =   345
         Left            =   -45
         TabIndex        =   30
         Top             =   615
         Visible         =   0   'False
         Width           =   800
      End
   End
End
Attribute VB_Name = "master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim CON As ADODB.Connection
Dim rs As ADODB.Recordset
Dim ctl As Control
Public addmaster As Boolean
Dim editing  As Boolean
Dim INVEVar As Integer

Private Sub cashCombocnepcontragenledgerdesc_Change()
Me.cashCombocnepcontrasubledgerdesc.Text = ""
End Sub

Private Sub cashCombocnepcontragenledgerdesc_Click()
  Me.cashCombocnepcontrasubledgerdesc.Text = ""
    If cashCombocnepcontragenledgerdesc.Text <> "" Then
        Set rs = New ADODB.Recordset
        rs.Open "select * from sledger where gledger='" + Trim(cashCombocnepcontragenledgerdesc.Text) + "' and " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
        Me.cashCombocnepcontrasubledgerdesc.Clear
        If Not rs.BOF Then
            rs.MoveFirst
            Do While Not rs.EOF
                Me.cashCombocnepcontrasubledgerdesc.AddItem rs(1)
                If Not rs.EOF Then
                    rs.MoveNext
                End If
            Loop
        End If
        rs.Close
    End If
End Sub

Private Sub cashCombocnepcontragenledgerdesc_LostFocus()
If cashCombocnepcontragenledgerdesc.Text = "" Then
    'MsgBox "Enter Gen.ledger"
    'cashCombocnepcontragenledgerdesc.SetFocus
End If
If cashCombocnepcontragenledgerdesc.Text <> "" Then
        cashCombocnepcontragenledgerdesc.Text = UCase(cashCombocnepcontragenledgerdesc.Text)

        Set rs = New ADODB.Recordset
        rs.Open "select gledger from gledger where gledger='" + Trim(cashCombocnepcontragenledgerdesc.Text) + "'  and " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
        If rs.EOF Then
            MsgBox cashCombocnepcontragenledgerdesc.Text + " Ledger not found"
            cashCombocnepcontragenledgerdesc.SetFocus
            'Exit Sub
        End If
        rs.Close
        rs.Open "select * from sledger where gledger='" + Trim(cashCombocnepcontragenledgerdesc.Text) + "' and " & stridnyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
        Me.cashCombocnepcontrasubledgerdesc.Clear
        If Not rs.BOF Then
            Me.cashCombocnepcontrasubledgerdesc.Enabled = True
            rs.MoveFirst
            Do While Not rs.EOF
                Me.cashCombocnepcontrasubledgerdesc.AddItem rs(1)
                If Not rs.EOF Then
                   
                    rs.MoveNext
                End If
            Loop
            cashCombocnepcontrasubledgerdesc.SetFocus
         Else
            Me.cashCombocnepcontrasubledgerdesc.Enabled = False
            
        End If
        rs.Close
    End If
    
    
    
    
    
End Sub

Private Sub cashCombocnepcontrasubledgerdesc_LostFocus()
'If Me.cashCombocnepcontragenledgerdesc <> "" Then
'        If cashCombocnepcontrasubledgerdesc.Text <> "" Then
'            Set rs = New ADODB.Recordset
'            rs.Open "select gledger,subledger from sledger where subledger='" + Trim(cashCombocnepcontrasubledgerdesc.Text) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
'            If rs.EOF Then
'                MsgBox "" + Me.cashCombocnepcontrasubledgerdesc + " Not Found in Sub Ledger's"
'                Me.cashCombocnepcontrasubledgerdesc.SetFocus
'            Else
'                If Me.cashCombocnepcontragenledgerdesc.Text <> rs(0) Then
'                    MsgBox "" + Me.cashCombocnepcontrasubledgerdesc.Text + " is not the subledger of " + Me.cashCombocnepcontragenledgerdesc.Text + ""
'                    Me.cashCombocnepcontrasubledgerdesc.SetFocus
'                End If
'            End If
'            rs.Close
'        End If
'Else
'            Set rs = New ADODB.Recordset
'            rs.Open "select gledger,subledger from sledger where subledger='" + Trim(cashCombocnepcontrasubledgerdesc.Text) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
'            If Not rs.EOF Then
'                Me.cashCombocnepcontragenledgerdesc.Text = rs(0)
'            End If
'            rs.Close
'End If



Dim rs As New ADODB.Recordset
      
If cashCombocnepcontragenledgerdesc <> "" And cashCombocnepcontrasubledgerdesc.ListCount > 0 And cashCombocnepcontrasubledgerdesc.Text = "" Then cashCombocnepcontrasubledgerdesc.SetFocus
  If cashCombocnepsubledgerdesc.ListCount > 0 And cashCombocnepcontrasubledgerdesc.Text <> "" Then
      rs.Open "Select* from sledger where GLEDGER='" + Trim(cashCombocnepcontragenledgerdesc.Text) + "' and SubLedger='" + Trim(cashCombocnepcontrasubledgerdesc.Text) + "' and " & stridnyear, CON, adOpenStatic
      If rs.RecordCount <= 0 Then
           MsgBox "No valid Sub Ledger"
           cashCombocnepcontrasubledgerdesc.SetFocus
      End If
End If


End Sub

Private Sub cashCombocnepdrorcr_LostFocus()
If cashCombocnepdrorcr.Text <> "Debit" And cashCombocnepdrorcr.Text <> "Credit" Then
  MsgBox "Please Enter Debit/Credit.."
 cashCombocnepdrorcr.SetFocus
End If
End Sub

Private Sub cashCombocnepgenledgerdesc_LostFocus()
If cashCombocnepgenledgerdesc.Text = "" Then
    'MsgBox "Enter Gen.ledger"
    'cashCombocnepgenledgerdesc.SetFocus
    
End If
  If cashCombocnepgenledgerdesc.Text <> "" Then
  cashCombocnepgenledgerdesc.Text = UCase(cashCombocnepgenledgerdesc.Text)
        Set rs = New ADODB.Recordset
        rs.Open "select gledger from gledger where gledger='" + Trim(cashCombocnepgenledgerdesc.Text) + "' and " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
        If rs.EOF Then
            MsgBox cashCombocnepcontragenledgerdesc.Text + " Ledger not found"
            cashCombocnepcontragenledgerdesc.SetFocus
            'Exit Sub
        End If
        rs.Close
        rs.Open "select * from sledger where gledger='" + Trim(cashCombocnepgenledgerdesc.Text) + "' and " & stridnyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
        Me.cashCombocnepsubledgerdesc.Clear
        If Not rs.BOF Then
            Me.cashCombocnepsubledgerdesc.Enabled = True
            rs.MoveFirst
            Do While Not rs.EOF
                Me.cashCombocnepsubledgerdesc.AddItem rs(1)
                If Not rs.EOF Then
                    rs.MoveNext
                End If
            Loop
            cashCombocnepsubledgerdesc.SetFocus
        Else
            Me.cashCombocnepsubledgerdesc.Enabled = False
        End If
        rs.Close
    End If









'Me.cashCombocnepsubledgerdesc.Text = ""
End Sub


Private Sub cashCombocnepsubledgerdesc_LostFocus()
' If Me.cashCombocnepgenledgerdesc <> "" Then
'        If cashCombocnepsubledgerdesc.Text <> "" Then
'            Set rs = New ADODB.Recordset
'            rs.Open "select gledger,subledger from sledger where subledger='" + Trim(cashCombocnepsubledgerdesc.Text) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
'            If rs.EOF Then
'                MsgBox "" + Me.Combocnepsubledgerdesc + " Not Found in Sub Ledger's"
'                Me.cashCombocnepsubledgerdesc.SetFocus
'            Else
'                If Me.cashCombocnepgenledgerdesc.Text <> rs(0) Then
'                    MsgBox "" + Me.cashCombocnepsubledgerdesc.Text + " is not the subledger of " + Me.cashCombocnepgenledgerdesc.Text + ""
'                    Me.cashCombocnepsubledgerdesc.SetFocus
'                End If
'            End If
'            rs.Close
'        End If
'    Else
'            Set rs = New ADODB.Recordset
'            rs.Open "select gledger,subledger from sledger where subledger='" + Trim(cashCombocnepsubledgerdesc.Text) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
'            If Not rs.EOF Then
'                Me.cashCombocnepgenledgerdesc.Text = rs(0)
'            End If
'            rs.Close
'    End If


If cashCombocnepgenledgerdesc <> "" And cashCombocnepsubledgerdesc.ListCount > 0 And Combocnepsubledgerdesc.Text = "" Then cashCombocnepsubledgerdesc.SetFocus
If cashCombocnepsubledgerdesc.ListCount > 0 And cashCombocnepsubledgerdesc.Text <> "" Then
    rs.Open "Select* from sledger where GLEDGER='" + Trim(cashCombocnepgenledgerdesc.Text) + "' and SubLedger='" + Trim(cashCombocnepsubledgerdesc.Text) + "' and " & stridnyear, CON, adOpenStatic
    If rs.RecordCount <= 0 Then
       MsgBox "No valid Sub Ledger"
       cashCombocnepsubledgerdesc.SetFocus
  End If
  

    
End If





End Sub

Private Sub cashTextcnep20chartext_LostFocus()
cashTextcnep20chartext.Text = UCase(cashTextcnep20chartext.Text)

End Sub

Private Sub cashTextInvePrintOrder_LostFocus()
If IsNumeric(cashTextInvePrintOrder.Text) = False Then
    MsgBox "Please Enter Any No..."
    cashTextInvePrintOrder.SetFocus
End If

End Sub

Private Sub CBODISTCODE_Change()
CBODISTCODE_Click
End Sub

Private Sub CBODISTCODE_Click()
On Error Resume Next
If addmaster = True Then
    Dim temp As ADODB.Recordset
    Set temp = New ADODB.Recordset
    temp.Open "select MAX(CONVERT(INT,SUBSTRING(SUBLEDGER,CHARINDEX('-',SUBLEDGER,1)+1,3))) AS MAXID from SLEDGER where " & stridnyear & "  AND SUBLEDGER LIKE '" & UCase(CBODISTCODE.Text) & "%'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If temp.EOF = False Then
        If temp!MAXID > 0 Then
            TXTCUSTCODE.Caption = Format(temp!MAXID + 1, "000")
        Else
            TXTCUSTCODE = "001"
        End If
    Else
        TXTCUSTCODE = "001"
    End If
    temp.Close
End If

If rs.State = adStateOpen Then rs.Close
rs.Open "SELECT DISTRICTNAME FROM DISTRICTS WHERE DISTCODE='" & CBODISTCODE.Text & "' AND " & stridnyear
If rs.EOF = False Then
Combosldistrictcode.Text = rs!DISTRICTNAME
End If
End Sub

Private Sub CneTextInvePrintOrder_LostFocus()
If IsNumeric(CneTextInvePrintOrder.Text) = False Then
    MsgBox "Please Enter Any No..."
    CneTextInvePrintOrder.SetFocus
End If
End Sub

Private Sub Combobgroupcode_Change()
    Dim temp As ADODB.Recordset
    Set temp = New ADODB.Recordset
    temp.Open "select * from groups where " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not temp.EOF Then
        temp.Find "groupcode='" + Trim(Me.Combobgroupcode.Text) + "'"
        If Not temp.EOF Then
          Me.Combobgroupname.Text = temp(1)
        End If
    End If
    temp.Close
End Sub
Private Sub Combobgroupcode_Click()
    Dim temp As ADODB.Recordset
    Set temp = New ADODB.Recordset
    temp.Open "select * from groups where  " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not temp.EOF Then
        temp.Find "groupcode='" + Trim(Me.Combobgroupcode.Text) + "'"
        If Not temp.EOF Then
          Me.Combobgroupname.Text = temp(1)
        End If
    End If
    temp.Close
End Sub
Private Sub Combobgroupcode_LostFocus()
    rs.Open "select * from groups where  " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        rs.Find "groupcode='" + Trim(Me.Combobgroupcode.Text) + "'"
        If Not rs.EOF Then
            Me.Combobgroupname.Text = rs(1)
        End If
    End If
    rs.Close
End Sub
Private Sub Combobgroupname_Change()
    Dim temp As ADODB.Recordset
    Set temp = New ADODB.Recordset
    temp.Open "select * from groups where  " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not temp.EOF Then
        temp.Find "groupname='" + Trim(Me.Combobgroupname.Text) + "'"
        If Not temp.EOF Then
            Me.Combobgroupcode.Text = temp(0)
        End If
    End If
    temp.Close
End Sub

Private Sub Combobgroupname_Click()
    Dim temp As ADODB.Recordset
    Set temp = New ADODB.Recordset
    temp.Open "select * from groups where " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not temp.EOF Then
        temp.Find "groupname='" + Trim(Me.Combobgroupname.Text) + "'"
        If Not temp.EOF Then
            Me.Combobgroupcode.Text = temp(0)
        End If
    End If
    temp.Close
End Sub

Private Sub Combobgroupname_LostFocus()
    rs.Open "select * from groups where  " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        rs.Find "groupname='" + Trim(Me.Combobgroupname.Text) + "'"
        If Not rs.EOF Then
            Me.Combobgroupcode.Text = rs(0)
        End If
    End If
    rs.Close
End Sub


Private Sub Combocnepcontragenledgerdesc_Change()
    Me.Combocnepcontrasubledgerdesc.Text = ""
End Sub

Private Sub Combocnepcontragenledgerdesc_Click()
    Me.Combocnepcontrasubledgerdesc.Text = ""
    If Combocnepcontragenledgerdesc.Text <> "" Then
        Set rs = New ADODB.Recordset
        rs.Open "select * from sledger where gledger='" + Trim(Combocnepcontragenledgerdesc.Text) + "' and " & stridnyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
        Me.Combocnepcontrasubledgerdesc.Clear
        If Not rs.BOF Then
            rs.MoveFirst
            Do While Not rs.EOF
                Me.Combocnepcontrasubledgerdesc.AddItem rs(1)
                If Not rs.EOF Then
                    rs.MoveNext
                End If
            Loop
        End If
        rs.Close
    End If
End Sub

Private Sub Combocnepcontragenledgerdesc_LostFocus()
If Combocnepcontragenledgerdesc.Text = "" Then
    'Combocnepcontragenledgerdesc.SetFocus
End If
If Combocnepcontragenledgerdesc.Text <> "" Then
        Set rs = New ADODB.Recordset
        rs.Open "select gledger from gledger where gledger='" + Trim(Combocnepcontragenledgerdesc.Text) + "' and " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
        If rs.EOF Then
            MsgBox Combocnepcontragenledgerdesc.Text + " Ledger not found"
            Combocnepcontragenledgerdesc.SetFocus
            'Exit Sub
        End If
        rs.Close
        rs.Open "select * from sledger where gledger='" + Trim(Combocnepcontragenledgerdesc.Text) + "' and " & stridnyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
        Me.Combocnepcontrasubledgerdesc.Clear
        
        If Not rs.BOF Then
            Me.Combocnepcontrasubledgerdesc.Enabled = True
            rs.MoveFirst
            Do While Not rs.EOF
                Me.Combocnepcontrasubledgerdesc.AddItem rs(1)
                If Not rs.EOF Then
                    rs.MoveNext
                End If
            Loop
            Me.Combocnepcontrasubledgerdesc.SetFocus
        Else
        
          Me.Combocnepcontrasubledgerdesc.Enabled = False
            
            
        End If
        rs.Close
    End If
End Sub
Private Sub Combocnepcontrasubledgerdesc_LostFocus()
'If Me.Combocnepcontragenledgerdesc <> "" Then
'        If Combocnepcontrasubledgerdesc.Text <> "" Then
'            Set rs = New ADODB.Recordset
'            rs.Open "select gledger,subledger from sledger where subledger='" + Trim(Combocnepcontrasubledgerdesc.Text) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
'            If rs.EOF Then
'                MsgBox "" + Me.Combocnepcontrasubledgerdesc + " Not Found in Sub Ledger's"
'                Me.Combocnepcontrasubledgerdesc.SetFocus
'            Else
'                If Me.Combocnepcontragenledgerdesc.Text <> rs(0) Then
'                    MsgBox "" + Me.Combocnepcontrasubledgerdesc.Text + " is not the subledger of " + Me.Combocnepcontragenledgerdesc.Text + ""
'                    Me.Combocnepcontrasubledgerdesc.SetFocus
'                End If
'            End If
'            rs.Close
'        End If
'Else
'            Set rs = New ADODB.Recordset
'            rs.Open "select gledger,subledger from sledger where subledger='" + Trim(Combocnepcontrasubledgerdesc.Text) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
'            If Not rs.EOF Then
'                Me.Combocnepcontragenledgerdesc.Text = rs(0)
'            End If
'            rs.Close
'End If



Dim rs As New ADODB.Recordset
      
If Combocnepcontragenledgerdesc <> "" And Combocnepcontrasubledgerdesc.ListCount > 0 And Combocnepcontrasubledgerdesc.Text = "" Then Combocnepcontrasubledgerdesc.SetFocus
If Combocnepsubledgerdesc.ListCount > 0 And Combocnepcontrasubledgerdesc.Text <> "" Then
    
    rs.Open "Select* from sledger where GLEDGER='" + Trim(Combocnepcontragenledgerdesc.Text) + "' and SubLedger='" + Trim(Combocnepcontrasubledgerdesc.Text) + "' and " & stridnyear, CON, adOpenStatic
    If rs.RecordCount <= 0 Then
       MsgBox "No valid Sub Ledger"
       Combocnepcontrasubledgerdesc.SetFocus
  End If
  

    
End If

End Sub

Private Sub Combocnepdrorcr_LostFocus()
If Combocnepdrorcr.Text <> "Debit" And Combocnepdrorcr.Text <> "Credit" Then
         MsgBox "Please Enter Debit/Credit.."
         'Combocnepdrorcr.SetFocus
         
    End If
    
End Sub

Private Sub Combocnepgenledgerdesc_LostFocus()

   
   

 If Combocnepgenledgerdesc.Text <> "" Then
        Set rs = New ADODB.Recordset
        rs.Open "select gledger from gledger where  gledger='" + Trim(Combocnepgenledgerdesc.Text) + "' and " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
        If rs.EOF Then
            MsgBox Combocnepgenledgerdesc.Text + " Ledger not found"
            Combocnepgenledgerdesc.SetFocus
            'Exit Sub
        End If
        rs.Close
        rs.Open "select * from sledger where gledger='" + Trim(Combocnepgenledgerdesc.Text) + "' and " & stridnyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
        Me.Combocnepsubledgerdesc.Clear
        If Not rs.BOF Then
            Me.Combocnepsubledgerdesc.Enabled = True
            rs.MoveFirst
            Do While Not rs.EOF
                Me.Combocnepsubledgerdesc.AddItem rs(1)
                If Not rs.EOF Then
                    rs.MoveNext
                End If
                
            Loop
            Combocnepsubledgerdesc.SetFocus
        Else
            Me.Combocnepsubledgerdesc.Enabled = False
        End If
        rs.Close
    End If
 
 
 
 
End Sub
Private Sub Combocnepsubledgerdesc_LostFocus()
    
    
    
    
    
   ' If Me.Combocnepgenledgerdesc <> "" Then
   '     If Combocnepsubledgerdesc.Text <> "" Then
   '         Set rs = New ADODB.Recordset
   '         rs.Open "select gledger,subledger from sledger where subledger='" + Trim(Combocnepsubledgerdesc.Text) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
   '         If rs.EOF Then
   '             MsgBox "" + Me.Combocnepsubledgerdesc + " Not Found in Sub Ledger's"
   '             Me.Combocnepsubledgerdesc.SetFocus
   '         Else
   '             If Me.Combocnepgenledgerdesc.Text <> rs(0) Then
   '                 MsgBox "" + Me.Combocnepsubledgerdesc.Text + " is not the subledger of " + Me.Combocnepgenledgerdesc.Text + ""
   '                 Me.Combocnepsubledgerdesc.SetFocus
   '             End If
   '         End If
   '         rs.Close
  '
  '
  '      End If
  '  Else
  '          Set rs = New ADODB.Recordset
  '          rs.Open "select gledger,subledger from sledger where subledger='" + Trim(Combocnepsubledgerdesc.Text) + "'", con, adOpenKeyset, adLockReadOnly, adCmdText
  '          If Not rs.EOF Then
   '             Me.Combocnepgenledgerdesc.Text = rs(0)
   '         End If
   '         rs.Close
  '  End If
  
Dim rs As New ADODB.Recordset
      
If Combocnepgenledgerdesc <> "" And Combocnepsubledgerdesc.ListCount > 0 And Combocnepsubledgerdesc.Text = "" Then Combocnepsubledgerdesc.SetFocus
If Combocnepsubledgerdesc.ListCount > 0 And Combocnepsubledgerdesc.Text <> "" Then
    rs.Open "Select* from sledger where GLEDGER='" + Trim(Combocnepgenledgerdesc.Text) + "' and SubLedger='" + Trim(Combocnepsubledgerdesc.Text) + "' and " & stridnyear, CON, adOpenStatic
    If rs.RecordCount <= 0 Then
       MsgBox "No valid Sub Ledger"
       Combocnepsubledgerdesc.SetFocus
  End If
  

    
End If

  
End Sub
Private Sub Comboinvepcontragenledgerdesc_Click()
Me.Comboinvepcontrasubledgerdesc.Text = ""

If Comboinvepcontragenledgerdesc.Text <> "" Then
        Set rs = New ADODB.Recordset
        rs.Open "select * from sledger where gledger='" + Trim(Comboinvepcontragenledgerdesc.Text) + "' and " & stridnyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
        Me.Comboinvepcontrasubledgerdesc.Clear
        If Not rs.BOF Then
            rs.MoveFirst
            Do While Not rs.EOF
                Me.Comboinvepcontrasubledgerdesc.AddItem rs(1)
                If Not rs.EOF Then
                    rs.MoveNext
                End If
            Loop
        End If
        rs.Close
End If
End Sub
Private Sub Comboinvepcontragenledgerdesc_LostFocus()
If Comboinvepcontragenledgerdesc.Text = "" Then
    'MsgBox "Enter Gen. Ledger..."
    'Comboinvepcontragenledgerdesc.SetFocus
    
End If


    If Comboinvepcontragenledgerdesc.Text <> "" Then
        Me.Comboinvepcontrasubledgerdesc.Enabled = True
        Set rs = New ADODB.Recordset
        rs.Open "select gledger from gledger where gledger='" + Trim(Comboinvepcontragenledgerdesc.Text) + "' and " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
        If rs.EOF Then
            MsgBox Comboinvepcontragenledgerdesc.Text + " Ledger not found"
            Comboinvepcontragenledgerdesc.SetFocus
        End If
        rs.Close
        rs.Open "select * from sledger where gledger='" + Trim(Comboinvepcontragenledgerdesc.Text) + "' and " & stridnyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
        Me.Comboinvepcontrasubledgerdesc.Clear
        If Not rs.BOF Then
            rs.MoveFirst
            Me.Comboinvepcontrasubledgerdesc.Enabled = True
            Do While Not rs.EOF
                Me.Comboinvepcontrasubledgerdesc.AddItem rs(1)
                If Not rs.EOF Then
                    rs.MoveNext
                End If
            Loop
            Comboinvepcontrasubledgerdesc.SetFocus
        Else
            Me.Comboinvepcontrasubledgerdesc.Enabled = False
            
        End If
        rs.Close
    End If
End Sub

Private Sub Comboinvepcontrasubledgerdesc_LostFocus()
Dim rs As New ADODB.Recordset
      
If Comboinvepcontragenledgerdesc <> "" And Comboinvepcontrasubledgerdesc.ListCount > 0 And Comboinvepcontrasubledgerdesc.Text = "" Then Comboinvepcontrasubledgerdesc.SetFocus

If Comboinvepcontrasubledgerdesc.ListCount > 0 And Comboinvepcontrasubledgerdesc.Text <> "" Then
    rs.Open "Select* from sledger where GLEDGER='" + Trim(Comboinvepcontragenledgerdesc.Text) + "' and SubLedger='" + Trim(Comboinvepcontrasubledgerdesc.Text) + "'  and " & stridnyear, CON, adOpenStatic
    If rs.RecordCount <= 0 Then
       MsgBox "No valid Sub Ledger"
       Comboinvepcontrasubledgerdesc.SetFocus
  End If
  

    
End If
End Sub

Private Sub Comboinvepdrorcr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ' Commandmastersave.SetFocus
End If
End Sub

Private Sub Comboinvepdrorcr_LostFocus()
If Comboinvepdrorcr.Text <> "Debit" And Comboinvepdrorcr.Text <> "Credit" Then
  MsgBox "Please Enter Debit/Credit.."
  Comboinvepdrorcr.SetFocus
End If
End Sub

Private Sub Comboinvepgenledgerdesc_Click()
Me.Comboinvepsubledgerdesc.Text = ""
 Me.Comboinvepsubledgerdesc.Enabled = True
If Comboinvepgenledgerdesc.Text <> "" Then
        Set rs = New ADODB.Recordset
        rs.Open "select * from sledger where gledger='" + Trim(Comboinvepgenledgerdesc.Text) + "'  and " & stridnyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
        Me.Comboinvepsubledgerdesc.Clear
        If Not rs.BOF Then
            Me.Comboinvepsubledgerdesc.Enabled = True
            
            rs.MoveFirst
            Do While Not rs.EOF
                Me.Comboinvepsubledgerdesc.AddItem rs(1)
                If Not rs.EOF Then
                    rs.MoveNext
                End If
            Loop
            Comboinvepsubledgerdesc.SetFocus
        Else
            Me.Comboinvepsubledgerdesc.Enabled = False
        
            
        End If
        rs.Close
    End If
End Sub
Private Sub Comboinvepgenledgerdesc_LostFocus()
If Comboinvepgenledgerdesc.Text = "" Then
     'MsgBox "Enter Gen. Ledger..."
    ' Comboinvepgenledgerdesc.SetFocus
End If
     

    If Comboinvepgenledgerdesc.Text <> "" Then
        Set rs = New ADODB.Recordset
        rs.Open "select gledger from gledger where  gledger='" + Trim(Comboinvepgenledgerdesc.Text) + "'  and " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
        If rs.EOF Then
            MsgBox Comboinvepgenledgerdesc.Text + " Ledger not found"
            Comboinvepgenledgerdesc.SetFocus
            'Exit Sub
        End If
        rs.Close
        rs.Open "select * from sledger where gledger='" + Trim(Comboinvepgenledgerdesc.Text) + "' and " & stridnyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
        Me.Comboinvepsubledgerdesc.Clear
        If Not rs.BOF Then
            Me.Comboinvepsubledgerdesc.Enabled = True
            rs.MoveFirst
            Do While Not rs.EOF
                Me.Comboinvepsubledgerdesc.AddItem rs(1)
                If Not rs.EOF Then
                    rs.MoveNext
                End If
                
            Loop
        Else
            Me.Comboinvepsubledgerdesc.Enabled = False
        End If
        rs.Close
    End If
End Sub

Private Sub Comboinvepsubledgerdesc_LostFocus()
Dim rs As New ADODB.Recordset
If Comboinvepgenledgerdesc <> "" And Comboinvepsubledgerdesc.ListCount > 0 And Comboinvepsubledgerdesc.Text = "" Then
  'Comboinvepsubledgerdesc.Enabled = True
  Comboinvepsubledgerdesc.SetFocus
End If
If Comboinvepsubledgerdesc.ListCount > 0 And Comboinvepsubledgerdesc.Text <> "" Then
    rs.Open "Select* from sledger where GLEDGER='" + Trim(Comboinvepgenledgerdesc.Text) + "' and SubLedger='" + Trim(Comboinvepsubledgerdesc.Text) + "' and " & stridnyear, CON, adOpenStatic
    If rs.RecordCount <= 0 Then
       MsgBox "No valid Sub Ledger"
       Comboinvepsubledgerdesc.SetFocus
  End If
End If

End Sub

Private Sub Combosldiscountcategory_LostFocus()

'''''    If Combosldiscountcategory = "" And Comboslgenledgerdiscription.Text = UCase("sundry debtors") Then
'''''        Combosldiscountcategory.SetFocus
'''''        Exit Sub
'''''    End If

''    If Combosldiscountcategory.Text <> "" Then
''        Combosldiscountcategory.Text = UCase(Combosldiscountcategory.Text)
''        Set RS = New ADODB.Recordset
''        RS.Open "select * from disccats  where categorycode ='" + Trim(Combosldiscountcategory.Text) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
''        If RS.EOF Then
''           MsgBox "Not valid Discount Category.."
''           Combosldiscountcategory.SetFocus
''        End If
''        RS.Close
''    End If
    
End Sub


Private Sub Combosldistrictcode_Click()
'On Error Resume Next
If addmaster = True And Combosldistrictcode.Text <> "" Then
    Dim temp As ADODB.Recordset
    Set temp = New ADODB.Recordset
    'temp.Open "select MAX(CONVERT(INT,SUBSTRING(SUBLEDGER,CHARINDEX('-',SUBLEDGER,1)+1,3))) AS MAXID from SLEDGER where " & stridnyear & "  AND SUBLEDGER LIKE '" & UCase(Combosldistrictcode.Text) & "%'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    temp.Open "Select distcode from DISTRICTS where " & stridnyear & " and  DISTRICTNAME='" & Combosldistrictcode.Text & "'", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If temp.EOF = False Then
         txtdistcode.Caption = temp!distcode
    End If
    temp.Close
    End If

'If rs.State = adStateOpen Then rs.Close
'rs.Open "SELECT DISTRICTNAME FROM DISTRICTS WHERE DISTCODE='" & CBODISTCODE.Text & "' AND " & stridnyear
'If rs.EOF = False Then
'C 'ombosldistrictcode.Text = rs!DISTRICTNAME
'End If

End Sub

Private Sub Combosldistrictcode_LostFocus()


'''''''''Dim rs3 As New ADODB.Recordset
''''''''' If Combosldiscountcategory = "" And Comboslgenledgerdiscription.Text = UCase("sundry debtors") Then
'''''''''        Combosldiscountcategory.SetFocus
'''''''''        Exit Sub
'''''''''    End If
'''''''''    If Combosldiscountcategory.Text <> "" Then
'''''''''        Combosldiscountcategory.Text = UCase(Combosldiscountcategory.Text)
'''''''''        rs3.Open "select * from disccats  where categorycode ='" + Trim(Combosldiscountcategory.Text) + "'", CON, adOpenStatic, adLockReadOnly, adCmdText
'''''''''        If rs3.EOF Then
'''''''''           MsgBox "Not valid Discount Category.."
'''''''''
'''''''''           Combosldiscountcategory.SetFocus
'''''''''           Exit Sub
'''''''''        End If
'''''''''        If rs3.State = 1 Then rs3.Close
'''''''''    End If
''''''''' If Combosldistrictcode <> "" And Comboslgenledgerdiscription.Text = UCase("sundry debtors") Then
'''''''''    rs3.Open "Select * from  Districts where  " & stridnyear & " and districtname = '" & Combosldistrictcode & "'", CON, adOpenStatic, adLockOptimistic
'''''''''    If rs3.RecordCount <= 0 Then
'''''''''        MsgBox "Not Valid District.."
'''''''''       Combosldistrictcode.SetFocus
''''''''' End If
'''''''''
'''''''''    'Combosldistrictcode.SetFocus
'''''''''    Exit Sub
''''''''' Else
'''''''''
'''''''''  If Comboslgenledgerdiscription.Text = UCase("sundry debtors") Then
'''''''''    rs3.Open "Select * from  Districts where  " & stridnyear & " and districtname = '" & Combosldistrictcode & "'", CON, adOpenStatic, adLockOptimistic
'''''''''    If rs3.RecordCount <= 0 Then
'''''''''        MsgBox "Not Valid District.."
'''''''''       Combosldistrictcode.SetFocus
''''''''' End If
''''''''' End If
'''''''''
'End If
End Sub

Private Sub Comboslgenledgerdiscription_LostFocus()
    If Len(Comboslgenledgerdiscription.Text) >= 40 Then
           MsgBox "Enter only 40 Character"
           Comboslgenledgerdiscription.SetFocus
           Exit Sub
    End If
    
    If Comboslgenledgerdiscription.Text <> "" Then
        Set rs = New ADODB.Recordset
        rs.Open "select gledger from gledger where slf= 1 and gledger='" + Trim(Comboslgenledgerdiscription.Text) + "' and " & stridnyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
        If rs.EOF Then
            MsgBox Comboslgenledgerdiscription.Text + " Ledger not found"
            Comboslgenledgerdiscription.SetFocus
        Else
            Comboslgenledgerdiscription.Text = rs!gledger
        End If
        rs.Close
    End If
    
End Sub

Private Sub ComboSPECIALCATEGORY_LostFocus()
'If ComboSPECIALCATEGORY.Text = "" Then MsgBox "Enter Category"
If Len(ComboSPECIALCATEGORY.Text) > 10 Then MsgBox "Enter only 10 Character"
End Sub

Private Sub Commandmasterabandon_Click()
    addmaster = False
    editing = False
    For i = 0 To 5
           SStab1.TabEnabled(i) = True
    Next
    For Each ctl In Me.Controls
        If TypeOf ctl Is textbox Then
            If ctl.Enabled = TURE Then
            ctl.Text = ""
            ctl.Enabled = False
            End If
        End If
        If TypeOf ctl Is ComboBox Then
            If ctl.Style <> 2 Then
            ctl.Text = ""
            ctl.Enabled = False
            End If
        End If
        
        If TypeOf ctl Is CheckBox Then
            ctl.Value = 0
            ctl.Enabled = False
        End If
        If TypeOf ctl Is ListBox Then
            ctl.Enabled = False
        End If
    Next
SetButton Commandmasteradd, Commandmasteredit, Commandmastersave, Commandmasterdelete
Commandmasteredit.Enabled = False
Commandmastersave.Enabled = False
Commandmasterdelete.Enabled = False


End Sub
Private Sub Commandmasteradd_Click()
     addmaster = True
     For Each ctl In Me.Controls
        If TypeOf ctl Is textbox Then
        If ctl.Enabled = True Then ctl.Text = ""
            'ctl.Enabled = False
        End If
        If TypeOf ctl Is ComboBox Then
        If ctl.Style <> 2 Then
        ctl.Text = ""
        End If
        End If
        If TypeOf ctl Is CheckBox Then
            ctl.Value = 0
            'ctl.Enabled = False
        End If
        If TypeOf ctl Is ListBox Then
            'ctl.Enabled = False
        End If
    Next
    
    
    If SStab1.Tab = 0 Then
    
    '/*  deactivate other tabs */
   
        For i = 0 To 5
            If i <> SStab1.Tab Then
                SStab1.TabEnabled(i) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.Container.Name) = UCase("gledger") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
       master.gledger.Enabled = True
       master.ComboSPECIALCATEGORY.Enabled = True
       Me.ComboSPECIALCATEGORY.SetFocus
     
    End If
    
    If SStab1.Tab = 1 Then
    '/**  deactivate other tabs**/
        For i = 0 To 5
            If i <> SStab1.Tab Then
                SStab1.TabEnabled(i) = False
            End If
        Next
        For Each ctl In Me.Controls
            If ctl.Container.Name = "sledger" Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
        master.sledger.Enabled = True
        Combosldistrictcode_Click
       
    End If
    If SStab1.Tab = 2 Then
    '/**  deactivate other tabs**/
        For i = 0 To 5
            If i <> SStab1.Tab Then
                SStab1.TabEnabled(i) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.Container.Name) = UCase("invnoteend") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
        master.invnoteend.Enabled = True
        'Comboinvepcontragenledgerdesc.SetFocus
        master.TextInvePrintOrder.SetFocus
        
        Textinveprate.Text = "0"
    End If
    
    
    If SStab1.Tab = 3 Then
    '/**  deactivate other tabs**/
        
       SStab1.Tab = 3
        
        For i = 0 To 5
            If i <> SStab1.Tab Then
                SStab1.TabEnabled(i) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.Container.Name) = UCase("crenoteend") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
                    ctl.Enabled = True
                End If
            'crenoteend.Enabled = True
            End If
        Next
    
       'Combocnepcontragenledgerdesc.SetFocus
       master.CneTextInvePrintOrder.SetFocus
       Textcneprate.Text = "0"
    End If
    
    
    If SStab1.Tab = 4 Then
    '/**  deactivate other tabs**/
        
        
        For i = 0 To 5
            If i <> SStab1.Tab Then
                SStab1.TabEnabled(i) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.Container.Name) = UCase("discount") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
    End If
    
    
    
If SStab1.Tab = 5 Then
    '/**  deactivate other tabs**/
        
       
        
        For i = 0 To 5
            If i <> SStab1.Tab Then
                SStab1.TabEnabled(i) = False
            End If
        Next
        For Each ctl In Me.Controls
            If UCase(ctl.Container.Name) = UCase("cashend") Then
                If TypeOf ctl Is textbox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
                    ctl.Enabled = True
                End If
            End If
        Next
    
       'cashCombocnepcontragenledgerdesc.SetFocus
       master.cashTextInvePrintOrder.SetFocus
       master.cashTextcneprate.Text = "0"
    End If
    

    
    
    
    
    
    Commandmasteradd.Enabled = False
    Commandmasteredit.Enabled = False
    CommandmasterPrint.Enabled = False
    Commandmastersave.Enabled = True
    Commandmasterabandon.Enabled = True
    CommandmasterReturn.Enabled = True
    Commandmastersearch.Enabled = True
    
End Sub
    
Private Sub Commandmasterdelete_Click()

If MsgBox("Want To Delete ?", vbInformation + vbYesNo) = vbYes Then
  CON.Execute "delete from GLEDGER where (Category='" & ComboSPECIALCATEGORY.Text & "' and gledger='" & Textglgeneralledgerdiscription.Text & "') and " & "" & stridnyear
  Commandmasteradd_Click
  fillGrid
End If



'==============rohan=======================
'Dim deleted As Boolean
'Dim X As Integer
'deleted = False
'If SSTab1.Tab = 0 Then
'
'    If Me.Textfindgl.Text <> "" Then
'        'rs.CancelUpdate
'
'        If rs.State = 1 Then rs.Close
'        rs.Open "select * from gledger where " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
'        rs.Find "gledger='" + Trim(Me.Textfindgl.Text) + "'"
'        If Not rs.EOF Then
'
'           If MsgBox("Are you sure...", vbOKCancel) = vbOK Then
'                  On Error GoTo tt
'                  rs.Delete
'
'tt:             If Err.Number = -2147217887 Then
'                  MsgBox " This Gen. Ledger Have Tranaction.."
'                  Exit Sub
'                End If
'                rs.Update
'                rs.Close
'
'               master.ComboSPECIALCATEGORY.Text = ""
'
'               For Each ctl In Me.Controls
'               If UCase(ctl.Container.Name) = UCase("gledger") Then
'                If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
'                    ctl.Enabled = True
'                End If
'                If TypeOf ctl Is TextBox Then
'                    ctl.Text = ""
'                End If
'                If TypeOf ctl Is ComboBox Then
'
'                    ctl.ListIndex = 0
'                    ctl.Text = ""
'                End If
'                If TypeOf ctl Is CheckBox Then
'                    ctl.Value = 0
'                End If
'            End If
'        Next
'        End If
'                deleted = True
'                Commandmasterabandon_Click
'        End If
'
'    End If
'End If
'If SSTab1.Tab = 1 Then
''On Error GoTo 10
'    If Me.TextFINDSUBLEADGER.Text <> "" Then
''     rs.CancelUpdate
'    If rs.State = 1 Then rs.Close
'   Dim strsledger As String
'
'        rs.Open "select * from sledger where " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
'        rs.Find "subledger='" & Me.TextFINDSUBLEADGER.Text & "'"
'           If Not rs.EOF Then
'                X = MsgBox("Are You Sure...!", vbYesNo, "Warning.....")
'
'                If X = 6 Then
'             On Error GoTo tt1
'
'                    rs.Delete
'
'tt1:             If Err.Number = -2147217887 Then
'                  MsgBox " This Sub. Ledger Have Tranaction.."
'                  Exit Sub
'                End If
'                    rs.Update
'                    deleted = True
'
''                    rsSLD.Open "Select * from  sledger where   gledger='" + Trim(COMBOGENLEDGER.Text) + "' and  Subledger ='" & Comboslgenledgerdiscription.Text & "'", con, adOpenStatic, adLockReadOnly, adCmdText
''
''                    RSopBalDR.Open "SELECT *  FROM VOUCHERS  where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  subledger = '" & Comboslgenledgerdiscription.Text & "'  GROUP BY GenLedger, SubLedger  ", con, adOpenKeyset, adLockReadOnly, adCmdText
''
''                    Ors1.Open "select * from invoicea where genledger='" + Trim(COMBOGENLEDGER.Text) + "' and  subledger = '" & Comboslgenledgerdiscription.Text & "' GROUP BY GenLedger, SubLedger", con, adOpenStatic, adLockReadOnly, adCmdText
''                    Ors2.Open "select * from CREDITa where genledger='" + Trim(COMBOGENLEDGER.Text) + "'  and  subledger = '" & Comboslgenledgerdiscription.Text & "'  GROUP BY GenLedger, SubLedger", con, adOpenStatic, adLockReadOnly, adCmdText
''                    Ors3.Open "select *  from casha where genledger='" + Trim(COMBOGENLEDGER.Text) + "'  and subledger = '" & Comboslgenledgerdiscription.Text & "'   GROUP BY GenLedger, SubLedger ", con, adOpenStatic, adLockReadOnly, adCmdText
''
''                    Ors4.Open "select * from Cnf1a where pgld='" + Trim(COMBOGENLEDGER.Text) + "'  and psld = '" & Comboslgenledgerdiscription.Text & "'  GROUP BY pgld, psld   ", con, adOpenStatic, adLockReadOnly, adCmdText
''
''
''                    Ors5.Open "select *  from dnfa where pgld='" + Trim(COMBOGENLEDGER.Text) + "'  and psld = '" & Comboslgenledgerdiscription.Text & "' GROUP BY pgld, psld ", con, adOpenStatic, adLockReadOnly, adCmdText
''
''                    Ors6.Open "Select * from  CNF1B where gld='" + Trim(COMBOGENLEDGER.Text) + "'  and sld = '" & Comboslgenledgerdiscription.Text & "'  GROUP BY gld, sld ", con, adOpenStatic, adLockReadOnly, adCmdText
''
''                    Ors7.Open "select * from dnfB where gld='" + Trim(COMBOGENLEDGER.Text) + "' and  sld = '" & Comboslgenledgerdiscription.Text & "' GROUP BY gld, sld  ", con, adOpenStatic, adLockReadOnly, adCmdText
''
'
'
'
'
'
'
'
'
'
'
'                    Commandmasterabandon_Click
'                End If
'
'                For Each ctl In Me.Controls
'                        If UCase(ctl.Container.Name) = UCase("gledger") Then
'                              If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
'                                       ctl.Enabled = True
'                              End If
'                              If TypeOf ctl Is TextBox Then 'Or TypeOf ctl Is ComboBox Then
'                                    ctl.Text = ""
'                              End If
'                              If TypeOf ctl Is ComboBox Then
'
'                                         'crl.Text = ""
'                                         ctl.ListIndex = 0
'
'                                End If
'                                If TypeOf ctl Is CheckBox Then
'                                    ctl.Value = 0
'                                End If
'                        End If
'                 Next
'            Else
'
'                    MsgBox "Record not found...."
'            End If
'            If rs.State = 1 Then rs.Close
'         End If
''10:    MsgBox Err.Number
'
'End If
'
'
'If SSTab1.Tab = 2 Then
'    If Me.Textinvep20chartext.Text <> "" Then
'
'        rs.Open " select * from invoiceend  where text='" + Trim(Me.Textinvep20chartext.Text) + "' and " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
'           If Not rs.EOF Then
'                X = MsgBox("Are You Sure...!", vbYesNo, "Warning.....")
'                If X = 6 Then
'                    rs.Delete
'                    rs.Update
'                    deleted = True
'                    Commandmasterabandon_Click
'
'                End If
'
'                For Each ctl In Me.Controls
'                        If UCase(ctl.Container.Name) = UCase("gledger") Then
'                              If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
'                                       ctl.Enabled = True
'                              End If
'                              If TypeOf ctl Is TextBox Then 'Or TypeOf ctl Is ComboBox Then
'                                    ctl.Text = ""
'                              End If
'                              If TypeOf ctl Is ComboBox Then
'
'                                         'crl.Text = ""
'                                         ctl.ListIndex = 0
'
'                                End If
'                                If TypeOf ctl Is CheckBox Then
'                                    ctl.Value = 0
'                                End If
'                        End If
'                 Next
'            Else
'
'                    MsgBox "Record not found...."
'            End If
'            rs.Close
'         End If
'End If
'
'If SSTab1.Tab = 3 Then
'    If Me.Textcnep20chartext <> "" Then
'        rs.Open "select * from CreditEnd where  " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
'        rs.Find "text='" + Trim(Me.Textcnep20chartext) + "'"
'           If Not rs.EOF Then
'                X = MsgBox("Are You Sure...!", vbYesNo, "Warning.....")
'                If X = 6 Then
'                    rs.Delete
'                    rs.Update
'                    deleted = True
'                    Commandmasterabandon_Click
'                End If
'
'                For Each ctl In Me.Controls
'                        If UCase(ctl.Container.Name) = UCase("gledger") Then
'                              If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
'                                       ctl.Enabled = True
'                              End If
'                              If TypeOf ctl Is TextBox Then 'Or TypeOf ctl Is ComboBox Then
'                                    ctl.Text = ""
'                              End If
'                              If TypeOf ctl Is ComboBox Then
'                                         'crl.Text = ""
'                                         ctl.ListIndex = 0
'                              End If
'                              If TypeOf ctl Is CheckBox Then
'                                    ctl.Value = 0
'                               End If
'                        End If
'                   Next
'            Else
'
'                    MsgBox "Record not found...."
'            End If
'            rs.Close
'         End If
'End If
'
''********* FOR  DISCOUNTS
'
'If SSTab1.Tab = 4 Then
'    If Me.Textdcdiscountcategorycode <> "" Then
'        rs.Open "SELECT * FROM disCcats WHERE " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
'        rs.Find " CATEGORYCODE  ='" + Trim(Me.Textdcdiscountcategorycode.Text) + "'"
'           If Not rs.EOF Then
'                X = MsgBox("Are You Sure...!", vbYesNo, "Warning.....")
'
'                If X = 6 Then
'                    rs.Delete
'                    rs.Update
'                    deleted = True
'                    Commandmasterabandon_Click
'                End If
'
'                For Each ctl In Me.Controls
'                        If UCase(ctl.Container.Name) = UCase("gledger") Then
'                              If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
'                                       ctl.Enabled = True
'                              End If
'                              If TypeOf ctl Is TextBox Then 'Or TypeOf ctl Is ComboBox Then
'                                    ctl.Text = ""
'                              End If
'                              If TypeOf ctl Is ComboBox Then
'
'                                         'crl.Text = ""
'                                         ctl.ListIndex = 0
'
'                                End If
'                                If TypeOf ctl Is CheckBox Then
'                                    ctl.Value = 0
'                                End If
'                        End If
'                 Next
'            Else
'
'                    MsgBox "Record not found...."
'            End If
'            rs.Close
'         End If
'End If
'
'
'
'
'
'
'If SSTab1.Tab = 5 Then
'    If Me.cashTextcnep20chartext.Text <> "" Then
'        rs.Open "SELECT * FROM CASHEND " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
'        rs.Find " TEXT  ='" + Trim(Me.cashTextcnep20chartext.Text) + "'"
'         If Not rs.EOF Then
'                X = MsgBox("Are You Sure...!", vbYesNo, "Warning.....")
'
'                If X = 6 Then
'                    rs.Delete
'                    rs.Update
'                    deleted = True
'                    Commandmasterabandon_Click
'
'                End If
'
'                For Each ctl In Me.Controls
'                        If UCase(ctl.Container.Name) = UCase("gledger") Then
'                              If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Or TypeOf ctl Is CheckBox Then
'                                       ctl.Enabled = True
'                              End If
'                              If TypeOf ctl Is TextBox Then 'Or TypeOf ctl Is ComboBox Then
'                                    ctl.Text = ""
'                              End If
'                              If TypeOf ctl Is ComboBox Then
'
'                                        ' crl.Text = ""
'                                         ctl.ListIndex = 0
'
'                                End If
'                                If TypeOf ctl Is CheckBox Then
'                                    ctl.Value = 0
'                                End If
'                        End If
'                 Next
'            Else
'
'                    MsgBox "Record not found...."
'            End If
'            rs.Close
'         End If
'End If
'
'
'
'
'If deleted = True Then
'    Commandmasterdelete.Enabled = False
'    Commandmastersave.Enabled = False
'End If
'
'
''Commandmasterabandon_Click
End Sub
Public Sub Commandmasteredit_Click()
    
    editing = True
    addmaster = False
    Me.Commandmasteradd.Enabled = False
    Me.Commandmasteredit.Enabled = False
    Me.Commandmasterabandon.Enabled = True
    Me.Commandmasteredit.Enabled = True
    
    'If SStab1.Tab = 0 Then
        master.Textglyearopeningbalance.Enabled = True
        master.GMASTERPL.Enabled = True
        master.GMASTERBS.Enabled = True
        master.GMASTERSL.Enabled = True
        master.Cashbankbook.Enabled = True
        master.ComboSPECIALCATEGORY.Enabled = True
        master.Textglgeneralledgerdiscription.Enabled = False
        master.Textglyearopeningbalance.Enabled = True
    'End If
    
''    If SStab1.Tab = 1 Then
''        Me.Comboslgenledgerdiscription.Enabled = True
''        Me.Textslsubledgerdiscription.Enabled = True
''        Textslfindgl.Enabled = True
''        TextFINDSUBLEADGER.Enabled = True
''        Textsldiscriptionforinvoice.Enabled = True
''        Textsladdress1.Enabled = True
''        Textsladdress2.Enabled = True
''        Textsladdress3.Enabled = True
''        Textslyearopeningbalance.Enabled = True
''        Combosldiscountcategory.Enabled = True
''        Combosldistrictcode.Enabled = True
''    End If
''
'''    If SSTab1.Tab = 2 Then
'''        INVEVar = Val(TextInvePrintOrder.Text)
'''    End If
'''    If SSTab1.Tab = 4 Then
'''        INVEVar = Val(TextcnePrintOrder.Text)
'''    End If
'''
'''     If SSTab1.Tab = 5 Then
'''        INVEVar = Val(TextOInvePrintOrder.Text)
'''    End If
        master.Commandmastersave.Enabled = True
        master.Commandmasteredit.Enabled = False
        master.Commandmasteradd.Enabled = False
        master.Commandmasterdelete.Enabled = True
        master.Commandmasterabandon.Enabled = True
End Sub
Private Sub CommandmasterPrint_Click()
    If SStab1.Tab = 0 Then
    '    Genledgerprinting.Show
    ElseIf SStab1.Tab = 1 Then
    MainMenu.cr1.Connect = constr
        MainMenu.cr1.ReportFileName = strrptpath & "\rEPORTS\subledgerlist.RPT"
        MainMenu.cr1.SelectionFormula = "{SLEDGER.fyear}='" & main.session & "' and {SLEDGER.setupid}=" & main.setupid & IIf(Trim(Comboslgenledgerdiscription) <> "", " AND {SLEDGER.GENLEDGER}='" & Comboslgenledgerdiscription & "'", "")
        MainMenu.cr1.WindowShowPrintBtn = True
        MainMenu.cr1.WindowShowPrintSetupBtn = True
        MainMenu.cr1.WindowState = crptMaximized
        MainMenu.cr1.Action = 1
    End If
    
End Sub
Private Sub CommandmasterReturn_Click()

MainMenu.Toolbar1.Visible = True
    Unload Me
End Sub
Private Sub Commandmastersave_Click()
Dim SAVED As Boolean
SAVED = False
'/////////////////*************
'   saving gen ledger
'/////////////////*************
      
'If SStab1.Tab = 0 Then
   
   master.Commandmasteradd.Enabled = True
   master.Commandmasteredit.Enabled = True
   master.gledger.Enabled = False  '************ for frame unlock
   master.Commandmasteradd.Enabled = True
   master.Commandmasteredit.Enabled = True
   master.Commandmastersave.Enabled = False
   
  If ComboSPECIALCATEGORY.Text <> "" And Textglgeneralledgerdiscription <> "" Then
       If rs.State = 1 Then rs.Close
        rs.Open "select * from gledger where " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
        If addmaster = True Then
            rs.Find "gledger='" + Trim(Me.Textglgeneralledgerdiscription.Text) + "'"
            If Not rs.EOF Then
                MsgBox "Record Already exist... "
            Else
                For i = 0 To UBound(arycname)
                rs.AddNew
                rs!gledger = Trim(UCase(Textglgeneralledgerdiscription.Text))
                rs!Category = ComboSPECIALCATEGORY.Text
                If GMASTERPL.Value = 0 Then
                    rs!PLC = False
                Else
                    rs!PLC = True
                End If
                If GMASTERBS.Value = 0 Then
                    rs!BSC = False
                Else
                    rs!BSC = True
                End If
                If GMASTERSL.Value = 0 Then
                    rs!SLF = False
                Else
                    rs!SLF = True
                End If
                rs!YEAROPENING = Val(Textglyearopeningbalance.Text)
                If Cashbankbook.Value = 0 Then
                    rs!Cashbankbook = False
                Else
                    rs!Cashbankbook = True
                End If
                rs!fyear = main.session
                rs!setupid = Val(Left(arycname(i), InStr(1, arycname(i), " (")))
                rs!createdby = main.username
                rs!createdon = Now
                rs.Update
                Next
                SAVED = True
            End If
        Else
            If Not rs.BOF Then
                rs.MoveFirst
            End If
            rs.Find "gledger='" + Trim(Me.Textfindgl.Text) + "'"
            If rs.EOF Then
                MsgBox "Not Found!.."
            Else
                rs!gledger = Trim(UCase(Textglgeneralledgerdiscription.Text))
                rs!Category = ComboSPECIALCATEGORY.Text
                If GMASTERPL.Value = 0 Then
                    rs!PLC = False
                Else
                    rs!PLC = True
                End If
                If GMASTERBS.Value = 0 Then
                    rs!BSC = False
                Else
                    rs!BSC = True
                End If
                If Cashbankbook.Value = 0 Then
                    rs!Cashbankbook = False
                Else
                    rs!Cashbankbook = True
                End If
                If GMASTERSL.Value = 0 Then
                    rs!SLF = False
                Else
                    rs!SLF = True
                End If
                rs!YEAROPENING = Val(Textglyearopeningbalance.Text)
                rs!updatedby = main.username
                rs!updatedon = Now
                rs.Update
                SAVED = True
            End If
        End If
    Else
        MsgBox "ONE OF THE REQUIRED DATA IS NULL. PLEASE CHECK.."
    End If
    If rs.State = 1 Then
        rs.Close
    End If
    Commandmasteredit.Enabled = False
    Commandmasterdelete.Enabled = False
    
'End If
'End If



'''/////////////////*************
'''   saving sub ledger
'''/////////////////*************
''
''If SStab1.Tab = 1 Then
''   master.Commandmasteradd.Enabled = True
''   master.Commandmasteredit.Enabled = True
''    'If Not CBODISTCODE.Text <> "" Or TXTCUSTCODE = "" Then
''    'MsgBox "Select Station Code"
''    'Exit Sub
''    'End If
''    If Textslsubledgerdiscription = "" Then
''    MsgBox "Enter Subledger Name"
''    Exit Sub
''    End If
''    Dim strsledger As String
''    If Combosldistrictcode.Text = "" Then
''    strsledger = (Textslsubledgerdiscription.Text)
''    Else
''    strsledger = txtdistcode & " " & (Textslsubledgerdiscription.Text)
''    End If
''
''    If Comboslgenledgerdiscription.Text <> "" And Textslsubledgerdiscription.Text <> "" Then
''        rs.Open "select * from SLEDGER where " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
''        If addmaster = True Then
''            rs.Find "SUBLEDGER='" + strsledger + "'"
''            If Not rs.EOF Then
''                MsgBox "Record Already exist... "
''            Else
''            ' ADD THE RECORDS IN SUBLEDGER
''                For i = 0 To UBound(arycname)
''                rs.AddNew
''                rs!gledger = Comboslgenledgerdiscription.Text
''                rs!subledger = strsledger
''                rs!DESCFORINVOICE = Textsldiscriptionforinvoice.Text
''                rs!YEAROPENING = Val(Textslyearopeningbalance.Text)
''                'If Trim(Textsladdress1.Text) <> "" Then
''                    rs!address1 = Trim(Textsladdress1.Text)
''                'End If
''                'If Trim(Textsladdress2.Text) <> "" Then
''                    rs!ADDRESS2 = Trim(Textsladdress2.Text)
''                'End If
''                'If Trim(Textsladdress3.Text) <> "" Then
''                    rs!ADDRESS3 = Trim(Textsladdress3.Text)
''                'End If
''                'If Combosldiscountcategory.Text <> "" Then
''                    rs!DISCATEGORY = Combosldiscountcategory.Text
''                'Else
''               '     rs!DISCATEGORY = ""
''                'End If
''                    rs!PHONE = txtphoneno.Text
''                'If Combosldistrictcode.Text <> "" Then
''                    rs!distcode = Combosldistrictcode.Text
''                'Else
''                 '   rs!distcode = ""
''                'End If
''                rs!fyear = main.session
''                rs!setupid = Val(Left(arycname(i), InStr(1, arycname(i), " (")))
''                rs!createdby = main.username
''                rs!createdon = Now
''                rs.Update
''                Next
''                SAVED = True
''                Combosldistrictcode_Click
''            End If
''        Else
''            If Not rs.BOF Then
''                rs.MoveFirst
''            End If
''            rs.Find "SUBLEDGER='" & TextFINDSUBLEADGER.Text & "'"
''            If rs.EOF Then
''                MsgBox "Not Found!.."
''            Else
''            'EDIT THE SUBLEDGER IN ALL THE FILES
''                rs!gledger = Comboslgenledgerdiscription.Text
''                rs!subledger = strsledger
''                rs!DESCFORINVOICE = Textsldiscriptionforinvoice.Text
''                rs!YEAROPENING = Val(Textslyearopeningbalance.Text)
''                'If Trim(Textsladdress1.Text) <> "" Then
''                   rs!address1 = Trim(Textsladdress1.Text)
''                'Else
''                 ' rs!address1 = ""
''                'End If
''                'If Trim(Textsladdress2.Text) <> "" Then
''                    rs!ADDRESS2 = Trim(Textsladdress2.Text)
''                '   Else
''                '     rs!ADDRESS2 = ""
''                'End If
''                'If Trim(Textsladdress3.Text) <> "" Then
''                    rs!ADDRESS3 = Trim(Textsladdress3.Text)
''               '  Else
''                '      rs!ADDRESS3 = ""
''               ' End If
''               ' If Combosldiscountcategory.Text <> "" Then
''                    rs!DISCATEGORY = Combosldiscountcategory.Text
''               ' Else
''                '    rs!DISCATEGORY = ""
''                'End If
''                'If Combosldistrictcode.Text <> "" Then
''                    rs!distcode = Combosldistrictcode.Text
''                'Else
''                '    rs!distcode = ""
''                'End If
''                rs!PHONE = txtphoneno.Text
''                rs!updatedby = main.username
''                rs!updatedon = Now
''                rs.Update
''                rs.Close
''                Dim sq As String
''                rs.Open "select * from vouchers where subledger = '" + TextFINDSUBLEADGER.Text + "' and " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
''                Do While Not rs.EOF
''                rs!subledger = strsledger
''                rs!updatedby = main.username
''                rs!updatedon = Now
''                rs.Update
''                If Not rs.EOF Then
''                rs.MoveNext
''                End If
''                Loop
''                SAVED = True
''                Combosldistrictcode_Click
''                End If
''        End If
''    Else
''        MsgBox "ONE OF THE REQUIRED DATA IS NULL. PLEASE CHECK.."
''        master.sledger.Enabled = True
''        Comboslgenledgerdiscription.Enabled = True
''        Comboslgenledgerdiscription.SetFocus
''
''    End If
''    If rs.State = 1 Then
''        rs.Close
''    End If
''      Commandmasteredit.Enabled = False
''      Commandmasterdelete.Enabled = False
''End If
'''///////////////////******************************
'''   saving INVOICE note end part.
'''//////////////////*******************************
''If SStab1.Tab = 2 Then
''      If rs.State = 1 Then rs.Close
''
''      If Comboinvepdrorcr.Text <> "Debit" And Comboinvepdrorcr.Text <> "Credit" Then
''           MsgBox "Please Enter Debit/Credit.."
''           Comboinvepdrorcr.SetFocus
''        End If
''
''      If Me.Comboinvepcontragenledgerdesc.Text <> "" And Me.Comboinvepgenledgerdesc.Text <> "" And Me.Textinvep20chartext <> "" And Me.TextInvePrintOrder.Text <> "" Then
''        If addmaster = True Then
''                      If rs.State = 1 Then rs.Close
''
''            rs.Open "SELECT * FROM INVOICEEND WHERE TEXT='" + Trim(Me.Textinvep20chartext) + "' And PRINTORDER = " & TextInvePrintOrder.Text & " and " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
''            If Not rs.EOF Then
''                MsgBox "Record Already exist... "
''            Else
''                 For i = 0 To UBound(arycname)
''                rs.AddNew
''
''                rs!printorder = Val(TextInvePrintOrder.Text)
''                rs!CGENLEDGER = Trim(Me.Comboinvepcontragenledgerdesc.Text)
''                If Me.Comboinvepcontrasubledgerdesc.Text <> "" Then
''                    rs!CSUBLEDGER = Trim(Me.Comboinvepcontrasubledgerdesc.Text)
''                Else
''                    rs!CSUBLEDGER = " "
''                End If
''                rs!Genledger = Trim(Me.Comboinvepgenledgerdesc.Text)
''                If Me.Comboinvepsubledgerdesc.Text <> "" Then
''                    rs!subledger = Trim(Me.Comboinvepsubledgerdesc.Text)
''                Else
''                    rs!subledger = " "
''                End If
''                rs!Text = Trim(Me.Textinvep20chartext) & ""
''                rs!rate = Val(Me.Textinveprate)
''                rs!RYN = " "
''                If Trim(Me.Comboinvepdrorcr.Text) = "" Then
''                    Comboinvepdrorcr.SetFocus
''                    Exit Sub
''                Else
''                   rs!DebitorCredit = Trim(Me.Comboinvepdrorcr.Text)
''
''                End If
''                rs!fyear = main.session
''                rs!setupid = Val(Left(arycname(i), InStr(1, arycname(i), " (")))
''                rs!createdby = main.username
''                rs!createdon = Now
''                rs.Update
''                Next
''                SAVED = True
''            End If
''        Else
''            'rs.Open "SELECT * FROM INVOICEEND WHERE CGENLEDGER='" + Trim(Me.Comboinvepcontragenledgerdesc) + "' AND GENLEDGER='" + Trim(Me.Comboinvepgenledgerdesc) + "'" and  , con, adOpenKeyset, adLockOptimistic, adCmdText
''
''            rs.Open "SELECT * FROM INVOICEEND WHERE printorder=" & Val(TextInvePrintOrder.Text) & " and " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
''            'and  printorder <> " & INVEVar & "
''            If rs.RecordCount > 0 Then
''
''                rs!CGENLEDGER = Trim(Me.Comboinvepcontragenledgerdesc.Text)
''                If Me.Comboinvepcontrasubledgerdesc.Text <> "" Then
''                    rs!CSUBLEDGER = Trim(Me.Comboinvepcontrasubledgerdesc.Text)
''                Else
''                    rs!CSUBLEDGER = " "
''                End If
''                rs!Genledger = Trim(Me.Comboinvepgenledgerdesc.Text)
''                If Me.Comboinvepsubledgerdesc.Text <> "" Then
''                    rs!subledger = Trim(Me.Comboinvepsubledgerdesc.Text)
''                Else
''                    rs!subledger = " "
''                End If
''                rs!Text = Trim(Me.Textinvep20chartext) & ""
''                rs!rate = Val(Me.Textinveprate)
''                rs!RYN = " "
''                rs!DebitorCredit = Trim(Me.Comboinvepdrorcr.Text)
''                rs!printorder = Val(TextInvePrintOrder.Text)
''                rs!updatedby = main.username
''                rs!updatedon = Now
''                rs.Update
''                SAVED = True
''            End If
''        End If
''    Else
''        MsgBox "ONE OF THE REQUIRED DATA IS NULL. PLEASE CHECK.."
''        Comboinvepcontragenledgerdesc.SetFocus
''        Exit Sub
''    End If
''    If rs.State = 1 Then
''        rs.Close
''    End If
''End If
''
'''/////////////////*************
'''   saving credit note end part.
'''////////////////*************
''
''If SStab1.Tab = 3 Then
''
''    If Combocnepdrorcr.Text <> "Debit" And Combocnepdrorcr.Text <> "Credit" Then
''         MsgBox "Please Enter Debit/Credit.."
''         Combocnepdrorcr.SetFocus
''          Exit Sub
''    End If
''
''
''    If Me.Combocnepcontragenledgerdesc.Text <> "" And Me.Combocnepgenledgerdesc.Text <> "" And Me.Textcnep20chartext <> "" And Me.Textcneprate <> "" And Me.Combocnepdrorcr.Text <> "" Then
''        If addmaster = True Then
''            rs.Open "SELECT * FROM CREDITEND WHERE  printorder= " & CneTextInvePrintOrder.Text & " and " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
''            If Not rs.EOF Then
''                MsgBox "Record Already exist... "
''            Else
''                 For i = 0 To UBound(arycname)
''                 rs.AddNew
''                rs!printorder = Val(CneTextInvePrintOrder.Text)
''                rs!CGENLEDGER = Trim(Me.Combocnepcontragenledgerdesc.Text)
''                If Me.Combocnepcontrasubledgerdesc.Text <> "" Then
''                    rs!CSUBLEDGER = Trim(Me.Combocnepcontrasubledgerdesc.Text)
''                Else
''                    rs!CSUBLEDGER = " "
''                End If
''                rs!Genledger = Trim(Me.Combocnepgenledgerdesc.Text)
''                If Me.Combocnepsubledgerdesc.Text <> "" Then
''                    rs!subledger = Trim(Me.Combocnepsubledgerdesc.Text)
''                Else
''                    rs!subledger = " "
''                End If
''                rs!Text = Trim(Me.Textcnep20chartext) & ""
''                rs!rate = Val(Me.Textcneprate)
''                rs!DebitorCredit = Trim(Me.Combocnepdrorcr.Text)
''                rs!RYN = " "
''                rs!createdby = main.username
''                rs!createdon = Now
''                rs!fyear = main.session
''                rs!setupid = Val(Left(arycname(i), InStr(1, arycname(i), " (")))
''                rs.Update
''                Next
''                SAVED = True
''            End If
''        Else
''            rs.Open "SELECT * FROM CREDITEND WHERE printorder= " & CneTextInvePrintOrder.Text & " and " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
''            If rs.EOF Then
''                MsgBox "Not Found!.."
''            Else
''                rs!printorder = Val(CneTextInvePrintOrder.Text)
''                rs!CGENLEDGER = Me.Combocnepcontragenledgerdesc.Text
''                If Me.Combocnepcontrasubledgerdesc.Text <> "" Then
''                    rs!CSUBLEDGER = Me.Combocnepcontrasubledgerdesc.Text
''                Else
''                    rs!CSUBLEDGER = " "
''                End If
''                rs!Genledger = Me.Combocnepgenledgerdesc.Text
''                If Me.Combocnepsubledgerdesc.Text <> "" Then
''                    rs!subledger = Me.Combocnepsubledgerdesc.Text
''                Else
''                    rs!subledger = " "
''                End If
''                rs!Text = Me.Textcnep20chartext & ""
''                rs!rate = Val(Me.Textcneprate)
''                rs!RYN = " "
''                rs!DebitorCredit = Trim(Me.Combocnepdrorcr.Text)
''                rs!updatedby = main.username
''                rs!updatedon = Now
''                rs.Update
''                SAVED = True
''            End If
''        End If
''    Else
''        MsgBox "ONE OF THE REQUIRED DATA IS NULL. PLEASE CHECK.."
''    End If
''    If rs.State = 1 Then
''        rs.Close
''    End If
''   master.Commandmasteradd.Enabled = True
''   master.Commandmasteredit.Enabled = True
''
''End If
''
'''/////////////////*************
'''   saving discount cat.
'''////////////////*************
''
''If SStab1.Tab = 4 Then
''    If Textdcdiscountcategorycode.Text <> "" And Combobgroupcode.Text <> "" And Combobgroupname.Text <> "" And Textdcdiscountrate.Text <> "" Then
''        If addmaster = True Then
''            rs.Open "select * from DISCCATS where categorycode='" + Trim(Textdcdiscountcategorycode.Text) + "' and groupcode='" + Trim(Combobgroupcode.Text) + "' and " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
''            If Not rs.EOF Then
''                MsgBox "Record Already exist... "
''            Else
''                For i = 0 To UBound(arycname)
''                rs.AddNew
''                rs!categorycode = Textdcdiscountcategorycode.Text
''                rs!GROUPCODE = Combobgroupcode.Text
''                rs!discountrate = Val(Textdcdiscountrate.Text)
''                rs!fyear = main.session
''                rs!setupid = Val(Left(arycname(i), InStr(1, arycname(i), " (")))
''                rs!createdby = main.username
''                rs!createdon = Now
''                rs.Update
''                Next
''                SAVED = True
''            End If
''        Else
''            rs.Open "select * from DISCCATS where categorycode='" + Trim(Textfinddiscountcategory.Text) + "' and groupcode='" + Trim(textfinddiscgroupcode.Text) + "' and " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
''            If rs.EOF Then
''                MsgBox "Not Found!.."
''            Else
''                rs!categorycode = Textdcdiscountcategorycode.Text
''                rs!GROUPCODE = Combobgroupcode.Text
''                rs!discountrate = Val(Textdcdiscountrate.Text)
''                rs!updatedby = main.username
''                rs!updatedon = Now
''                rs.Update
''                SAVED = True
''            End If
''        End If
''    Else
''        MsgBox "ONE OF THE REQUIRED DATA IS NULL. PLEASE CHECK.."
''        On Error Resume Next
''        Textdcdiscountcategorycode.SetFocus
''
''    End If
''    If rs.State = 1 Then
''        rs.Close
''    End If
''   master.Commandmasteradd.Enabled = True
''   master.Commandmasteredit.Enabled = True
''
''End If
''
'''************************ CASH /BANK
'''If SSTab1.Tab = 5 Then
'' '   If Me.Textcbgenledger.Text <> "" Then
''  '      If addmaster = True Then
''   '         rs.Open "SELECT * FROM CBMF WHERE GLD='" + Trim(Me.Textcbgenledger) + "' AND SLD='" + Trim(Me.Textcbsubledger.Text) + "'", con, adOpenKeyset, adLockOptimistic, adCmdText
''    '        If Not rs.EOF Then
''     '           MsgBox "Record Already exist... "
''      '      Else
''       '         rs.AddNew
''        '        rs(0) = Trim(Me.Textcbgenledger.Text)
''         '       rs(1) = Trim(Me.Textcbsubledger.Text)
''          '      rs.Update
''           '     SAVED = True
''    '        End If
''     '   Else
''      '      rs.Open "SELECT * FROM CBMF WHERE GLD='" + Trim(Me.Textcbgenledger.Text) + "' AND SLD='" + Trim(Me.Textcbsubledger.Text) + "'", con, adOpenKeyset, adLockOptimistic, adCmdText
''       '     If rs.EOF Then
''        '        MsgBox "Not Found!.."
''         '   Else
''          '      rs(0) = Trim(Me.Textcbgenledger.Text)
''           '     rs(1) = Trim(Me.Textcbsubledger.Text)
''            '    rs.Update
''             '   SAVED = True'
'' '           End If
''  '      End If
''   ' Else
''    '    MsgBox "ONE OF THE REQUIRED DATA IS NULL. PLEASE CHECK.."
'' '   End If
'' '   If rs.State = 1 Then
''  '      rs.Close
''   ' End If
'''End If
''
''If SStab1.Tab = 5 Then
''
''    If cashCombocnepdrorcr.Text <> "Debit" And cashCombocnepdrorcr.Text <> "Credit" Then
''         MsgBox "Please Enter Debit/Credit.."
''          cashCombocnepdrorcr.SetFocus
''          Exit Sub
''    End If
''
''    If Me.cashCombocnepcontragenledgerdesc.Text <> "" And Me.cashCombocnepgenledgerdesc.Text <> "" And Me.cashTextcnep20chartext <> "" And Me.cashTextcneprate <> "" And Me.cashCombocnepdrorcr.Text <> "" Then
''        If addmaster = True Then
''            rs.Open "SELECT * FROM CASHEND WHERE PRINTORDER = " & cashTextInvePrintOrder.Text & " AND " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
''
''            If Not rs.EOF Then
''                MsgBox "Record Already exist... "
''            Else
''                For i = 0 To UBound(arycname)
''                rs.AddNew
''                rs!printorder = Val(cashTextInvePrintOrder.Text)
''                rs!CGENLEDGER = Trim(Me.cashCombocnepcontragenledgerdesc.Text)
''                If Me.cashCombocnepcontrasubledgerdesc.Text <> "" Then
''                    rs!CSUBLEDGER = Trim(Me.cashCombocnepcontrasubledgerdesc.Text)
''                Else
''                    rs!CSUBLEDGER = " "
''                End If
''                rs!Genledger = Trim(Me.cashCombocnepgenledgerdesc.Text)
''                If Me.cashCombocnepsubledgerdesc.Text <> "" Then
''                    rs!subledger = Trim(Me.cashCombocnepsubledgerdesc.Text)
''                Else
''                    rs!subledger = " "
''                End If
''                rs!Text = Trim(Me.cashTextcnep20chartext) & ""
''                rs!rate = Val(Me.cashTextcneprate)
''                rs!DebitorCredit = Trim(Me.cashCombocnepdrorcr.Text)
''                rs!RYN = " "
''                rs!createdby = main.username
''                rs!createdon = Now
''                rs!fyear = main.session
''                 rs!setupid = Val(Left(arycname(i), InStr(1, arycname(i), " (")))
''                rs.Update
''                Next
''                SAVED = True
''            End If
''        Else
''            rs.Open "SELECT * FROM CASHEND WHERE  printorder= " & master.cashTextInvePrintOrder.Text & " and " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
''            If rs.EOF Then
''                MsgBox "Not Found!.."
''            Else
''                rs!printorder = Val(cashTextInvePrintOrder.Text)
''                rs!CGENLEDGER = Me.cashCombocnepcontragenledgerdesc.Text
''                If Me.cashCombocnepcontrasubledgerdesc.Text <> "" Then
''                    rs!CSUBLEDGER = Me.cashCombocnepcontrasubledgerdesc.Text
''                Else
''                    rs!CSUBLEDGER = " "
''                End If
''                rs!Genledger = Me.cashCombocnepgenledgerdesc.Text
''                If Me.cashCombocnepsubledgerdesc.Text <> "" Then
''                    rs!subledger = Me.cashCombocnepsubledgerdesc.Text
''                Else
''                    rs!subledger = " "
''                End If
''                rs!Text = Me.cashTextcnep20chartext & ""
''                rs!rate = Val(Me.cashTextcneprate)
''                rs!DebitorCredit = Trim(Me.cashCombocnepdrorcr.Text)
''                rs!RYN = " "
''                rs!updatedby = main.username
''                rs!updatedon = Now
''                rs.Update
''                SAVED = True
''            End If
''        End If
''    Else
''        MsgBox "ONE OF THE REQUIRED DATA IS NULL. PLEASE CHECK.."
''    End If
''    If rs.State = 1 Then
''        rs.Close
''    End If
''End If
''
''If SAVED Then
''        editing = False
''        For Each ctl In Me.Controls
''            If TypeOf ctl Is textbox Then
''                    If ctl.Enabled = True Then
''                    ctl.Text = ""
''                    ctl.Enabled = False
''                    End If
''            End If
''            If TypeOf ctl Is ComboBox Then
''                    If ctl.Style <> 2 Then
''                    ctl.Text = ""
''                    ctl.Enabled = False
''                    End If
''            End If
''            If TypeOf ctl Is CheckBox Then
''                    ctl.Value = 0
''                    ctl.Enabled = False
''          End If
''        Next
''        TXTCUSTCODE = ""
'End If

Commandmastersave.Enabled = False
Commandmasteredit.Enabled = False
fillGrid

End Sub
Private Sub Commandmastersearch_Click()
  Me.Enabled = False
 'searchscreen.Grid1.row = 0
 'searchscreen.Grid1.col = 0
  Call searchscreen.tempr(SStab1.Tab, Me.Name)
  SetButton Commandmasteradd, Commandmasteredit, Commandmastersave, Commandmasterdelete
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub
Sub fillGrid()
    Dim f As New ADODB.Recordset
    If f.State = 1 Then f.Close
    S = "SELECT [Category],[gledger],[PLC],[BSC],[SLF],[YEAROPENING],[cashbankbook] FROM [GLEDGER] ORDER BY [Category],[gledger]"
    f.Open S, CON
    Set vs.DataSource = f
      
    
End Sub
Private Sub Form_Load()
' /****      FRAMEINI      ****/
   fillGrid
    Me.TOP = 20
    Me.Left = 200
    Dim TMPA As Control
    editing = False
    INVEVar = 0
''    For Each TMPA In Me.Controls
''        If TypeOf TMPA Is VB.frame Then
''            TMPA.TOP = 1200
''            TMPA.Left = 800
''            TMPA.Width = 7515
''            TMPA.Height = 4005
''        End If
''        If TypeOf TMPA Is textbox Then
''            TMPA.Enabled = False
''        End If
''        If TypeOf TMPA Is CheckBox Then
''            TMPA.Enabled = False
''        End If
''        If TypeOf TMPA Is ComboBox Then
''            TMPA.Enabled = False
''        End If
''    Next
    ' ComboSPECIALCATEGORY INI
    ComboSPECIALCATEGORY.AddItem "Assets"
    ComboSPECIALCATEGORY.AddItem "Liability"
    ComboSPECIALCATEGORY.AddItem "Income"
    ComboSPECIALCATEGORY.AddItem "Expences"
'    Set CON = New ADODB.Connection
''    With CON
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\data.mdb"
'        .Open
'    End With
    Set rs = New ADODB.Recordset
    rs.Open "select * from gledger where slf=1 and " & stridnyear, CON, adOpenDynamic, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Comboslgenledgerdiscription.AddItem rs!gledger
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    
    
'   If rs.State = 1 Then rs.Close
'   rs.Open "select DISTCODE from DISTRICTS where  " & stridnyear & "   ORDER BY DISTCODE", CON, adOpenKeyset, adLockOptimistic, adCmdText
'   If Not rs.EOF Then
'        Do While Not rs.EOF
'          If Not IsNull(rs.Fields(0).Value) Then
'           CBODISTCODE.AddItem rs.Fields(0).Value
'           rs.MoveNext
'          End If
'       Loop
'    End If
'    rs.Close
'
    
    Combosldistrictcode.AddItem ""
        
   If rs.State = 1 Then rs.Close
   rs.Open "select DISTRICTNAME from DISTRICTS where  " & stridnyear & "   ORDER BY DISTRICTNAME", CON, adOpenKeyset, adLockOptimistic, adCmdText
   If Not rs.EOF Then
        Do While Not rs.EOF
          If Not IsNull(rs.Fields(0).Value) Then
           Combosldistrictcode.AddItem rs.Fields(0).Value
           rs.MoveNext
          End If
       Loop
    End If
    
    rs.Close
    
    
    
    
    
    
    '/ ***** Combobgroupcode
    rs.Open "select * from GROUPS where " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Me.Combobgroupcode.AddItem rs!GROUPCODE
            Me.Combobgroupname.AddItem rs!GroupName
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    rs.Open "select distinct categorycode from DISCCATS where  " & stridnyear & "   order by categorycode", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Me.Combosldiscountcategory.AddItem rs!categorycode
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
'    rs.Open "select * from DISTRICTS where " & stridnyear & " order by DISTRICTNAME", CON, adOpenKeyset, adLockReadOnly, adCmdText
'    If Not rs.EOF Then
 '       Do While Not rs.EOF
  '          Me.Combosldistrictcode.AddItem rs!DISTRICTNAME
   '         If Not rs.EOF Then
    '            rs.MoveNext
    '        End If
    '    Loop
  '  End If
  '  rs.Close
    rs.Open "select gledger from gledger where  " & stridnyear & "   order by gledger", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Me.Combocnepcontragenledgerdesc.AddItem rs!gledger
            Me.cashCombocnepcontragenledgerdesc.AddItem rs!gledger
            Me.Comboinvepcontragenledgerdesc.AddItem rs!gledger
            Me.Comboinvepgenledgerdesc.AddItem rs!gledger
            Me.Combocnepgenledgerdesc.AddItem rs!gledger
            Me.cashCombocnepgenledgerdesc.AddItem rs!gledger
  
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    Commandmastersearch.Enabled = True
    CommandmasterPrint.Enabled = True
    SetButton Commandmasteradd, Commandmasteredit, Commandmastersave, Commandmasterdelete
    Commandmasteredit.Enabled = False
    Commandmastersave.Enabled = False
    Commandmasterdelete.Enabled = False
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'refreshme
End Sub

Private Sub Textcbgenledger_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Dim rs As New ADODB.Recordset
   If rs.State = 1 Then rs.Close
   rs.Open "select * from SLEDGER where GLEDGER ='" & Textcbgenledger.Text & "' and " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
   If Not rs.EOF Then
        Do While Not rs.EOF
            Textcbsubledger.AddItem rs(1)
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
End If
End Sub

Private Sub Textcbgenledger_LostFocus()
   
       If Textcbgenledger.Text <> "" Then
            Dim rs4 As New ADODB.Recordset
            rs4.Open "Select* from gledger where GLEDGER = '" + Trim(Textcbgenledger.Text) + "' and " & stridnyear, CON, adOpenStatic
            If rs4.RecordCount <= 0 Then
                 MsgBox "No valid Gen.Ledger"
                 Textcbgenledger.SetFocus
            End If
        End If
   

End Sub

Private Sub Textcbsubledger_LostFocus()
If Textcbsubledger.Text <> "" And Textcbgenledger.Text <> "" Then
   Dim rs4 As New ADODB.Recordset
   rs4.Open "Select* from sledger where GLEDGER='" + Trim(Textcbgenledger.Text) + "' and SubLedger='" + Trim(Textcbsubledger.Text) + "' and " & stridnyear, CON, adOpenStatic
   If rs4.RecordCount <= 0 Then
      MsgBox "No valid Sub Ledger"
      Textcbsubledger.SetFocus
   End If
End If
If Textcbsubledger.ListCount > 0 And Textcbsubledger.Text = "" Then
'    Textcbsubledger.SetFocus
  
 End If

End Sub

Private Sub Textcnep20chartext_LostFocus()
 Textcnep20chartext.Text = UCase(Textcnep20chartext.Text)
End Sub

Private Sub Textcneprate_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub
Private Sub Textdcdiscountcategorycode_LostFocus()
    Textdcdiscountcategorycode.Text = UCase(Textdcdiscountcategorycode.Text)
End Sub
Private Sub Textdcdiscountrate_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
 Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub Textdcdiscountrate_LostFocus()
   Textdcdiscountrate.Text = Format(Textdcdiscountrate.Text, "0.00")
End Sub

Private Sub Textglgeneralledgerdiscription_LostFocus()
    Textglgeneralledgerdiscription.Text = UCase(Textglgeneralledgerdiscription.Text)
        Set rs = New ADODB.Recordset
        rs.Open "select * from gledger where " & stridnyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
        rs.Find "gledger='" + Trim(Textglgeneralledgerdiscription.Text) + "'"
        If Not rs.EOF Then
            Textglgeneralledgerdiscription.Text = rs!gledger
            Textfindgl.Text = rs!gledger
            ComboSPECIALCATEGORY.Text = rs!Category
            If rs!PLC = False Then
                GMASTERPL.Value = 0
            Else
                GMASTERPL.Value = 1
            End If
            If rs!BSC = False Then
                GMASTERBS.Value = 0
            Else
                GMASTERBS.Value = 1
            End If
            If rs!Cashbankbook = False Then
                Cashbankbook.Value = 0
            Else
                Cashbankbook.Value = 1
            End If
            If rs!SLF = False Then
                GMASTERSL.Value = 0
            Else
                GMASTERSL.Value = 1
            End If
            Textglyearopeningbalance.Text = rs!YEAROPENING
            ComboSPECIALCATEGORY.Text = rs!Category
        End If
        rs.Close
End Sub

Private Sub Textglyearopeningbalance_GotFocus()
Textglyearopeningbalance.Text = Format(Textglyearopeningbalance.Text, "0.00")
End Sub

Private Sub Textglyearopeningbalance_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub



Private Sub Textinvep20chartext_LostFocus()
Textinvep20chartext.Text = UCase(Textinvep20chartext.Text)
End Sub

Private Sub Textinveprate_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub TextInvePrintOrder_LostFocus()


If IsNumeric(TextInvePrintOrder.Text) = False Then
    MsgBox "Please Enter Any No..."
    TextInvePrintOrder.SetFocus
End If


End Sub

Private Sub Textsladdress1_LostFocus()
    Textsladdress1.Text = UCase(Textsladdress1.Text)
End Sub
Private Sub Textsladdress2_LostFocus()
    Textsladdress2.Text = UCase(Textsladdress2.Text)
End Sub
Private Sub Textsladdress3_LostFocus()
    Textsladdress3.Text = UCase(Textsladdress3.Text)
End Sub
Private Sub Textsldiscriptionforinvoice_LostFocus()
    Textsldiscriptionforinvoice.Text = UCase(Textsldiscriptionforinvoice.Text)
End Sub

Private Sub Textslsubledgerdiscription_GotFocus()
If Trim(Comboslgenledgerdiscription.Text) = "" Then
    MsgBox "Please Select Gen. Ledger ."
    Comboslgenledgerdiscription.SetFocus
End If
End Sub

Private Sub Textslsubledgerdiscription_LostFocus()
If Trim(Comboslgenledgerdiscription.Text) <> "" Then
    Textslsubledgerdiscription.Text = UCase(Textslsubledgerdiscription.Text)
    Set rs = New ADODB.Recordset
        rs.Open "select * from SLEDGER where gledger='" + Trim(Comboslgenledgerdiscription.Text) + "' and subledger='" + IIf(Combosldistrictcode.Text <> "", Combosldistrictcode.Text & "-" & TXTCUSTCODE, "") & Trim(Textslsubledgerdiscription.Text) + "' and " & stridnyear, CON, adOpenKeyset, adLockReadOnly, adCmdText
        If Not rs.EOF Then
            Textslfindgl.Text = rs!gledger
            TextFINDSUBLEADGER.Text = rs!SUBLEDGER
            
            'Textsldiscriptionforinvoice.Text = rs!DESCFORINVOICE
            
            If IsNull(rs!DESCFORINVOICE) Then
               Textsldiscriptionforinvoice.Text = ""
            Else
                Textsldiscriptionforinvoice.Text = rs!DESCFORINVOICE

            End If
            
            If IsNull(rs!ADDRESS1) Then
               Textsladdress1.Text = ""
            Else
               Textsladdress1.Text = rs!ADDRESS1
            End If
            
            If IsNull(rs!ADDRESS2) Then
                    Textsladdress2.Text = ""
               Else
                  Textsladdress2.Text = rs!ADDRESS2

            End If
            
            If IsNull(rs!ADDRESS3) Then
               Textsladdress3.Text = ""
            Else
               Textsladdress3.Text = rs!ADDRESS3
            End If
            'Textsladdress2.Text = rs!ADDRESS2
            'Textsladdress3.Text = rs!ADDRESS3
            Textslyearopeningbalance.Text = Format(rs!YEAROPENING, "0.00")
            If IsNull(rs!DISCATEGORY) Then
                 Combosldiscountcategory.Text = ""
            Else
                Combosldiscountcategory.Text = rs!DISCATEGORY
                
            End If
            
            If IsNull(rs!distcode) Then
                Combosldistrictcode.Text = ""
               
            ElseIf rs!distcode <> "" Then
                Combosldistrictcode.Text = rs!distcode
                Else
                Combosldistrictcode.ListIndex = 0
            End If

            
            If editing Then
                'editing = False
                Me.Comboslgenledgerdiscription.Enabled = False
            End If
        End If
        rs.Close
End If
If Trim(Textslsubledgerdiscription) = "" Then
    Me.Comboslgenledgerdiscription.Enabled = True
End If
End Sub
Private Sub Textslyearopeningbalance_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
    Else
        If KeyAscii <> 46 Then
            If KeyAscii <> 8 And KeyAscii <> 45 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Function refreshme()
' /****      FRAMEINI      ****/
    Me.TOP = 20
    Me.Left = 200
    Dim TMPA As Control
    editing = False
    For Each TMPA In Me.Controls
        If TypeOf TMPA Is VB.frame Then
            TMPA.TOP = 1200
            TMPA.Left = 800
            TMPA.Width = 7515
            TMPA.Height = 4005
        End If
        If TypeOf TMPA Is textbox Then
            TMPA.Enabled = False
        End If
        If TypeOf TMPA Is CheckBox Then
            TMPA.Enabled = False
        End If
        If TypeOf TMPA Is ComboBox Then
            TMPA.Enabled = False
        End If
    Next
    ' ComboSPECIALCATEGORY INI
    ComboSPECIALCATEGORY.AddItem "Assets"
    ComboSPECIALCATEGORY.AddItem "Liability"
    ComboSPECIALCATEGORY.AddItem "Income"
    ComboSPECIALCATEGORY.AddItem "Expences"
'    Set CON = New ADODB.Connection
''    With CON
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + App.Path + "\" + Trim(main.directory) + "\data.mdb"
'        .Open
'    End With
    Set rs = New ADODB.Recordset
    rs.Open "select * from gledger where slf=1 and " & stridnyear, CON, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Comboslgenledgerdiscription.AddItem rs!gledger
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    '/ ***** Combobgroupcode
    rs.Open "select * from GROUPS where " & stridnyear & " ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Me.Combobgroupcode.AddItem rs!GROUPCODE
            Me.Combobgroupname.AddItem rs!GroupName
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    rs.Open "select distinct categorycode from DISCCATS where  " & stridnyear & "  order by categorycode", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Me.Combosldiscountcategory.AddItem rs!categorycode
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    rs.Open "select * from DISTRICTS where  " & stridnyear & "   order by DISTRICTNAME ", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Me.Combosldistrictcode.AddItem rs!DISTRICTNAME
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    rs.Open "select gledger from gledger where slf=1 and  " & stridnyear & "   order by gledger", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Me.Combocnepcontragenledgerdesc.AddItem rs!gledger
            Me.Comboinvepcontragenledgerdesc.AddItem rs!gledger
            Me.Comboinvepgenledgerdesc.AddItem rs!gledger
            Me.Combocnepgenledgerdesc.AddItem rs!gledger
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    rs.Open "select subledger from sledger where  " & stridnyear & "   order by subledger", CON, adOpenKeyset, adLockReadOnly, adCmdText
    If Not rs.EOF Then
        Do While Not rs.EOF
            Me.Combocnepcontrasubledgerdesc.AddItem rs!SUBLEDGER
            Me.Comboinvepcontrasubledgerdesc.AddItem rs!SUBLEDGER
            Me.Comboinvepsubledgerdesc.AddItem rs!SUBLEDGER
            Me.Combocnepsubledgerdesc.AddItem rs!SUBLEDGER
            If Not rs.EOF Then
                rs.MoveNext
            End If
        Loop
    End If
    rs.Close
    'SSTab1.Tab = 0
    Commandmastersearch.Enabled = True
    CommandmasterPrint.Enabled = True
End Function

Private Sub Textslyearopeningbalance_LostFocus()
Textslyearopeningbalance.Text = Format(Textslyearopeningbalance.Text, "0.00")
End Sub
Private Sub vs_Click()
                
                
If rs.State = 1 Then rs.Close
rs.Open "select * from GLEDGER where (Category='" & vs.TextMatrix(vs.RowSel, 0) & "' and gledger='" & vs.TextMatrix(vs.RowSel, 1) & "')"
If rs.EOF = False Then
    
    Commandmasteredit.Enabled = True
    
    Textglgeneralledgerdiscription.Text = rs!gledger
    ComboSPECIALCATEGORY.Text = rs!Category
    If rs!PLC = False Then
       GMASTERPL.Value = 0
    Else
       GMASTERPL.Value = 1
    End If
    
    If rs!BSC = False Then
       GMASTERBS.Value = 0
    Else
        GMASTERBS.Value = 1
    End If
    
    
    If rs!SLF = False Then
       GMASTERSL.Value = 0
    Else
       GMASTERSL.Value = 1
    End If
    
    Textglyearopeningbalance.Text = rs!YEAROPENING
    
    If rs!Cashbankbook = False Then
       Cashbankbook.Value = 0
    Else
       Cashbankbook.Value = 1
    End If

End If

End Sub
