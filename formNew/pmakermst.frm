VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form papermaker 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10308
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9120
   ScaleWidth      =   10308
   Begin VB.Frame panel 
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   9075
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   10125
      Begin VB.CommandButton cmdPaper 
         BackColor       =   &H00FFFFFF&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   960
         Width           =   585
      End
      Begin VB.ComboBox na 
         Height          =   315
         ItemData        =   "pmakermst.frx":0000
         Left            =   1860
         List            =   "pmakermst.frx":000D
         TabIndex        =   1
         Top             =   1020
         Width           =   4350
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         Height          =   780
         Left            =   390
         ScaleHeight     =   732
         ScaleWidth      =   5040
         TabIndex        =   53
         Top             =   4515
         Width           =   5085
         Begin VB.CommandButton close 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Close"
            Height          =   675
            Left            =   3720
            Picture         =   "pmakermst.frx":001F
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   30
            Width           =   1200
         End
         Begin VB.CommandButton REPORTCD 
            Caption         =   "&Print"
            Height          =   450
            Left            =   7560
            TabIndex        =   58
            Top             =   60
            Width           =   1200
         End
         Begin VB.CommandButton Abandon 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Clear"
            Height          =   675
            Left            =   45
            Picture         =   "pmakermst.frx":0C03
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   30
            Width           =   1200
         End
         Begin VB.CommandButton Del 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Delete"
            Height          =   675
            Left            =   2475
            Picture         =   "pmakermst.frx":17E7
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   30
            Width           =   1200
         End
         Begin VB.CommandButton save 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            Height          =   675
            Left            =   1275
            Picture         =   "pmakermst.frx":23CB
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   30
            Width           =   1200
         End
         Begin VB.CommandButton Help 
            Caption         =   "&Help"
            Height          =   450
            Left            =   -360
            TabIndex        =   55
            Top             =   0
            Width           =   300
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&Search"
            Height          =   450
            Left            =   6300
            TabIndex        =   54
            Top             =   240
            Visible         =   0   'False
            Width           =   60
         End
      End
      Begin VB.TextBox txtpmid 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1845
         MaxLength       =   10
         TabIndex        =   0
         Top             =   630
         Width           =   2025
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   2235
         Width           =   585
      End
      Begin VB.CommandButton Command1 
         Height          =   435
         Left            =   5295
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   4455
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.CommandButton Command2 
         Height          =   435
         Left            =   5325
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   4455
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.ComboBox cboSize 
         Height          =   315
         Left            =   5025
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   4515
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.ComboBox cboGSM 
         Height          =   288
         Left            =   1845
         Style           =   1  'Simple Combo
         TabIndex        =   5
         Top             =   2715
         Width           =   3495
      End
      Begin VB.ComboBox cboEco 
         Height          =   315
         ItemData        =   "pmakermst.frx":2FAF
         Left            =   1845
         List            =   "pmakermst.frx":2FB1
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2295
         Width           =   4335
      End
      Begin VB.ComboBox cboBright 
         Height          =   315
         Left            =   1845
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3585
         Width           =   4335
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3540
         Width           =   585
      End
      Begin VB.ComboBox cboPType 
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1455
         Width           =   4335
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1440
         Width           =   585
      End
      Begin VB.TextBox txtSize1 
         Height          =   315
         Left            =   1845
         TabIndex        =   6
         Top             =   3135
         Width           =   675
      End
      Begin VB.TextBox txtSize2 
         Height          =   315
         Left            =   3345
         TabIndex        =   8
         Top             =   3135
         Width           =   675
      End
      Begin VB.TextBox txtSize3 
         Height          =   315
         Left            =   4965
         TabIndex        =   10
         Top             =   3135
         Width           =   735
      End
      Begin VB.ComboBox cboCM1 
         Height          =   315
         ItemData        =   "pmakermst.frx":2FB3
         Left            =   2505
         List            =   "pmakermst.frx":2FC0
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3135
         Width           =   870
      End
      Begin VB.ComboBox cboCM2 
         Height          =   315
         ItemData        =   "pmakermst.frx":2FD2
         Left            =   4005
         List            =   "pmakermst.frx":2FDF
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3135
         Width           =   990
      End
      Begin VB.ComboBox cboCM3 
         Height          =   315
         ItemData        =   "pmakermst.frx":2FF1
         Left            =   5685
         List            =   "pmakermst.frx":2FFE
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3135
         Width           =   1110
      End
      Begin VB.ComboBox cboReel_Sheet 
         Height          =   315
         ItemData        =   "pmakermst.frx":3010
         Left            =   1860
         List            =   "pmakermst.frx":301A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1875
         Width           =   4335
      End
      Begin VSFlex7Ctl.VSFlexGrid VS 
         Height          =   3540
         Left            =   360
         TabIndex        =   69
         Top             =   5460
         Width           =   8205
         _cx             =   14473
         _cy             =   6244
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
         BackColorFixed  =   7917545
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
         SelectionMode   =   1
         GridLines       =   1
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
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For Search Click the Row ..."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5700
         TabIndex        =   70
         Top             =   4980
         Width           =   1995
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   870
         Left            =   345
         Top             =   4470
         Width           =   5190
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Code :"
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
         Height          =   225
         Left            =   450
         TabIndex        =   68
         Top             =   690
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Mill :"
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
         Height          =   225
         Left            =   450
         TabIndex        =   67
         Top             =   1065
         Width           =   1305
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Press F2 for Search Paper Maker"
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
         Height          =   195
         Left            =   1875
         TabIndex        =   66
         Top             =   360
         Width           =   2835
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Quality :"
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
         Height          =   225
         Left            =   450
         TabIndex        =   65
         Top             =   2340
         Width           =   1365
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "GSM :"
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
         Height          =   225
         Left            =   450
         TabIndex        =   64
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Size :"
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
         Height          =   225
         Left            =   450
         TabIndex        =   63
         Top             =   3135
         Width           =   1245
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Brightness :"
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
         Height          =   225
         Left            =   450
         TabIndex        =   62
         Top             =   3630
         Width           =   1245
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Paper Type :"
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
         Height          =   225
         Left            =   465
         TabIndex        =   61
         Top             =   1500
         Width           =   1365
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "Sheet/Reel :"
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
         Height          =   225
         Left            =   465
         TabIndex        =   60
         Top             =   1920
         Width           =   1365
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C000&
      Height          =   285
      Left            =   17580
      TabIndex        =   22
      Top             =   4020
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.TextBox city 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   17145
      MaxLength       =   100
      TabIndex        =   16
      Top             =   2100
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox pno1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   17145
      MaxLength       =   100
      TabIndex        =   17
      Top             =   2430
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox pno2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   17145
      MaxLength       =   100
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox Faxno 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   17145
      MaxLength       =   100
      TabIndex        =   19
      Top             =   3090
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox emailid 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   17145
      MaxLength       =   100
      TabIndex        =   21
      Top             =   3750
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.ListBox List1 
      Height          =   1776
      Left            =   17940
      Sorted          =   -1  'True
      TabIndex        =   33
      Top             =   1440
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.TextBox add2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   17145
      MaxLength       =   100
      TabIndex        =   15
      Top             =   1785
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox add1 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   17145
      MaxLength       =   100
      TabIndex        =   14
      Top             =   1470
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.TextBox mobile 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   17145
      MaxLength       =   100
      TabIndex        =   20
      Top             =   3420
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFD7AE&
      Caption         =   "Contact Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3060
      Left            =   16080
      TabIndex        =   43
      Top             =   1380
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox txtMobile5 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3465
         MaxLength       =   50
         TabIndex        =   32
         Top             =   1950
         Width           =   2415
      End
      Begin VB.TextBox txtMobile4 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3465
         MaxLength       =   50
         TabIndex        =   30
         Top             =   1620
         Width           =   2415
      End
      Begin VB.TextBox txtMobile3 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3465
         MaxLength       =   50
         TabIndex        =   28
         Top             =   1290
         Width           =   2415
      End
      Begin VB.TextBox txtContact2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   225
         MaxLength       =   50
         TabIndex        =   25
         Top             =   960
         Width           =   3180
      End
      Begin VB.TextBox txtMobile2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3465
         MaxLength       =   50
         TabIndex        =   26
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtContact3 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   225
         MaxLength       =   50
         TabIndex        =   27
         Top             =   1290
         Width           =   3180
      End
      Begin VB.TextBox txtContact4 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   225
         MaxLength       =   50
         TabIndex        =   29
         Top             =   1605
         Width           =   3180
      End
      Begin VB.TextBox txtContact5 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   225
         MaxLength       =   50
         TabIndex        =   31
         Top             =   1950
         Width           =   3180
      End
      Begin VB.TextBox txtMobile1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3465
         MaxLength       =   50
         TabIndex        =   24
         Top             =   630
         Width           =   2415
      End
      Begin VB.TextBox txtContact1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   225
         MaxLength       =   50
         TabIndex        =   23
         Top             =   645
         Width           =   3180
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   3465
         TabIndex        =   45
         Top             =   390
         Width           =   735
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   2
         Left            =   225
         TabIndex        =   44
         Top             =   375
         Width           =   660
      End
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11400
      TabIndex        =   42
      Top             =   4560
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "City   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   16125
      TabIndex        =   41
      Top             =   2115
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Resi Phone Nos. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   16125
      TabIndex        =   40
      Top             =   2835
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax Nos. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   16125
      TabIndex        =   39
      Top             =   3165
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Office Phone Nos. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   16125
      TabIndex        =   38
      Top             =   2475
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Address1 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   16140
      TabIndex        =   37
      Top             =   1500
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   16125
      TabIndex        =   36
      Top             =   3810
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Address2 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   16125
      TabIndex        =   35
      Top             =   1785
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   16050
      TabIndex        =   34
      Top             =   3480
      Visible         =   0   'False
      Width           =   1080
   End
End
Attribute VB_Name = "papermaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ref As ADODB.Recordset
Dim flag As Boolean
Dim RS As New ADODB.Recordset
Dim value As String
Sub lstfocus()

    On Error Resume Next
    
    txtpmid = ref!papermaker_id
    na = ref!papermaker_name
    cboReel_Sheet.text = ref!Size & ""
    
    If Not IsNull(ref!eco) Then
    cboEco.text = ref!eco
    End If
    
    If Not IsNull(ref!GSM) Then
      cboGSM.text = ref!GSM
    End If
    
    If Not IsNull(ref!Size) Then
    cboSize.text = ref!Size
    End If
    
    If Not IsNull(ref!bright) Then
    cboBright.text = ref!bright
    End If
    
    cboPType.text = ref!ptype
    
    txtSize1.text = ref!SizeValue1
    txtSize2.text = ref!SizeValue2
    txtSize3.text = ref!SizeValue3
    
    cboCM1.text = ref!size1 & ""
    cboCM2.text = ref!Size2 & ""
    cboCM3.text = ref!Size3 & ""
    
    
    cboCM1.Enabled = True
    cboCM2.Enabled = True
    cboCM3.Enabled = True
    
    txtSize1.Enabled = True
    
    txtSize2.Enabled = True
    
    txtSize3.Enabled = True


End Sub
Sub COMPINI()
    na.ListIndex = -1
    cboEco.ListIndex = -1
    cboGSM.ListIndex = -1
    cboSize.ListIndex = -1
    cboReel_Sheet.ListIndex = -1
    txtSize1.text = ""
    txtSize2.text = ""
    txtSize3.text = ""
    cboCM1.ListIndex = -1
    cboCM2.ListIndex = -1
    cboCM3.ListIndex = -1
    cboEco.ListIndex = -1
    cboPType.ListIndex = -1
    cboSize.ListIndex = -1
    cboBright.ListIndex = -1
    cboGSM.text = ""
    na.SetFocus
    MaxNo
End Sub
Private Sub ABANDON_Click()
    COMPINI
End Sub
Private Sub cboBright_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     save_Click
  End If
End Sub
Private Sub cboCM1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then sendkeys "{tab}"
End Sub

Private Sub cboCM2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then sendkeys "{tab}"
End Sub
Private Sub cboCM3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then sendkeys "{tab}"
End Sub
Private Sub cboEco_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cboGSM.SetFocus
End If
End Sub
Private Sub cboGSM_GotFocus()

If PopUpValue1 <> "" Then
   cboGSM.text = PopUpValue1
   PopUpValue1 = ""
   PopUpValue2 = ""
End If

End Sub
Private Sub cboGSM_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
   value = "Select GSM as [Paper Size],gsm_info as Remarks from GSMMaster order by GSM"
   popuplistModel10 value, con
End If

End Sub
Private Sub cboGSM_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   txtSize1.SetFocus
End If

End Sub
Private Sub cboPType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then sendkeys "{tab}"
End Sub

Private Sub cboReel_Sheet_Click()
If cboReel_Sheet.text = "Reel" Then
    txtSize2.Enabled = False
    cboCM2.Enabled = False
    txtSize3.Enabled = False
    cboCM3.Enabled = False
Else
    txtSize2.Enabled = True
    cboCM2.Enabled = True
    txtSize3.Enabled = True
    cboCM3.Enabled = True

End If

End Sub

Private Sub cboReel_Sheet_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then sendkeys "{tab}"
End Sub

Private Sub cboSize_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
    Call save_Click
 End If

End Sub
Private Sub Check1_Click()
  
If Check1.value = 1 Then
   Me.Caption = "Manufacturer"
Else
   Me.Caption = "Vendor"
End If

End Sub

Private Sub close_Click()
Unload Me
End Sub

Private Sub cmdPaper_Click()
HeadTbl = "papermill"
frmMasters.Show 1

End Sub

Private Sub cmdSearch_Click()
   value = "Select papermaker_name as [Vendor Name],add1 as Address from papermakemaster where " & stringyear & " and papermaker_id <> ''"
   popuplist1 value, con
End Sub

Private Sub Command1_Click()
HeadTbl = "GSM"
frmMasters.Show 1
End Sub

Private Sub Command2_Click()
HeadTbl = "Size"
frmMasters.Show 1
End Sub

Private Sub Command3_Click()
HeadTbl = "bright"
frmMasters.Show 1
End Sub

Private Sub Command4_Click()
HeadTbl = "Book_q"
frmMasters.Show 1
End Sub

Private Sub Command5_Click()
HeadTbl = "ptype"
frmMasters.Show 1

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   SendKeys "{tab}"
'End If
End Sub
Sub AddItem()
 
 
 If RS.State = 1 Then RS.close
 RS.Open "Select * from MasterTbl where category='" & HeadTbl & "' order by Name", con, adOpenDynamic, adLockOptimistic
 If HeadTbl = "Book_q" Then
    
    cboEco.Clear
    While RS.EOF = False
       cboEco.AddItem RS!Name
       RS.MoveNext
    Wend
 
 ElseIf HeadTbl = "GSM" Then
 
    cboGSM.Clear
    While RS.EOF = False
       cboGSM.AddItem RS!Name
       RS.MoveNext
    Wend
    
 ElseIf HeadTbl = "Size" Then
 
    cboSize.Clear
    While RS.EOF = False
       cboSize.AddItem RS!Name
       RS.MoveNext
    Wend
 
 ElseIf HeadTbl = "bright" Then
 
    cboBright.Clear
    While RS.EOF = False
       cboBright.AddItem RS!Name
       RS.MoveNext
    Wend
 
 ElseIf HeadTbl = "ptype" Then
 
    cboPType.Clear
    While RS.EOF = False
       cboPType.AddItem RS!Name
       RS.MoveNext
    Wend
 
 
 End If
 

End Sub
Sub MaxNo()
If RS.State = 1 Then RS.close
RS.Open "select max(convert(int,papermaker_id)) from PaperMakeMaster where " & stringyear & "", con
If IsNull(RS.Fields(0).value) Then
   txtpmid.text = 1
Else
   txtpmid.text = RS.Fields(0).value + 1
End If
End Sub
Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub List1_Click()
For I = 0 To List1.ListCount - 1
            If List1.Selected(I) = True Then
                If ref.State = 1 Then ref.close
                ref.Open "Select * from papermakemaster where " & stringyear & " and papermaker_name = '" + List1.List(I) + "'", con, adOpenDynamic, adLockOptimistic, adCmdText
                If ref.RecordCount > 0 Then
                  lstfocus
                End If
               Exit Sub
               End If
          Next I
End Sub
Private Sub NA_Click()

fillGrid

End Sub
Private Sub Del_Click()

Dim X As Integer
Dim rs1 As New ADODB.Recordset
If na.text = "" Then
MsgBox "Please Select Binder Name... ", vbInformation
na.SetFocus
Exit Sub
End If
X = MsgBox("Are you sure you wish to delete the selected item ", 4, "Confirmation")
If X = 6 Then
   con.Execute "Delete  from papermakemaster where papermaker_id= '" & txtpmid.text & "' and " & stringyear
na = ""

txtpmid = ""
List1.Clear
na.Clear

rs1.Open "select distinct papermaker_name from papermakemaster where " & stringyear & "", con, adOpenStatic, adLockReadOnly
If rs1.RecordCount > 0 Then
    
    rs1.MoveFirst
    Do While Not rs1.EOF
        List1.AddItem rs1!papermaker_name
        na.AddItem rs1!papermaker_name
        rs1.MoveNext
    Loop
    
End If
COMPINI
End If




na.SetFocus
End Sub
Private Sub Form_Load()


Me.Left = 100
Me.top = 100
Me.Width = 9800
Me.Height = 9800



frmNo = "paper"


Set ref = New ADODB.Recordset
'ref.Open "select distinct papermaker_name from papermakemaster where " & stringyear & " order by papermaker_name", con, adOpenDynamic, adLockOptimistic
ref.Open "Select * from MasterTbl where category='papermill' order by Name", con, adOpenDynamic, adLockOptimistic




If Not ref.BOF Then
    ref.MoveFirst
End If
na.Clear
Do While Not ref.EOF
    List1.AddItem ref!Name
    na.AddItem ref!Name
    If Not ref.EOF Then
       ref.MoveNext
    End If
Loop



MaxNo

'------------------------------------------------
If RS.State = 1 Then RS.close
RS.Open "Select * from MasterTbl where category='Book_q'", con, adOpenDynamic, adLockOptimistic
cboEco.Clear
While RS.EOF = False
   cboEco.AddItem RS!Name
   RS.MoveNext
Wend
 
 
If RS.State = 1 Then RS.close
RS.Open "Select * from MasterTbl where category='bright'", con, adOpenDynamic, adLockOptimistic
cboBright.Clear
While RS.EOF = False
   cboBright.AddItem RS!Name
   RS.MoveNext
Wend
 
 
If RS.State = 1 Then RS.close
RS.Open "Select * from MasterTbl where category='ptype'", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   cboPType.AddItem RS!Name
   RS.MoveNext
Wend
 

 
 
If RS.State = 1 Then RS.close
RS.Open "Select gsm from GsmMaster where " & stringyear & " order by gsm", con, adOpenDynamic, adLockOptimistic
cboGSM.Clear
While RS.EOF = False
   cboGSM.AddItem RS(0)
   RS.MoveNext
Wend
 
 
If RS.State = 1 Then RS.close
RS.Open "Select size1 from SizeMaster where " & stringyear & " order by size1", con, adOpenDynamic, adLockOptimistic
cboSize.Clear
While RS.EOF = False
   cboSize.AddItem RS(0)
   RS.MoveNext
Wend
 
Me.Caption = "Paper Master..."
 
BackColorFrom Me


End Sub

Sub fillGrid()

I = 1

If na.text = "" Then Exit Sub

If rs1.State = 1 Then rs1.close
rs1.Open "Select papermaker_name as [Paper Name],SizeValue1 + ' '+ Size1 + ' , '+ SizeValue2 + ' '+ Size2 + ' , '+ SizeValue3 + ' '+ Size3 as [Paper Size]," & _
"PType,Size as [Sheet/Real],Eco + '   -  ' + GSM  as [Quality & GSM],papermaker_Id as Code from " & _
" papermakemaster where papermaker_name = '" & na.text & "' and " & stringyear

vs.rows = 1

While rs1.EOF = False
vs.rows = vs.rows + 1
vs.TextMatrix(I, 0) = rs1(0)
vs.TextMatrix(I, 1) = rs1(1)
vs.TextMatrix(I, 2) = rs1(2)
vs.TextMatrix(I, 3) = rs1(3)
vs.TextMatrix(I, 4) = rs1(4)
vs.TextMatrix(I, 5) = rs1(5)
I = I + 1
rs1.MoveNext
Wend


vs.FormatString = "Paper Mill|Paper Size|PType|Sheet/Reel|Quality & GSM|Code"

vs.ColWidth(0) = 1200
vs.ColWidth(1) = 2200
vs.ColWidth(2) = 1200
vs.ColWidth(3) = 900
vs.ColWidth(4) = 2000

End Sub

Private Sub NA_KeyPress(KeyAscii As Integer)
''If KeyAscii = 8 Then
''If Len(Trim(na.Text)) <> 0 Then
''        na.Text = Left(na.Text, (Len(na.Text) - 1))
''End If
''End If
''
''If Len(Trim(na.Text)) = 50 Then
''KeyAscii = 0
''End If

If KeyAscii = 13 Then
   cboPType.SetFocus
End If

End Sub
Sub paperUpdate_()
 
If RS.State = 1 Then RS.close
RS.Open "select * from PaperMakeMaster where papermaker_id=" & txtpmid.text & "", con, adOpenStatic, adLockReadOnly

If RS.EOF = False Then
    pname = ""
    
    
    pname = RS!papermaker_name
    
    If Len(RS!eco) > 1 Then
    If RS!eco <> "" Then
       pname = pname & "-" & RS!eco
    End If
    End If
    
    If Val(RS!SizeValue1) > 0 Then
       pname = pname & "-" & RS!SizeValue1
    End If
    
    If Val(RS!SizeValue2) > 0 Then
       pname = pname & "X" & RS!SizeValue2
    End If
    
    
    If RS!GSM <> "" Then
       pname = pname & "-" & RS!GSM
    End If
    
    con.Execute "update PaperMakeMaster set papername1='" & pname & "' where papermaker_id='" & txtpmid.text & "'"
End If
   

 
End Sub

Private Sub save_Click()

On Error GoTo save1

Dim rs_ss As New ADODB.Recordset


If na.text = "" Then
   MsgBox "Enter Paper Name ..", vbInformation
   na.SetFocus
   Exit Sub
End If
'-------------------------------------------
If cboPType.text = "" Then
   MsgBox "Select Paper Type ..", vbInformation
   cboPType.SetFocus
   Exit Sub
End If
'-------------------------------------------
'-------------------------------------------
If cboGSM.text = "" Then
   MsgBox "Select Paper GSM ..", vbInformation
   cboGSM.SetFocus
   Exit Sub
End If
'-------------------------------------------





If na <> "" Then
If txtpmid <> "" Then
If txtpmid <> "" Then
Dim FN As String


If rs_ss.State = 1 Then rs_ss.close
rs_ss.Open "select * from PaperMakeMaster where papermaker_id='" & txtpmid.text & "' and " & stringyear, con, adOpenDynamic, adLockOptimistic


If rs_ss.EOF = True Then
    rs_ss.AddNew
    'List1.AddItem na
    rs_ss!papermaker_id = UCase(Trim(txtpmid))
End If
    
rs_ss!papermaker_name = UCase(Trim(na.text))
rs_ss!eco = (Trim(cboEco.text))
rs_ss!GSM = (Trim(cboGSM.text))
rs_ss!Size = (Trim(cboReel_Sheet.text))
rs_ss!bright = (Trim(cboBright.text))
rs_ss!ptype = Trim(cboPType.text)

rs_ss!SizeValue1 = txtSize1.text
rs_ss!SizeValue2 = txtSize2.text
rs_ss!SizeValue3 = txtSize3.text

rs_ss!size1 = cboCM1.text
rs_ss!Size2 = cboCM2.text
rs_ss!Size3 = cboCM3.text

rs_ss!fyear = session
rs_ss!setupid = setupid

rs_ss.update


paperUpdate_

MsgBox " Record Saved .... ", vbInformation
COMPINI
na.SetFocus

End If
Else
   MsgBox "Fill the Paper Maker id"
   txtpmid.SetFocus
End If
Else
   MsgBox "Fill the Paper Maker Name"
na.SetFocus

End If

Exit Sub

save1:

MsgBox "" & err.DESCRIPTION

'Form_Load
End Sub

Private Sub Text2_Change()

End Sub

Private Sub txtpmid_GotFocus()

If PopUpValue1 <> "" Then

txtpmid = popupvalue5
If ref.State = 1 Then ref.close
ref.Open "Select * from papermakemaster where papermaker_id = '" + popupvalue5 + "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If ref.RecordCount > 0 Then
  lstfocus
Else
flag = True
COMPINI
End If

PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
popupvalue4 = ""
popupvalue5 = ""

End If

End Sub

Private Sub txtpmid_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    popuplistModel10 "SELECT papermaker_name as PaperName,SizeValue1 + ' X ' + SizeValue2 as Size,ECO,GSM,papermaker_id as Id from PaperMakeMaster order by papermaker_name", con
End If



End Sub
Private Sub txtpmid_KeyPress(KeyAscii As Integer)
Label12.Visible = False
If KeyAscii = 13 Then sendkeys "{tab}"
End Sub
Private Sub txtSize1_GotFocus()

If PopUpValue1 <> "" Then
   If InStr(PopUpValue1, "X") > 0 Then
      txtSize1.text = Trim(Mid(PopUpValue1, 1, InStr(PopUpValue1, "X") - 1))
      txtSize2.text = Trim(Mid(PopUpValue1, InStr(PopUpValue1, "X") + 1, 2))
   End If
   PopUpValue1 = ""
   PopUpValue2 = ""
End If

End Sub
Private Sub txtSize1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   value = "Select size1 as [Paper Size],size_info as Remarks from SizeMaster where " & stringyear & " order by size1"
   popuplistModel10 value, con
End If

End Sub

Private Sub txtSize1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then sendkeys "{tab}"
End Sub

Private Sub txtSize2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then sendkeys "{tab}"
End Sub

Private Sub txtSize3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then sendkeys "{tab}"
End Sub
Private Sub vs_Click()

'Label12.Visible = True

Set ref = New ADODB.Recordset
ref.Open "Select * from papermakemaster where " & stringyear & " and papermaker_id= '" + vs.TextMatrix(vs.RowSel, 5) + "'", con, adOpenDynamic, adLockOptimistic, adCmdText
If ref.RecordCount > 0 Then
  lstfocus
ElseIf na.text <> "" Then
Else
flag = True
COMPINI
End If

PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue6 = ""

na.Enabled = True
End Sub
