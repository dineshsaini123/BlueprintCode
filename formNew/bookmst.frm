VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D76D7130-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7d.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmbook 
   BackColor       =   &H00FFD7AE&
   Caption         =   "Book Master"
   ClientHeight    =   10512
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   14088
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10512
   ScaleWidth      =   14088
   Begin VB.Frame panel 
      Height          =   10536
      Left            =   60
      TabIndex        =   56
      Top             =   0
      Width           =   13944
      Begin VB.TextBox txtRem 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   4608
         MaxLength       =   100
         TabIndex        =   248
         Top             =   8928
         Width           =   8076
      End
      Begin VB.Frame frmBookList 
         Caption         =   "Book List"
         Height          =   2388
         Left            =   6216
         TabIndex        =   190
         Top             =   1956
         Visible         =   0   'False
         Width           =   7056
         Begin VB.CommandButton cmdok 
            Caption         =   "&OK"
            Height          =   384
            Left            =   5724
            TabIndex        =   193
            Top             =   120
            Width           =   675
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close"
            Height          =   384
            Left            =   6444
            TabIndex        =   192
            Top             =   120
            Width           =   555
         End
         Begin VSFlex7DAOCtl.VSFlexGrid VS_bk 
            Height          =   1956
            Left            =   24
            TabIndex        =   191
            Top             =   540
            Width           =   7020
            _cx             =   12382
            _cy             =   3450
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
            BackColorBkg    =   -2147483636
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
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
            Editable        =   2
            ShowComboButton =   -1  'True
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   2250
         Left            =   10320
         TabIndex        =   57
         Top             =   1620
         Visible         =   0   'False
         Width           =   1950
         Begin VSFlex7Ctl.VSFlexGrid vs 
            Height          =   6540
            Left            =   0
            TabIndex        =   62
            Top             =   240
            Width           =   6000
            _cx             =   10583
            _cy             =   11536
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
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
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
         Begin VB.CommandButton cmdExitfrm 
            BackColor       =   &H0080C0FF&
            Caption         =   "Close"
            Height          =   465
            Left            =   6165
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox txtGpno1 
            Height          =   285
            Left            =   6120
            TabIndex        =   60
            Top             =   780
            Width           =   1290
         End
         Begin VB.CommandButton Command17 
            BackColor       =   &H0080C0FF&
            Caption         =   "Change BG No."
            Height          =   465
            Left            =   6165
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   1440
            Width           =   1455
         End
         Begin VB.ComboBox cboClass1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   6120
            TabIndex        =   58
            Text            =   "cboClass"
            Top             =   420
            Width           =   1320
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Class Name :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Index           =   1
            Left            =   6165
            TabIndex        =   64
            Top             =   135
            Width           =   1395
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Group  :"
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
            Index           =   2
            Left            =   4545
            TabIndex        =   63
            Top             =   1035
            Width           =   825
         End
      End
      Begin VB.CommandButton cmd12 
         Caption         =   "Clear"
         Height          =   315
         Left            =   12240
         TabIndex        =   246
         Top             =   8568
         Width           =   495
      End
      Begin VB.ComboBox txtHead12 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   480
         TabIndex        =   228
         Top             =   8496
         Width           =   1335
      End
      Begin VB.TextBox txtHeadData12 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1860
         MaxLength       =   80
         TabIndex        =   229
         Top             =   8496
         Width           =   465
      End
      Begin VB.ComboBox cboPrinter12 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   3576
         TabIndex        =   232
         Top             =   8532
         Width           =   2235
      End
      Begin VB.ComboBox cboColour12 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         ItemData        =   "bookmst.frx":0000
         Left            =   5856
         List            =   "bookmst.frx":000D
         TabIndex        =   233
         Top             =   8532
         Width           =   1695
      End
      Begin VB.TextBox txtPCode12 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8316
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   235
         Top             =   8568
         Width           =   570
      End
      Begin VB.TextBox txtPaper12 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8916
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   236
         Top             =   8568
         Width           =   3300
      End
      Begin VB.TextBox txtTextSupp12 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3036
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   231
         Top             =   8496
         Width           =   480
      End
      Begin VB.ComboBox cbosupp12 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   2340
         TabIndex        =   230
         Top             =   8496
         Width           =   675
      End
      Begin VB.CommandButton Command24_12 
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
         Height          =   360
         Left            =   7596
         Style           =   1  'Graphical
         TabIndex        =   234
         Top             =   8532
         Width           =   435
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   8100
         MaxLength       =   80
         TabIndex        =   245
         Top             =   8496
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.CommandButton cmd11 
         Caption         =   "Clear"
         Height          =   315
         Left            =   12240
         TabIndex        =   243
         Top             =   8208
         Width           =   495
      End
      Begin VB.ComboBox txtHead11 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   480
         TabIndex        =   219
         Top             =   8136
         Width           =   1335
      End
      Begin VB.TextBox txtHeadData11 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1860
         MaxLength       =   80
         TabIndex        =   220
         Top             =   8136
         Width           =   465
      End
      Begin VB.ComboBox cboPrinter11 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   3576
         TabIndex        =   223
         Top             =   8208
         Width           =   2235
      End
      Begin VB.ComboBox cboColour11 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         ItemData        =   "bookmst.frx":0039
         Left            =   5856
         List            =   "bookmst.frx":0046
         TabIndex        =   224
         Top             =   8208
         Width           =   1695
      End
      Begin VB.TextBox txtPCode11 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8316
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   226
         Top             =   8208
         Width           =   570
      End
      Begin VB.TextBox txtPaper11 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8916
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   227
         Top             =   8208
         Width           =   3300
      End
      Begin VB.TextBox txtTextSupp11 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3036
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   222
         Top             =   8172
         Width           =   480
      End
      Begin VB.ComboBox cbosupp11 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   2340
         TabIndex        =   221
         Top             =   8172
         Width           =   675
      End
      Begin VB.CommandButton Command24_11 
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
         Height          =   348
         Left            =   7596
         Style           =   1  'Graphical
         TabIndex        =   225
         Top             =   8172
         Width           =   435
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   8100
         MaxLength       =   80
         TabIndex        =   242
         Top             =   8172
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   8100
         MaxLength       =   80
         TabIndex        =   240
         Top             =   7812
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.CommandButton Command24_10 
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
         Height          =   348
         Left            =   7596
         Style           =   1  'Graphical
         TabIndex        =   216
         Top             =   7812
         Width           =   435
      End
      Begin VB.ComboBox cbosupp10 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   2340
         TabIndex        =   212
         Top             =   7812
         Width           =   675
      End
      Begin VB.TextBox txtTextSupp10 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3036
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   213
         Top             =   7812
         Width           =   480
      End
      Begin VB.TextBox txtPaper10 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8916
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   218
         Top             =   7848
         Width           =   3300
      End
      Begin VB.TextBox txtPCode10 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8316
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   217
         Top             =   7848
         Width           =   570
      End
      Begin VB.ComboBox cboColour10 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         ItemData        =   "bookmst.frx":0072
         Left            =   5856
         List            =   "bookmst.frx":007F
         TabIndex        =   215
         Top             =   7848
         Width           =   1695
      End
      Begin VB.ComboBox cboPrinter10 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   3576
         TabIndex        =   214
         Top             =   7848
         Width           =   2235
      End
      Begin VB.TextBox txtHeadData10 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1860
         MaxLength       =   80
         TabIndex        =   211
         Top             =   7812
         Width           =   465
      End
      Begin VB.ComboBox txtHead10 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   480
         TabIndex        =   210
         Top             =   7812
         Width           =   1335
      End
      Begin VB.CommandButton cmd10 
         Caption         =   "Clear"
         Height          =   315
         Left            =   12240
         TabIndex        =   239
         Top             =   7848
         Width           =   495
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   8100
         MaxLength       =   80
         TabIndex        =   237
         Top             =   7452
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.CommandButton Command14_9 
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
         Height          =   348
         Left            =   7596
         Style           =   1  'Graphical
         TabIndex        =   207
         Top             =   7452
         Width           =   435
      End
      Begin VB.ComboBox cbosupp9 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   2340
         TabIndex        =   203
         Top             =   7488
         Width           =   675
      End
      Begin VB.TextBox txtTextSupp9 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3036
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   204
         Top             =   7488
         Width           =   480
      End
      Begin VB.TextBox txtPaper9 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8916
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   209
         Top             =   7524
         Width           =   3300
      End
      Begin VB.TextBox txtPCode9 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8316
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   208
         Top             =   7524
         Width           =   570
      End
      Begin VB.ComboBox cboColour9 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         ItemData        =   "bookmst.frx":00AB
         Left            =   5856
         List            =   "bookmst.frx":00B8
         TabIndex        =   206
         Top             =   7488
         Width           =   1695
      End
      Begin VB.ComboBox cboPrinter9 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   3576
         TabIndex        =   205
         Top             =   7524
         Width           =   2235
      End
      Begin VB.TextBox txtHeadData9 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1860
         MaxLength       =   80
         TabIndex        =   202
         Top             =   7488
         Width           =   465
      End
      Begin VB.ComboBox txtHead9 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   480
         TabIndex        =   201
         Top             =   7488
         Width           =   1335
      End
      Begin VB.CommandButton cmd9 
         Caption         =   "Clear"
         Height          =   315
         Left            =   12240
         TabIndex        =   200
         Top             =   7488
         Width           =   495
      End
      Begin VB.ComboBox cbofirm 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "bookmst.frx":00E4
         Left            =   6555
         List            =   "bookmst.frx":00E6
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   495
         Width           =   2925
      End
      Begin VB.CommandButton cmd8 
         Caption         =   "Clear"
         Height          =   315
         Left            =   12240
         TabIndex        =   99
         Top             =   6000
         Width           =   495
      End
      Begin VB.CommandButton cmd7 
         Caption         =   "Clear"
         Height          =   315
         Left            =   12240
         TabIndex        =   129
         Top             =   7080
         Width           =   495
      End
      Begin VB.CommandButton cmd6 
         Caption         =   "Clear"
         Height          =   315
         Left            =   12240
         TabIndex        =   119
         Top             =   6720
         Width           =   495
      End
      Begin VB.CommandButton cmd5 
         Caption         =   "Clear"
         Height          =   315
         Left            =   12240
         TabIndex        =   109
         Top             =   6345
         Width           =   495
      End
      Begin VB.CommandButton cmd4 
         Caption         =   "Clear"
         Height          =   315
         Left            =   12240
         TabIndex        =   198
         Top             =   5640
         Width           =   495
      End
      Begin VB.CommandButton cmd3 
         Caption         =   "Clear"
         Height          =   315
         Left            =   12240
         TabIndex        =   197
         Top             =   5340
         Width           =   495
      End
      Begin VB.CommandButton cnd2 
         Caption         =   "Clear"
         Height          =   315
         Left            =   12240
         TabIndex        =   196
         Top             =   5040
         Width           =   495
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Clear"
         Height          =   315
         Left            =   12240
         TabIndex        =   195
         Top             =   4680
         Width           =   495
      End
      Begin VB.CommandButton cmdUpDatePrice 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update Price Book Wise"
         Height          =   465
         Left            =   8175
         Style           =   1  'Graphical
         TabIndex        =   194
         Top             =   810
         Width           =   1245
      End
      Begin VB.ComboBox cboBinding 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         ItemData        =   "bookmst.frx":00E8
         Left            =   6996
         List            =   "bookmst.frx":00EA
         TabIndex        =   17
         Top             =   4056
         Width           =   1524
      End
      Begin VB.CommandButton Command20 
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
         Height          =   360
         Left            =   8592
         Style           =   1  'Graphical
         TabIndex        =   188
         Top             =   3996
         Width           =   375
      End
      Begin VB.CommandButton Command19 
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
         Height          =   360
         Left            =   5808
         Style           =   1  'Graphical
         TabIndex        =   187
         Top             =   4008
         Width           =   375
      End
      Begin VB.ComboBox txtTrimSize 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         ItemData        =   "bookmst.frx":00EC
         Left            =   4128
         List            =   "bookmst.frx":00EE
         TabIndex        =   16
         Top             =   4068
         Width           =   1656
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add &Type"
         Height          =   348
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   185
         Top             =   8844
         Width           =   915
      End
      Begin VB.ComboBox txtHead5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   480
         TabIndex        =   85
         Top             =   6030
         Width           =   1335
      End
      Begin VB.ComboBox txtHead6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   480
         TabIndex        =   100
         Top             =   6345
         Width           =   1335
      End
      Begin VB.ComboBox txtHead8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   480
         TabIndex        =   120
         Top             =   7080
         Width           =   1335
      End
      Begin VB.ComboBox txtHead7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   480
         TabIndex        =   110
         Top             =   6720
         Width           =   1335
      End
      Begin VB.ComboBox txtHead4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   480
         TabIndex        =   49
         Top             =   5640
         Width           =   1335
      End
      Begin VB.ComboBox txtHead3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   480
         MousePointer    =   8  'Size NW SE
         TabIndex        =   42
         Top             =   5340
         Width           =   1335
      End
      Begin VB.ComboBox txtHead2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   480
         MousePointer    =   2  'Cross
         TabIndex        =   35
         Top             =   5040
         Width           =   1335
      End
      Begin VB.ComboBox txtHead1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   480
         TabIndex        =   28
         Top             =   4680
         Width           =   1335
      End
      Begin VB.TextBox txtname 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Left            =   1530
         MaxLength       =   100
         TabIndex        =   3
         Top             =   840
         Width           =   6630
      End
      Begin VB.TextBox txtRetLY 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8430
         MaxLength       =   80
         TabIndex        =   27
         Top             =   3600
         Width           =   720
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update Data"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9012
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   3960
         Width           =   1035
      End
      Begin VB.TextBox txtRet 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   7170
         MaxLength       =   80
         TabIndex        =   26
         Top             =   3615
         Width           =   780
      End
      Begin VB.CommandButton cmdBookNoAd 
         BackColor       =   &H0080C0FF&
         Caption         =   "Gp. Book No. Adjustment"
         Height          =   585
         Left            =   8205
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   1500
         Width           =   1245
      End
      Begin VB.TextBox txtHeadData6 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1860
         MaxLength       =   80
         TabIndex        =   101
         Top             =   6345
         Width           =   465
      End
      Begin VB.TextBox txtHeadData7 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1860
         MaxLength       =   80
         TabIndex        =   111
         Top             =   6720
         Width           =   465
      End
      Begin VB.TextBox txtHeadData8 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1860
         MaxLength       =   80
         TabIndex        =   121
         Top             =   7080
         Width           =   465
      End
      Begin VB.ComboBox cboPrinter6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3570
         TabIndex        =   104
         Top             =   6345
         Width           =   2235
      End
      Begin VB.ComboBox cboPrinter7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3570
         TabIndex        =   114
         Top             =   6720
         Width           =   2235
      End
      Begin VB.ComboBox cboPrinter8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   3570
         TabIndex        =   124
         Top             =   7110
         Width           =   2235
      End
      Begin VB.ComboBox cboColour6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "bookmst.frx":00F0
         Left            =   5850
         List            =   "bookmst.frx":00FD
         TabIndex        =   105
         Top             =   6345
         Width           =   1695
      End
      Begin VB.ComboBox cboColour7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "bookmst.frx":0129
         Left            =   5850
         List            =   "bookmst.frx":0136
         TabIndex        =   115
         Top             =   6720
         Width           =   1695
      End
      Begin VB.ComboBox cboColour8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         ItemData        =   "bookmst.frx":0162
         Left            =   5850
         List            =   "bookmst.frx":016F
         TabIndex        =   125
         Top             =   7080
         Width           =   1695
      End
      Begin VB.TextBox txtPCode6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8310
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   107
         Top             =   6405
         Width           =   570
      End
      Begin VB.TextBox txtPCode7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8310
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   117
         Top             =   6780
         Width           =   570
      End
      Begin VB.TextBox txtPCode8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8310
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   127
         Top             =   7110
         Width           =   570
      End
      Begin VB.TextBox txtPaper6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8910
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   108
         Top             =   6405
         Width           =   3300
      End
      Begin VB.TextBox txtPaper7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8910
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   118
         Top             =   6765
         Width           =   3300
      End
      Begin VB.TextBox txtPaper8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8910
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   128
         Top             =   7110
         Width           =   3300
      End
      Begin VB.TextBox txtTextSupp8 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   123
         Top             =   7080
         Width           =   480
      End
      Begin VB.TextBox txtTextSupp7 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   113
         Top             =   6720
         Width           =   465
      End
      Begin VB.TextBox txtTextSupp6 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3015
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   103
         Top             =   6345
         Width           =   480
      End
      Begin VB.ComboBox cbosupp6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2340
         TabIndex        =   102
         Top             =   6345
         Width           =   675
      End
      Begin VB.ComboBox cbosupp7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2355
         TabIndex        =   112
         Top             =   6720
         Width           =   675
      End
      Begin VB.ComboBox cbosupp8 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   2340
         TabIndex        =   122
         Top             =   7080
         Width           =   675
      End
      Begin VB.CommandButton Command16 
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
         Height          =   360
         Left            =   7590
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   6420
         Width           =   435
      End
      Begin VB.CommandButton Command15 
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
         Height          =   360
         Left            =   7590
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   6750
         Width           =   435
      End
      Begin VB.CommandButton Command14 
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
         Height          =   360
         Left            =   7590
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   7080
         Width           =   435
      End
      Begin VB.TextBox txtSpecimenLY 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8430
         MaxLength       =   80
         TabIndex        =   25
         Top             =   3240
         Width           =   720
      End
      Begin VB.TextBox txtInternalPrint 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   7170
         MaxLength       =   80
         TabIndex        =   23
         Top             =   2880
         Width           =   780
      End
      Begin VB.TextBox txtSpecimenCY 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   7170
         MaxLength       =   80
         TabIndex        =   24
         Top             =   3240
         Width           =   780
      End
      Begin VB.TextBox txtPrintedLY 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8430
         MaxLength       =   80
         TabIndex        =   22
         Top             =   2520
         Width           =   720
      End
      Begin VB.TextBox txtPrintedCY 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   7170
         MaxLength       =   80
         TabIndex        =   21
         Top             =   2520
         Width           =   780
      End
      Begin VB.TextBox txtPriceLY 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8430
         MaxLength       =   80
         TabIndex        =   20
         Top             =   2160
         Width           =   720
      End
      Begin VB.TextBox txtgpNo 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   7170
         MaxLength       =   80
         TabIndex        =   18
         Top             =   1800
         Width           =   765
      End
      Begin VB.TextBox txtPrice 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   7170
         MaxLength       =   80
         TabIndex        =   19
         Top             =   2160
         Width           =   780
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
         Height          =   330
         Left            =   5565
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   1215
         Width           =   495
      End
      Begin VB.ComboBox cboClass 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "bookmst.frx":019B
         Left            =   3405
         List            =   "bookmst.frx":019D
         TabIndex        =   5
         Top             =   1245
         Width           =   2115
      End
      Begin VB.CommandButton Command12 
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
         Height          =   330
         Left            =   4770
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   2550
         Width           =   495
      End
      Begin VB.CommandButton Command11 
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
         Height          =   330
         Left            =   4770
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   1950
         Width           =   495
      End
      Begin VB.ComboBox txtNegativeby 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1530
         TabIndex        =   9
         Top             =   2565
         Width           =   3255
      End
      Begin VB.ComboBox txtTypeSetter 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1530
         TabIndex        =   7
         Top             =   1950
         Width           =   3255
      End
      Begin VB.ComboBox txtWriter 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1530
         TabIndex        =   6
         Top             =   1590
         Width           =   3255
      End
      Begin VB.CommandButton Command10 
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
         Height          =   330
         Left            =   4770
         Style           =   1  'Graphical
         TabIndex        =   130
         Top             =   1590
         Width           =   495
      End
      Begin VB.ComboBox txtDivide 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "bookmst.frx":019F
         Left            =   3750
         List            =   "bookmst.frx":01A1
         TabIndex        =   13
         Top             =   3540
         Width           =   1035
      End
      Begin VB.CommandButton Command9 
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
         Height          =   360
         Left            =   7590
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   6030
         Width           =   435
      End
      Begin VB.CommandButton Command8 
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
         Height          =   348
         Left            =   7590
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   5676
         Width           =   435
      End
      Begin VB.CommandButton Command7 
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
         Height          =   348
         Left            =   7590
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   5340
         Width           =   435
      End
      Begin VB.CommandButton Command5 
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
         Height          =   348
         Left            =   7590
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4992
         Width           =   435
      End
      Begin VB.CommandButton Command4 
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
         Height          =   348
         Left            =   7590
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   4644
         Width           =   435
      End
      Begin VB.TextBox txtInnbright 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   8115
         MaxLength       =   80
         TabIndex        =   94
         Top             =   4680
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox txttextbright 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   8115
         MaxLength       =   80
         TabIndex        =   92
         Top             =   5040
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox txtExambright 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   8115
         MaxLength       =   80
         TabIndex        =   90
         Top             =   5430
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox txtSuppbright 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   8115
         MaxLength       =   80
         TabIndex        =   88
         Top             =   5745
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox txtTitlebright 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   8115
         MaxLength       =   80
         TabIndex        =   86
         Top             =   5985
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.ComboBox cbotitle 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2355
         TabIndex        =   89
         Top             =   6030
         Width           =   675
      End
      Begin VB.ComboBox cbosupp 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2355
         TabIndex        =   51
         Top             =   5685
         Width           =   675
      End
      Begin VB.ComboBox cboExam 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   2355
         TabIndex        =   44
         Top             =   5370
         Width           =   675
      End
      Begin VB.ComboBox cbotext 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "bookmst.frx":01A3
         Left            =   2355
         List            =   "bookmst.frx":01A5
         TabIndex        =   37
         Top             =   5040
         Width           =   675
      End
      Begin VB.ComboBox cboInner 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "bookmst.frx":01A7
         Left            =   2355
         List            =   "bookmst.frx":01A9
         TabIndex        =   30
         Top             =   4680
         Width           =   675
      End
      Begin VB.TextBox txtInForms 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   31
         Top             =   4680
         Width           =   480
      End
      Begin VB.TextBox txtTextForms 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   38
         Top             =   5040
         Width           =   480
      End
      Begin VB.TextBox txtTextExam 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   45
         Top             =   5370
         Width           =   480
      End
      Begin VB.TextBox txtTextSupp 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   52
         Top             =   5685
         Width           =   465
      End
      Begin VB.TextBox txtTextTitle 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   91
         Top             =   6030
         Width           =   480
      End
      Begin VB.TextBox txtTitlePaper 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8910
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   98
         Top             =   6030
         Width           =   3300
      End
      Begin VB.TextBox txtSuppPaper 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8910
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   84
         Top             =   5685
         Width           =   3300
      End
      Begin VB.TextBox txtExamPaper 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8910
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   83
         Top             =   5355
         Width           =   3300
      End
      Begin VB.TextBox txtTextPaper 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8910
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   82
         Top             =   5010
         Width           =   3300
      End
      Begin VB.TextBox txtInnerPaper 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8910
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   81
         Top             =   4680
         Width           =   3300
      End
      Begin VB.TextBox txtTitlePCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8310
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   97
         Top             =   6030
         Width           =   570
      End
      Begin VB.TextBox txtSuppPCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8310
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   80
         Top             =   5685
         Width           =   570
      End
      Begin VB.TextBox txtExamPCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8310
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   79
         Top             =   5370
         Width           =   570
      End
      Begin VB.TextBox txtTextPCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8310
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   78
         Top             =   4980
         Width           =   570
      End
      Begin VB.ComboBox cboTitleColour 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "bookmst.frx":01AB
         Left            =   5850
         List            =   "bookmst.frx":01B8
         TabIndex        =   95
         Top             =   6030
         Width           =   1695
      End
      Begin VB.ComboBox cboSuppColour 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "bookmst.frx":01E4
         Left            =   5850
         List            =   "bookmst.frx":01F1
         TabIndex        =   54
         Top             =   5685
         Width           =   1695
      End
      Begin VB.ComboBox cboExamColour 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "bookmst.frx":021D
         Left            =   5850
         List            =   "bookmst.frx":022A
         TabIndex        =   47
         Top             =   5370
         Width           =   1695
      End
      Begin VB.ComboBox cboTextColour 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "bookmst.frx":0256
         Left            =   5850
         List            =   "bookmst.frx":0263
         TabIndex        =   40
         Top             =   5040
         Width           =   1695
      End
      Begin VB.ComboBox cboTitlePrint 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3570
         TabIndex        =   93
         Top             =   6030
         Width           =   2235
      End
      Begin VB.ComboBox cboSuppPrint 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3570
         TabIndex        =   53
         Top             =   5685
         Width           =   2235
      End
      Begin VB.ComboBox cboExamPrint 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3570
         TabIndex        =   46
         Top             =   5370
         Width           =   2235
      End
      Begin VB.ComboBox cboTextPrint 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3570
         TabIndex        =   39
         Top             =   5040
         Width           =   2235
      End
      Begin VB.ComboBox cboInnerPrint 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3570
         TabIndex        =   32
         Top             =   4680
         Width           =   2235
      End
      Begin VB.ComboBox cboBinder 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1530
         TabIndex        =   10
         Top             =   2865
         Width           =   3255
      End
      Begin VB.TextBox txtEdition 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1530
         MaxLength       =   40
         TabIndex        =   8
         Top             =   2265
         Width           =   3240
      End
      Begin VB.CommandButton Command6 
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
         Height          =   330
         Left            =   4770
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   3180
         Width           =   495
      End
      Begin VB.ComboBox cboLemination 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   1530
         TabIndex        =   11
         Top             =   3180
         Width           =   3255
      End
      Begin VB.ComboBox cboColour 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         ItemData        =   "bookmst.frx":028F
         Left            =   5850
         List            =   "bookmst.frx":029C
         TabIndex        =   33
         Top             =   4680
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
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
         Height          =   330
         Left            =   4770
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   3540
         Width           =   495
      End
      Begin VB.TextBox txtInnerPCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   8310
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   75
         Top             =   4680
         Width           =   570
      End
      Begin VB.TextBox txtHeadData5 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1875
         MaxLength       =   80
         TabIndex        =   87
         Top             =   6030
         Width           =   465
      End
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   288
         Left            =   1872
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   74
         Top             =   8916
         Width           =   465
      End
      Begin VB.TextBox txtHeadData4 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1875
         MaxLength       =   80
         TabIndex        =   50
         Top             =   5685
         Width           =   465
      End
      Begin VB.TextBox txtHeadData3 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1875
         MaxLength       =   80
         TabIndex        =   43
         Top             =   5370
         Width           =   465
      End
      Begin VB.TextBox txtHeadData2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1875
         MaxLength       =   80
         TabIndex        =   36
         Top             =   5040
         Width           =   465
      End
      Begin VB.TextBox txtHeadData1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1875
         MaxLength       =   80
         TabIndex        =   29
         Top             =   4680
         Width           =   465
      End
      Begin VB.TextBox txtdes 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1530
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1230
         Width           =   1230
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   708
         Left            =   270
         ScaleHeight     =   708
         ScaleWidth      =   10248
         TabIndex        =   71
         Top             =   9432
         Width           =   10245
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   630
            Left            =   45
            Picture         =   "bookmst.frx":02C8
            Style           =   1  'Graphical
            TabIndex        =   73
            Top             =   75
            Width           =   1140
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Next"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   2355
            Style           =   1  'Graphical
            TabIndex        =   143
            Top             =   75
            Width           =   1260
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "P&revious"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   1215
            Style           =   1  'Graphical
            TabIndex        =   141
            Top             =   75
            Width           =   1140
         End
         Begin VB.CommandButton cmdSEarch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "S&earch"
            Height          =   630
            Left            =   6855
            Picture         =   "bookmst.frx":0EAC
            Style           =   1  'Graphical
            TabIndex        =   137
            Top             =   75
            Width           =   1140
         End
         Begin VB.CommandButton save 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save/Edit"
            Height          =   630
            Left            =   3645
            Picture         =   "bookmst.frx":1A90
            Style           =   1  'Graphical
            TabIndex        =   131
            Top             =   60
            Width           =   1140
         End
         Begin VB.CommandButton Del 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Delete"
            Height          =   630
            Left            =   4815
            Picture         =   "bookmst.frx":2674
            Style           =   1  'Graphical
            TabIndex        =   133
            Top             =   75
            Width           =   1020
         End
         Begin VB.CommandButton Abandon 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Abandon"
            Height          =   630
            Left            =   5835
            Picture         =   "bookmst.frx":3258
            Style           =   1  'Graphical
            TabIndex        =   135
            Top             =   75
            Width           =   1020
         End
         Begin VB.CommandButton REPORTCD 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   630
            Left            =   7980
            Picture         =   "bookmst.frx":37E2
            Style           =   1  'Graphical
            TabIndex        =   139
            Top             =   60
            Width           =   1080
         End
         Begin VB.CommandButton close 
            BackColor       =   &H00FFFFFF&
            Cancel          =   -1  'True
            Caption         =   "&Exit"
            Height          =   630
            Left            =   9090
            Picture         =   "bookmst.frx":43C6
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   75
            Width           =   1065
         End
      End
      Begin VB.TextBox bsize 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1530
         MaxLength       =   80
         TabIndex        =   12
         Top             =   3495
         Width           =   1680
      End
      Begin VB.TextBox bunit 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   70
         Top             =   8916
         Width           =   540
      End
      Begin VB.TextBox bwastage 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   1530
         MaxLength       =   4
         TabIndex        =   14
         Top             =   3780
         Width           =   885
      End
      Begin VB.ComboBox ftype 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   288
         ItemData        =   "bookmst.frx":4FAA
         Left            =   1530
         List            =   "bookmst.frx":4FB4
         TabIndex        =   15
         Top             =   4095
         Width           =   1620
      End
      Begin VB.TextBox bfont 
         Height          =   255
         Left            =   7770
         TabIndex        =   69
         Top             =   1320
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   8100
         MaxLength       =   80
         TabIndex        =   68
         Top             =   6360
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   8100
         MaxLength       =   80
         TabIndex        =   67
         Top             =   7080
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   8100
         MaxLength       =   80
         TabIndex        =   66
         Top             =   6720
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox txtISBN 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3675
         MaxLength       =   20
         TabIndex        =   1
         Top             =   540
         Width           =   1740
      End
      Begin VB.PictureBox picOriginal 
         AutoRedraw      =   -1  'True
         Height          =   4275
         Left            =   9480
         MousePointer    =   2  'Cross
         ScaleHeight     =   4224
         ScaleWidth      =   3744
         TabIndex        =   65
         Top             =   120
         Visible         =   0   'False
         Width           =   3795
      End
      Begin VB.ComboBox txtBookCode 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   1530
         Style           =   1  'Simple Combo
         TabIndex        =   0
         Top             =   480
         Width           =   1515
      End
      Begin Crystal.CrystalReport cr 
         Left            =   12936
         Top             =   8676
         _ExtentX        =   593
         _ExtentY        =   593
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   12912
         Top             =   8220
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remark :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3684
         TabIndex        =   249
         Top             =   8928
         Width           =   876
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   17
         Left            =   36
         TabIndex        =   247
         Top             =   8508
         Width           =   168
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   16
         Left            =   36
         TabIndex        =   244
         Top             =   8184
         Width           =   168
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   15
         Left            =   36
         TabIndex        =   241
         Top             =   7860
         Width           =   168
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   3
         Left            =   120
         TabIndex        =   238
         Top             =   7536
         Width           =   84
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FirmName  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   5460
         TabIndex        =   199
         Top             =   540
         Width           =   1260
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Binding"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6144
         TabIndex        =   189
         Top             =   4080
         Width           =   792
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trim Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3108
         TabIndex        =   186
         Top             =   4092
         Width           =   996
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Press F4 for English "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   372
         Left            =   2736
         TabIndex        =   184
         Top             =   10236
         Width           =   2652
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Press F3 for Hindi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   372
         Left            =   360
         TabIndex        =   183
         Top             =   10236
         Visible         =   0   'False
         Width           =   2352
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Return :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5490
         TabIndex        =   182
         Top             =   3615
         Width           =   810
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   14
         Left            =   105
         TabIndex        =   181
         Top             =   7125
         Width           =   90
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   105
         TabIndex        =   180
         Top             =   5700
         Width           =   90
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   105
         TabIndex        =   179
         Top             =   6060
         Width           =   90
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   11
         Left            =   105
         TabIndex        =   178
         Top             =   6375
         Width           =   90
      End
      Begin VB.Label lblspecimentLY 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "LY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9150
         TabIndex        =   177
         Top             =   3300
         Width           =   285
      End
      Begin VB.Label lblspecimentCY 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
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
         Left            =   7950
         TabIndex        =   176
         Top             =   3300
         Width           =   495
      End
      Begin VB.Label lblTotalPrintedLY 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "LY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9150
         TabIndex        =   175
         Top             =   2580
         Width           =   285
      End
      Begin VB.Label lbltotalprintedCY 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
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
         Left            =   7950
         TabIndex        =   174
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblPriceLY 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "LY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   9150
         TabIndex        =   173
         Top             =   2160
         Width           =   285
      End
      Begin VB.Label lblPriceCY 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
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
         Left            =   7950
         TabIndex        =   172
         Top             =   2205
         Width           =   495
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Specimen :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5490
         TabIndex        =   171
         Top             =   3240
         Width           =   1170
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Initial  Print Run. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5490
         TabIndex        =   170
         Top             =   2940
         Width           =   1800
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Printed :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5490
         TabIndex        =   169
         Top             =   2580
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gp. Book No. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5490
         TabIndex        =   168
         Top             =   1860
         Width           =   1875
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5490
         TabIndex        =   167
         Top             =   2220
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   2790
         TabIndex        =   166
         Top             =   1290
         Width           =   825
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   $"bookmst.frx":4FC4
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
         Height          =   210
         Left            =   105
         TabIndex        =   165
         Top             =   4440
         Width           =   12660
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Binder"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   10
         Left            =   192
         TabIndex        =   164
         Top             =   2880
         Width           =   1296
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   9
         Left            =   192
         TabIndex        =   163
         Top             =   2268
         Width           =   1332
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Lamination"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   192
         TabIndex        =   162
         Top             =   3180
         Width           =   1140
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   105
         TabIndex        =   161
         Top             =   6750
         Width           =   90
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   7
         Left            =   192
         TabIndex        =   160
         Top             =   1596
         Width           =   780
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type Setters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   6
         Left            =   150
         TabIndex        =   159
         Top             =   1965
         Width           =   1350
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Negative By"
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
         Height          =   192
         Index           =   5
         Left            =   192
         TabIndex        =   158
         Top             =   2580
         Width           =   1056
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   192
         TabIndex        =   157
         Top             =   528
         Width           =   1176
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3(Exa)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   156
         Top             =   5385
         Width           =   450
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2(tex)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   155
         Top             =   5040
         Width           =   390
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1(inn)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   154
         Top             =   4695
         Width           =   390
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1200
         TabIndex        =   153
         Top             =   8916
         Width           =   672
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "/"
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
         Left            =   3510
         TabIndex        =   152
         Top             =   3540
         Width           =   90
      End
      Begin VB.Label mname 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Master :-"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   14.4
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   75
         TabIndex        =   151
         Top             =   105
         Width           =   2310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No of Pages "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   192
         TabIndex        =   150
         Top             =   1260
         Width           =   1248
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Name "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   192
         TabIndex        =   149
         Top             =   888
         Width           =   1296
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F2 for search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   3690
         TabIndex        =   148
         Top             =   75
         Visible         =   0   'False
         Width           =   2880
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Size "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   150
         TabIndex        =   147
         Top             =   3480
         Width           =   525
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Form "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2376
         TabIndex        =   146
         Top             =   8916
         Width           =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wastage %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   192
         TabIndex        =   145
         Top             =   3768
         Width           =   1200
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   144
         Top             =   4095
         Width           =   915
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   840
         Left            =   228
         Top             =   9408
         Width           =   10368
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ISBN :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   2
         Left            =   3090
         TabIndex        =   142
         Top             =   540
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ref As ADODB.Recordset
Dim b As Boolean
Dim RS As New ADODB.Recordset
Dim rs11 As New ADODB.Recordset
Dim rs_11 As New ADODB.Recordset
Dim totalSum As Long
Dim d1 As Integer
Dim f As New ADODB.Recordset
Private Sub ABANDON_Click()

    Clearvalue
    txtBookCode.SetFocus

End Sub
Sub calcform()

End Sub
Private Sub bsize_GotFocus()
Label5.Visible = True
If PopUpValue1 <> "" Then
   bsize.text = PopUpValue1
   'txtdes.Text = PopUpValue2
End If

PopUpValue1 = ""
PopUpValue2 = ""

End Sub
Private Sub bsize_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
      popuplist1 "Select size1 as [Size],size_info as [Size Info] from SizeMaster where size1 <> '' and  " & stringyear, con
End If

End Sub
Private Sub bwastage_KeyPress(KeyAscii As Integer)
 Dim bb As Boolean
 bb = val_int(bwastage, KeyAscii)
 If bb = False Then
    KeyAscii = 0
 End If
 
End Sub
Private Sub cboClass_Click()

If cboClass.text = "" Then Exit Sub
If RS.State = 1 Then RS.close
RS.Open "select max(gpno) from BookMaster where class='" & cboClass.text & "' and  " & stringyear, con, adOpenDynamic, adLockOptimistic
If IsNull(RS(0)) Then
   txtgpNo.text = 1
Else
   txtgpNo.text = RS(0) + 1
End If



End Sub
Sub searchBookGp()
  
If rs_11.State = 1 Then rs_11.close
rs_11.Open "Select * from BookMaster where class='" & cboClass.text & "' and  " & stringyear & " order by BookNo", con, adOpenKeyset, adLockReadOnly
If rs_11.EOF = False Then
'rs_11.MoveFirst
'txtBookCode.Text = rs_11!BookNo & ""
'rs_11.Find "BookNo='" & txtBookCode.Text & "'"
'If Not (rs_11.EOF) Then
'    lstfocus
' If rs_11.EOF Then
'    Exit Sub
' End If
'End If

End If

End Sub


Private Sub cboClass_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  If cboClass.text = "" Then
     cboClass.SetFocus
  End If
End If

End Sub
Private Sub cboClass_LostFocus()

If rs_11.State = 1 Then rs_11.close
rs_11.Open "Select * from BookMaster where class='" & cboClass.text & "' and  " & stringyear, con, adOpenKeyset, adLockReadOnly
If rs_11.EOF = False Then
   'txtBookCode = rs_11!bookNo
End If
'rs_11.Open "Select * from BookMaster", con, adOpenKeyset, adLockReadOnly

End Sub

Private Sub cboClass1_Click()
fillGrid
End Sub

Private Sub cboExam_Click()
calForm
End Sub

Private Sub cboExam_LostFocus()
calForm
End Sub

Private Sub cboInner_Click()
calForm
End Sub
Private Sub cbosupp_Click()
calForm
End Sub

Private Sub cbosupp_LostFocus()
calForm
End Sub

Private Sub cbosupp10_Change()
calForm
totalValue
End Sub

Private Sub cbosupp11_Change()
calForm
totalValue
End Sub

Private Sub cbosupp12_Change()
calForm
totalValue
End Sub

Private Sub cbosupp6_LostFocus()
calForm
End Sub

Private Sub cbosupp9_Change()
calForm
 totalValue
End Sub

Private Sub cbotext_Click()
calForm
End Sub


Private Sub cbotext_LostFocus()
calForm
End Sub

Private Sub cbotitle_Click()
calForm
End Sub

Private Sub close_Click()
Unload Me
End Sub

Private Sub cmd1_Click()
txtHead1 = ""
txtHeadData1 = ""
cboInner = ""
txtInForms = ""
cboInnerPrint = ""
cboColour = ""
txtInnerPCode = ""
txtInnerPaper = ""
End Sub

Private Sub cmd10_Click()
txtHead10 = ""
txtHeadData10 = ""
cbosupp10 = ""
txtTextSupp10 = ""
cboPrinter10 = ""
cboColour10 = ""
txtPCode10 = ""
txtPaper10 = ""
End Sub

Private Sub cmd11_Click()
    txtHead11 = ""
    txtHeadData11 = ""
    cbosupp11 = ""
    txtTextSupp11 = ""
    cboPrinter11 = ""
    cboColour11 = ""
    txtPCode11 = ""
    txtPaper11 = ""
End Sub

Private Sub cmd12_Click()

txtHead12 = ""
txtHeadData12 = ""
cbosupp12 = ""
txtTextSupp12 = ""
cboPrinter12 = ""
cboColour12 = ""
txtPCode12 = ""
txtPaper12 = ""
 
End Sub

Private Sub cmd3_Click()

txtHead3 = ""
txtHeadData3 = ""
cboExam = ""
txtTextExam = ""
cboExamPrint = ""
cboExamColour = ""
txtExamPCode = ""
txtExamPaper = ""

End Sub

Private Sub cmd4_Click()
txtHead4 = ""
txtHeadData4 = ""
cbosupp = ""
txtTextSupp = ""
cboSuppPrint = ""
cboSuppColour = ""
txtSuppPCode = ""
txtSuppPaper = ""
End Sub

Private Sub cmd5_Click()
txtHead6 = ""
txtHeadData6 = ""
cbosupp6 = ""
txtTextSupp6 = ""
cboPrinter6 = ""
cboColour6 = ""
txtPCode6 = ""
txtPaper6 = ""
End Sub

Private Sub cmd6_Click()
txtHead7 = ""
txtHeadData7 = ""
cbosupp7 = ""
txtTextSupp7 = ""
cboPrinter7 = ""
cboColour7 = ""
txtPCode7 = ""
txtPaper7 = ""
End Sub

Private Sub cmd7_Click()
txtHead8 = ""
txtHeadData8 = ""
cbosupp8 = ""
txtTextSupp8 = ""
cboPrinter8 = ""
cboColour8 = ""
txtPCode8 = ""
txtPaper8 = ""
End Sub

Private Sub cmd8_Click()
txtHead5 = ""
txtHeadData5 = ""
cbotitle = ""
txtTextTitle = ""
cboTitlePrint = ""
cboTitleColour = ""
txtTitlePCode = ""
txtTitlePaper = ""
End Sub

Private Sub cmd9_Click()

txtHead9 = ""
txtHeadData9 = ""
cbosupp9 = ""
txtTextSupp9 = ""
cboPrinter9 = ""
cboColour9 = ""
txtPCode9 = ""
txtPaper9 = ""


End Sub

Private Sub cmdAdd_Click()
Clearvalue
txtBookCode.SetFocus
End Sub
Sub calForm()

On Error Resume Next
txtInForms = (Val(txtHeadData1.text) / Val(cboInner.text))
txtTextForms = (Val(txtHeadData2.text) / IIf(cbotext.text = "", 0, cbotext.text))

If Val(cboExam.text) > 0 Then
txtTextExam = (Val(txtHeadData3.text) / Val(cboExam.text))
End If

If Val(cbosupp.text) > 0 Then
txtTextSupp = (Val(txtHeadData4.text) / Val(cbosupp.text))
End If

If Val(cbotitle.text) > 0 Then
txtTextTitle = (Val(txtHeadData5.text) / Val(cbotitle.text))
End If

If Val(cbosupp6.text) > 0 Then
txtTextSupp6 = (Val(txtHeadData6.text) / Val(cbosupp6.text))
End If

If Val(cbosupp7.text) > 0 Then
txtTextSupp7 = (Val(txtHeadData7.text) / Val(cbosupp7.text))
End If

txtTextSupp8 = (Val(txtHeadData8.text) / Val(cbosupp8.text))


If Val(cbosupp9.text) > 0 Then
txtTextSupp9 = (Val(txtHeadData9.text) / Val(cbosupp9.text))
End If

If Val(cbosupp10.text) > 0 Then
txtTextSupp10 = (Val(txtHeadData10.text) / Val(cbosupp10.text))
End If


If Val(cbosupp11.text) > 0 Then
txtTextSupp11 = (Val(txtHeadData11.text) / Val(cbosupp11.text))
End If


If Val(cbosupp12.text) > 0 Then
txtTextSupp12 = (Val(txtHeadData12.text) / Val(cbosupp12.text))
End If
'--------------

txtInForms = Round(txtInForms.text, 3)
txtTextForms = Round(txtTextForms.text, 3)
txtTextExam = Round(txtTextExam.text, 3)
txtTextSupp = Round(txtTextSupp.text, 3)
txtTextTitle = Round(txtTextTitle.text, 3)



End Sub

Private Sub cmdBookNoAd_Click()
 If d1 = 1 Then
    Screen.MousePointer = vbHourglass
    Frame1.Height = 6750
    Frame1.Width = 7770
    Frame1.top = 840
    Frame1.Left = 120
    Frame1.Visible = True
    fillGrid
    d1 = d1 + 1
    Screen.MousePointer = vbDefault
    
 Else
    Frame1.Visible = False
    d1 = 1
 End If
End Sub
Sub fillGrid()
 
 If f.State = 1 Then f.close
 If cboClass1.text = "" Then
    f.Open "select BookNo,Book,Class,GPNo from BookMaster where  " & stringyear & " order by BookNo", con, adOpenDynamic, adLockOptimistic
 Else
    f.Open "select BookNo,Book,Class,GPNo from BookMaster where class='" & cboClass1.text & "' and  " & stringyear & " order by BookNo", con, adOpenDynamic, adLockOptimistic
 End If
 
 Set vs.DataSource = f

End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdClose_Click()
frmBookList.Visible = False
End Sub

Private Sub cmdExitfrm_Click()
Frame1.Visible = False
End Sub

Private Sub cmdok_Click()
  
On Error GoTo aa101

  If MsgBox("Want to update ?", vbQuestion + vbYesNo) = vbYes Then
     For J = 1 To vs.rows - 1
     If VS_bk.TextMatrix(J, 0) <> "" Then
       con.Execute "update BookMaster set price='" & VS_bk.TextMatrix(J, 2) & "' where BookNo = '" & VS_bk.TextMatrix(J, 0) & "'"
       
       If RS.State = 1 Then RS.close
       RS.Open "select * from BookDetails where BookNo = '" & VS_bk.TextMatrix(J, 0) & "'", con, adOpenDynamic, adLockOptimistic
       If RS.EOF = False Then
          con.Execute "update BookDetails set price='" & VS_bk.TextMatrix(J, 2) & "' where BookNo = '" & VS_bk.TextMatrix(J, 0) & "'"
       Else
          RS.AddNew
          RS!bookNo = VS_bk.TextMatrix(J, 0)
          RS!fyear = main.session
          RS!Price = Val(VS_bk.TextMatrix(J, 2))
          RS.update
       End If
     End If
     Next
  End If
  
  
Exit Sub
aa101:

MsgBox "" & err.DESCRIPTION

  
  
  
End Sub

Private Sub cmdSearch_Click()

'If cboClass <> "" Then
'    popuplist1 "Select BookNo as [Book Code],book as [Book Name], book_info as [Book Information] from BookMaster where book <> '' and bookfont= '" + bfont.Text + "' and class='" & cboClass & "' and " & stringyear & " order by BookNo", con, , bfont
'End If

    searchType = "party"
    'popuplist_client "Select BookNo as [Book Code],book as [Book Name] from BookMaster order by BookNo", con
    popuplistFast "Select BookNo as [Book Code],book as [Book Name],[Class] as GroupCode from BookMaster order by BookNo", con, , , "book"
    
    

End Sub
Private Sub cmdSearch_GotFocus()

If PopUpValue1 = "" Then Exit Sub

Label5.Visible = True
Set ref = New ADODB.Recordset
If ref.State = 1 Then ref.close
ref.Open "Select * from bookMaster where bookno = '" + PopUpValue1 + "' order by BookNo", con, adOpenDynamic, adLockOptimistic, adCmdText
If ref.RecordCount > 0 Then
   lstfocus
   txtBookCode.SetFocus
End If
PopUpValue1 = ""
PopUpValue2 = ""



End Sub

Private Sub cmdUpdate_Click()

Screen.MousePointer = vbHourglass

Dim con_conven As New ADODB.Connection
Dim rs_spec As New ADODB.Recordset

If RS.State = 1 Then RS.close


RS.Open "select Path2 from DataBasePath", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
    With con_conven
    .CursorLocation = adUseClient
    .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & RS(0)
    .Open
    .CursorLocation = adUseClient
    End With
End If




If RS.State = 1 Then RS.close
RS.Open "select sum(QUANTITY),Bookcode from invoiceb where " & stringyear & "  group by Bookcode", con_conven, adOpenDynamic, adLockOptimistic
While RS.EOF = False
 con.Execute "update BookMaster set Sepimen_CY=" & RS(0) & " where " & stringyear & " and bookno='" & RS(1) & "'"
RS.MoveNext
Wend


If RS.State = 1 Then RS.close
RS.Open "select sum(QUANTITY),Bookcode from CREDITB where " & stringyear & " group by Bookcode", con_conven, adOpenDynamic, adLockOptimistic
While RS.EOF = False
 con.Execute "update BookMaster set Return_CY=" & RS(0) & " where " & stringyear & " and bookno='" & RS(1) & "'"
RS.MoveNext
Wend




Screen.MousePointer = vbDefault



End Sub

Private Sub cmdUpDatePrice_Click()

Screen.MousePointer = vbHourglass

VS_bk.Clear
frmBookList.Visible = True
If RS.State = 1 Then RS.close
If RS.State = 1 Then RS.close
RS.Open "select bookNo,book,price from BookMaster where [class]='" & cboClass & "' and " & stringyear & " order by bookno"
VS_bk.rows = 1
For J = 0 To RS.RecordCount - 1
    VS_bk.rows = VS_bk.rows + 1
    VS_bk.TextMatrix(J, 0) = RS!bookNo
    VS_bk.TextMatrix(J, 1) = RS!Book
    VS_bk.TextMatrix(J, 2) = RS!Price & ""
    RS.MoveNext
Next


VS_bk.FormatString = "BookCode|Book Name|Price"
VS_bk.ColWidth(0) = 1000
VS_bk.ColWidth(1) = 4500
VS_bk.ColWidth(2) = 1000

Screen.MousePointer = vbDefault

End Sub

Private Sub cnd2_Click()
txtHead2 = ""
txtHeadData2 = ""
cbotext = ""
txtTextForms = ""
cboTextPrint = ""
cboTextColour = ""
txtTextPCode = ""
txtTextPaper = ""
End Sub

Private Sub Command1_Click()

On Error GoTo err1

Screen.MousePointer = vbHourglass
        
        If txtBookCode.text = "" Then
           MsgBox "Please Enter Book Code !!", vbInformation
           txtBookCode.SetFocus
           Exit Sub
        End If
        
        
        rs_11.MoveFirst
        
        rs_11.Find "BookNo='" & Trim(txtBookCode.text) & "'"
        If rs_11.EOF = False Then
           rs_11.MovePrevious
           If rs_11.EOF = False Then
              txtBookCode.text = rs_11!bookNo
           End If
           
           'SearchBook
           lstfocus
        
Screen.MousePointer = vbDefault
        
        If rs_11.EOF Then
           Exit Sub
        End If
        End If
        
Exit Sub
err1:

If err.Number = 3021 Then
Screen.MousePointer = vbDefault
MsgBox "Record not found ...", vbCritical
End If
        
End Sub


Private Sub Command10_Click()
HeadTbl = "Author"
frmMasters.Show 1

End Sub

Private Sub Command11_Click()

HeadTbl = "typesetter"
frmMasters.Show 1

End Sub

Private Sub Command12_Click()

HeadTbl = "negative"
frmMasters.Show 1

End Sub

Private Sub Command13_Click()

HeadTbl = "class"
frmMasters.Show 1

End Sub

Private Sub Command14_9_Click()
 value = "Select papermaker_name as [Paper Name],Eco,SizeValue1 + ' X ' +SizeValue2 as Size, GSM,papermaker_Id as Code from papermakemaster where papermaker_id <> '' and  " & stringyear
 popuplist1 value, con
End Sub

Private Sub Command14_9_GotFocus()
If PopUpValue1 <> "" Then
  
  txtPCode9.text = popupvalue5
  txtPaper9.text = PopUpValue1 & " , " & PopUpValue3 & " , " & popupvalue4 & " , " & popupvalue5 & " , " & PopUpValue2
  

  txtHead10.SetFocus
  
End If
  
  PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""
  popupvalue4 = ""
  popupvalue5 = ""
End Sub

Private Sub Command14_Click()
 value = "Select papermaker_name as [Paper Name],Eco,SizeValue1 + ' X ' +SizeValue2 as Size, GSM,papermaker_Id as Code from papermakemaster where papermaker_id <> '' and  " & stringyear
 popuplist1 value, con

End Sub

Private Sub Command14_GotFocus()
If PopUpValue1 <> "" Then
  
  txtPCode8.text = popupvalue5
  txtPaper8.text = PopUpValue1 & " , " & PopUpValue3 & " , " & popupvalue4 & " , " & popupvalue5 & " , " & PopUpValue2
  
  txtHead9.SetFocus
  
End If
  
  PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""
  popupvalue4 = ""
  popupvalue5 = ""
End Sub
Private Sub Command15_Click()
 value = "Select papermaker_name as [Paper Name],Eco,SizeValue1 + ' X ' +SizeValue2 as Size, GSM,papermaker_Id as Code from papermakemaster where papermaker_id <> '' and  " & stringyear
 popuplist1 value, con

End Sub

Private Sub Command15_GotFocus()
If PopUpValue1 <> "" Then
  
  txtPCode7.text = popupvalue5
  txtPaper7.text = PopUpValue1 & " , " & PopUpValue3 & " , " & popupvalue4 & " , " & popupvalue5 & " , " & PopUpValue2
  
  txtHead8.SetFocus
  
End If
  
  PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""
  popupvalue4 = ""
  popupvalue5 = ""
End Sub

Private Sub Command16_Click()
 value = "Select papermaker_name as [Paper Name],Eco,SizeValue1 + ' X ' +SizeValue2 as Size, GSM,papermaker_Id as Code from papermakemaster where papermaker_id <> '' and  " & stringyear
 popuplist1 value, con

End Sub

Private Sub Command16_GotFocus()
  
If PopUpValue1 <> "" Then
  
  txtPCode6.text = popupvalue5
  txtPaper6.text = PopUpValue1 & " , " & PopUpValue3 & " , " & popupvalue4 & " , " & popupvalue5 & " , " & PopUpValue2
  
  
  txtHead7.SetFocus
  
End If
  
  PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""
  popupvalue4 = ""
  popupvalue5 = ""

End Sub

Private Sub Command17_Click()


Dim kk11 As Integer
kk11 = (Val(vs.TextMatrix(vs.RowSel, 4)) + 1)

If RS.State = 1 Then RS.close
RS.Open "select GpNo from BookMaster and  " & stringyear & " order by GpNo desc"
While RS.EOF = False
If RS(0) >= kk11 Then
   con.Execute "update BookMaster set GpNo=" & (RS(0) + 1) & " where GpNo=" & RS(0) & " and  " & stringyear
End If
RS.MoveNext
Wend


con.Execute "update BookMaster set GpNo=" & Val(txtGpno1.text) & " where GpNo=" & vs.TextMatrix(vs.RowSel, 4) & " and  " & stringyear

MsgBox "Increament Successfully ...", vbInformation
    
End Sub

Private Sub Command18_Click()
HeadTbl = "bkpart"
frmMasters.Show 1
End Sub

Private Sub Command19_Click()
HeadTbl = "trimsize"
frmMasters.Show 1
End Sub

'Private Sub Command18_Click()
'End Sub

Private Sub Command2_Click()
        
    On Error GoTo err1
        
    Screen.MousePointer = vbHourglass
        
        If txtBookCode.text = "" Then
           MsgBox "Please Enter Book Code !!", vbInformation
           txtBookCode.SetFocus
           Exit Sub
        End If
        
        rs_11.MoveFirst
        
        rs_11.Find "BookNo='" & txtBookCode.text & "'"
        If Not (rs_11.EOF) Then
           rs_11.MoveNext
           If rs_11.EOF = False Then
              txtBookCode.text = rs_11!bookNo
           End If
             
           lstfocus
           Screen.MousePointer = vbDefault
            
        If rs_11.EOF Then
           Exit Sub
        End If
        End If
        
        
Exit Sub
err1:

If err.Number = 3021 Then
Screen.MousePointer = vbDefault
MsgBox "Record not found ...", vbCritical
End If
        
End Sub

Private Sub Command20_Click()
HeadTbl = "binding"
frmMasters.Show 1
End Sub

Private Sub Command24_10_Click()
 value = "Select papermaker_name as [Paper Name],Eco,SizeValue1 + ' X ' +SizeValue2 as Size, GSM,papermaker_Id as Code from papermakemaster where papermaker_id <> '' and  " & stringyear
 popuplist1 value, con
End Sub

Private Sub Command24_10_GotFocus()
If PopUpValue1 <> "" Then
  
  txtPCode10.text = popupvalue5
  txtPaper10.text = PopUpValue1 & " , " & PopUpValue3 & " , " & popupvalue4 & " , " & popupvalue5 & " , " & PopUpValue2
  
  txtHead11.SetFocus
  ''txtHeadData10.SetFocus
  
End If
  
  PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""
  popupvalue4 = ""
  popupvalue5 = ""
End Sub

Private Sub Command24_11_Click()
 value = "Select papermaker_name as [Paper Name],Eco,SizeValue1 + ' X ' +SizeValue2 as Size, GSM,papermaker_Id as Code from papermakemaster where papermaker_id <> '' and  " & stringyear
 popuplist1 value, con

End Sub

Private Sub Command24_11_GotFocus()
If PopUpValue1 <> "" Then
  
  txtPCode11.text = popupvalue5
  txtPaper11.text = PopUpValue1 & " , " & PopUpValue3 & " , " & popupvalue4 & " , " & popupvalue5 & " , " & PopUpValue2
  
  txtHead12.SetFocus
  ''txtHeadData10.SetFocus
  
End If
  
  PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""
  popupvalue4 = ""
  popupvalue5 = ""
End Sub

Private Sub Command24_12_Click()
 value = "Select papermaker_name as [Paper Name],Eco,SizeValue1 + ' X ' +SizeValue2 as Size, GSM,papermaker_Id as Code from papermakemaster where papermaker_id <> '' and  " & stringyear
 popuplist1 value, con
End Sub

Private Sub Command24_12_GotFocus()
If PopUpValue1 <> "" Then
  
  txtPCode12.text = popupvalue5
  txtPaper12.text = PopUpValue1 & " , " & PopUpValue3 & " , " & popupvalue4 & " , " & popupvalue5 & " , " & PopUpValue2
  
  ''txtHead11.SetFocus
  ''txtHeadData10.SetFocus
  
End If
  
  PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""
  popupvalue4 = ""
  popupvalue5 = ""
End Sub

Private Sub Command3_Click()
popuplist1 "Select size1 as [Size],size_info as [Size Info] from SizeMaster where size1 <> '' and  " & stringyear & "", con
End Sub

Private Sub Command3_GotFocus()
Label5.Visible = True
If PopUpValue1 <> "" Then
   bsize.text = PopUpValue1
   'txtdes.Text = PopUpValue2
   txtDivide.SetFocus
   bsize.SetFocus
End If
PopUpValue1 = ""
PopUpValue2 = ""
End Sub
Private Sub Command4_Click()
 value = "Select papermaker_name as [Paper Name],Eco,SizeValue1 + ' X ' +SizeValue2 as Size, GSM,papermaker_Id as Code from papermakemaster where papermaker_id <> '' and  " & stringyear
 popuplist1 value, con
End Sub
Private Sub Command4_GotFocus()
If PopUpValue1 <> "" Then
  txtInnerPCode.text = popupvalue5
  txtInnerPaper.text = PopUpValue1 & " , " & PopUpValue3 & " , " & popupvalue4 & " , " & popupvalue5 & " , " & PopUpValue2
  txtHead2.SetFocus
End If
  
  PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""
  popupvalue4 = ""
  popupvalue5 = ""
End Sub

Private Sub Command5_Click()
 value = "Select papermaker_name as [Paper Name],Eco,SizeValue1 + ' X ' +SizeValue2 as Size, GSM,papermaker_Id as Code from papermakemaster where papermaker_id <> '' and  " & stringyear
 popuplist1 value, con
End Sub

Private Sub Command5_GotFocus()

If PopUpValue1 <> "" Then
  txtTextPCode.text = popupvalue5
  txtTextPaper.text = PopUpValue1 & " , " & PopUpValue3 & " , " & popupvalue4 & " , " & popupvalue5 & " , " & PopUpValue2
  txtHead3.SetFocus
End If

  PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""
  popupvalue4 = ""
  popupvalue5 = ""
  
End Sub

Private Sub Command6_Click()
HeadTbl = "Lemination"
frmMasters.Show 1
End Sub

Private Sub Command7_Click()
 value = "Select papermaker_name as [Paper Name],Eco,SizeValue1 + ' X ' +SizeValue2 as Size, GSM,papermaker_Id as Code from papermakemaster where papermaker_id <> '' and  " & stringyear
 popuplist1 value, con
End Sub

Private Sub Command7_GotFocus()
  
 If PopUpValue1 <> "" Then
  txtExamPCode.text = popupvalue5
  txtExamPaper.text = PopUpValue1 & " , " & PopUpValue3 & " , " & popupvalue4 & " , " & popupvalue5 & " , " & PopUpValue2
  
  txtHead4.SetFocus
  
 End If
  
  PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""
  popupvalue4 = ""
  popupvalue5 = ""
  
End Sub

Private Sub Command8_Click()
 value = "Select papermaker_name as [Paper Name],Eco,SizeValue1 + ' X ' +SizeValue2 as Size, GSM,papermaker_Id as Code from papermakemaster where papermaker_id <> '' and  " & stringyear
 popuplist1 value, con
End Sub

Private Sub Command8_GotFocus()
  
  
If PopUpValue1 <> "" Then
  
  txtSuppPCode.text = popupvalue5
  txtSuppPaper.text = PopUpValue1 & " , " & PopUpValue3 & " , " & popupvalue4 & " , " & popupvalue5 & " , " & PopUpValue2
  
  txtHead5.SetFocus
  
End If
  
  PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""
  popupvalue4 = ""
  popupvalue5 = ""
  
End Sub

Private Sub Command9_Click()
 value = "Select papermaker_name as [Paper Name],Eco,SizeValue1 + ' X ' +SizeValue2 as Size, GSM,papermaker_Id as Code from papermakemaster where papermaker_id <> '' and  " & stringyear
 popuplist1 value, con
End Sub

Private Sub Command9_GotFocus()

If PopUpValue1 <> "" Then


  txtTitlePCode.text = popupvalue5
  txtTitlePaper.text = PopUpValue1 & " , " & PopUpValue3 & " , " & popupvalue4 & " , " & popupvalue5 & " , " & PopUpValue2
  
  txtHead6.SetFocus
  
  PopUpValue1 = ""
  PopUpValue2 = ""
  PopUpValue3 = ""
  popupvalue4 = ""
  popupvalue5 = ""
  
End If


End Sub

Private Sub Del_Click()
If txtname.text = "" Then Exit Sub
If MsgBox("Are you sure....", vbOKCancel) = vbOK Then
        con.Execute "Delete from BookMaster where bookno = '" & txtBookCode.text & "'"
        Clearvalue
txtname.SetFocus
End If

End Sub

Private Sub Form_Activate()
bfont = "e"
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then

If Frame1.Visible = True Then
   Frame1.Visible = False
Else
   Unload frmbook
End If

End If

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    sendkeys "{tab}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
   txtBookCode.SetFocus
   Call cmdSearch_Click
End If

End Sub
Sub AddItem()
 
 'cboLemination.AddItem "None"
 
 If RS.State = 1 Then RS.close
 RS.Open "Select * from MasterTbl where category='" & HeadTbl & "' order by name", con, adOpenDynamic, adLockOptimistic
 
 If HeadTbl = "Book_q" Then
    cboQuality.Clear
    While RS.EOF = False
       cboQuality.AddItem RS!Name
       RS.MoveNext
    Wend
 
 ElseIf HeadTbl = "Author" Then
 
    txtWriter.Clear
    While RS.EOF = False
       txtWriter.AddItem RS!Name
       RS.MoveNext
    Wend
 
 ElseIf HeadTbl = "typesetter" Then
 
    txtTypeSetter.Clear
    While RS.EOF = False
       txtTypeSetter.AddItem RS!Name
       RS.MoveNext
    Wend
 
 ElseIf HeadTbl = "negative" Then
 
    txtNegativeby.Clear
    While RS.EOF = False
       txtNegativeby.AddItem RS!Name
       RS.MoveNext
    Wend
 
ElseIf HeadTbl = "class" Then

    cboClass.Clear
    While RS.EOF = False
       cboClass.AddItem RS!Name
       RS.MoveNext
    Wend
 
 
 Else
 
    cboLemination.Clear
    While RS.EOF = False
      cboLemination.AddItem RS!Name
      RS.MoveNext
    Wend

 
 End If
 
 
 '--------------------
' cboLemination.ListIndex = 0
 
 

End Sub
Private Sub Form_Load()

Me.top = 100
Me.Left = 100
Me.Height = 10800
Me.Width = 13905

d1 = 1
frmNo = "boobmaster"

'===========================================

cbofirm.Clear
If RS.State = 1 Then RS.close
RS.Open "select FirmName,Add1,Add2 from FirmMaster order by firmname", con, adOpenStatic, adLockReadOnly
While RS.EOF = False
  cbofirm.AddItem RS(0)
  RS.MoveNext
Wend

cbofirm.ListIndex = 0


'===========================================
cboInnerPrint.Clear
cboTextPrint.Clear
cboExamPrint.Clear
cboSuppPrint.Clear
cboTitlePrint.Clear
cboBinder.Clear



If RS.State = 1 Then RS.close
RS.Open "Select Godwn from Godownmaster where (Binder_Printer='b' or Binder_Printer='pb') and len(Godwn)>=4 order by Godwn", con, adOpenKeyset, adLockReadOnly
While RS.EOF = False
    cboBinder.AddItem RS(0)
    RS.MoveNext
Wend
    
If RS.State = 1 Then RS.close
'RS.Open "Select Godwn from Godownmaster where (Binder_Printer='p' or Binder_Printer='pb') and len(Godwn)>=4 order by Godwn", con, adOpenKeyset, adLockReadOnly
RS.Open "Select Godwn from Godownmaster where len(Godwn)>=4 order by Godwn", con, adOpenKeyset, adLockReadOnly
While RS.EOF = False
    
    cboInnerPrint.AddItem RS(0)
    cboTextPrint.AddItem RS(0)
    cboExamPrint.AddItem RS(0)
    cboSuppPrint.AddItem RS(0)
    cboTitlePrint.AddItem RS(0)
    
    cboPrinter6.AddItem RS(0)
    cboPrinter7.AddItem RS(0)
    cboPrinter8.AddItem RS(0)
    cboPrinter9.AddItem RS(0)
    cboPrinter10.AddItem RS(0)
    
    cboPrinter11.AddItem RS(0)
    cboPrinter12.AddItem RS(0)
    
    
    
    RS.MoveNext
Wend


'============================================

If RS.State = 1 Then RS.close
RS.Open "Select * from MasterTbl where category='binding'", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   cboBinding.AddItem RS(0)
   RS.MoveNext
Wend


txtTrimSize.Clear

If RS.State = 1 Then RS.close
RS.Open "Select * from MasterTbl where category='trimsize'", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   txtTrimSize.AddItem RS(0)
   RS.MoveNext
Wend


 
 
cboLemination.Clear
If RS.State = 1 Then RS.close
RS.Open "Select * from MasterTbl where category='Lemination'", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   cboLemination.AddItem RS(0)
   RS.MoveNext
Wend


If RS.State = 1 Then RS.close
RS.Open "Select * from MasterTbl where category='Author' order by name", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   txtWriter.AddItem RS!Name
   RS.MoveNext
Wend


If RS.State = 1 Then RS.close
RS.Open "Select * from MasterTbl where category='typesetter' order by name", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   txtTypeSetter.AddItem RS!Name
   RS.MoveNext
Wend


If RS.State = 1 Then RS.close
RS.Open "Select * from MasterTbl where category='negative' order by name", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   txtNegativeby.AddItem RS!Name
   RS.MoveNext
Wend


If RS.State = 1 Then RS.close
RS.Open "Select * from MasterTbl where category='class' order by name", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   cboClass.AddItem RS!Name
   cboClass1.AddItem RS(0) & ""
   RS.MoveNext
Wend

If cboClass.ListIndex >= 0 Then
cboClass.ListIndex = 0
End If


'cboLemination.ListIndex = 0
 



If rs_11.State = 1 Then rs_11.close
rs_11.Open "Select bookNo from BookMaster where class='" & cboClass.text & "' and  " & stringyear, con, adOpenKeyset, adLockReadOnly

If rs_11.EOF = False Then
  rs_11.MoveFirst
   txtBookCode.text = rs_11!bookNo & ""
   rs_11.Find "BookNo='" & txtBookCode.text & "'"
   If Not (rs_11.EOF) Then
      lstfocus
   If rs_11.EOF Then
      Exit Sub
   End If
End If

End If


If RS.State = 1 Then RS.close
RS.Open "Select * from MasterTbl where category='bkpart'", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False

  frmbook.txtHead1.AddItem RS!Name
  frmbook.txtHead2.AddItem RS!Name
  frmbook.txtHead3.AddItem RS!Name
  frmbook.txtHead4.AddItem RS!Name
  frmbook.txtHead5.AddItem RS!Name
  frmbook.txtHead6.AddItem RS!Name
  frmbook.txtHead7.AddItem RS!Name
  frmbook.txtHead8.AddItem RS!Name
  frmbook.txtHead9.AddItem RS!Name
  frmbook.txtHead10.AddItem RS!Name
  
  frmbook.txtHead11.AddItem RS!Name
  frmbook.txtHead12.AddItem RS!Name
  
  
   RS.MoveNext
Wend

txtDivide.Clear
txtDivide.AddItem "8"
txtDivide.AddItem "16"
txtDivide.AddItem "6"
txtDivide.AddItem "24"
txtDivide.AddItem "36"

BackColorFrom Me

End Sub
Sub SaveBookDet()

'If rs11.State = 1 Then rs11.close
Set rs11 = New ADODB.Recordset
rs11.Open "select * from bookdetails where bookno='" & txtBookCode.text & "'", con, adOpenDynamic, adLockOptimistic
If rs11.EOF = True Then
   rs11.AddNew
Else
   rs11!bookNo = Trim(txtBookCode.text)
   rs11!fyear = main.session
   rs11!TotalPrinted = Val(txtPrintedCY.text)
   rs11!Specimen = Val(txtSpecimenCY.text)
   rs11!Price = Val(txtPrice.text)
   rs11!returnQty = Val(txtRet.text) & ""
   'rs11!ID = MaxSNo("bookdetails", "id")
   rs11.update
End If

'------------------------------------------------------------------------------
''Dim ses1 As String
''ses1 = ""
''ses1 = Left(main.session, 4) - 1 & "-" & Mid(main.session, 3, 2)
''Set rs11 = New ADODB.Recordset
''rs11.Open "select * from bookdetails where " & stringyear & " and bookno='" & txtBookCode.Text & "'", con, adOpenDynamic, adLockOptimistic
''If rs11.EOF = True Then
''   rs11.AddNew
''End If
''   rs11!bookNo = Trim(txtBookCode.Text)
''   rs11!fyear = ses1
''   rs11!TotalPrinted = Val(txtPrintedLY)
''   rs11!Specimen = Val(txtSpecimenLY)
''   rs11!Price = Val(txtPriceLY)
''   rs11!returnQty = Val(txtRetLY)
''   rs11.update
''
''rs11.close





End Sub
Sub SearchBookDet()

If rs11.State = 1 Then rs11.close
rs11.Open "select * from bookdetails where (" & stringyear & " and bookno='" & txtBookCode.text & "') order by ID desc", con, adOpenKeyset, adLockReadOnly
If rs11.EOF = False Then
   txtSpecimenCY.text = rs11!Specimen & ""
   txtPriceLY.text = rs11!Price & ""
   txtPrintedCY.text = rs11!TotalPrinted & ""
   
   lblPriceCY.Caption = "(CY)"
   lblspecimentCY.Caption = "(CY)"
   lbltotalprintedCY.Caption = "(CY)"
   txtSpecimenCY.text = rs11!Specimen & ""
End If




Dim ses1 As String
ses1 = ""
ses1 = Left(main.session, 4) - 1 & "-" & Mid(main.session, 3, 2)

If rs11.State = 1 Then rs11.close
rs11.Open "select * from bookdetails where " & stringyear & " and bookno='" & txtBookCode.text & "'", con, adOpenDynamic, adLockOptimistic
If rs11.EOF = False Then
   txtPrintedLY = rs11!TotalPrinted
   txtSpecimenLY = rs11!Specimen
   txtPriceLY = rs11!Price
   txtRetLY = rs11!returnQty
End If





End Sub


Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.top = (Me.ScaleHeight - panel.Height) / 2

End Sub






Private Sub picOriginal_Click()
cd.Flags = cdlOFNOverwritePrompt + cdlOFNNoReadOnlyReturn
 cd.DialogTitle = "Select Photo..."
 cd.filename = ""
 cd.Filter = "*.bmp|*.jpg"
 
 cd.ShowOpen
'============
picOriginal.Picture = LoadPicture(cd.filename, 4, 3, 0, 0)
'picOriginal.Picture = LoadPicture(cd.FileName)
picOriginal.Visible = True


End Sub

Private Sub REPORTCD_Click()

DSNNew

cr.Reset
cr.ReportFileName = rptPath & "/booklist.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
If cboClass.text <> "" Then
   cr.ReplaceSelectionFormula "{BookMaster.class}='" & cboClass.text & "'"
End If
cr.WindowShowPrintSetupBtn = True
cr.WindowState = crptMaximized
cr.Action = 1

End Sub
Private Sub save_Click()

On Error GoTo aa10

Dim rs_isbn As New ADODB.Recordset
If rs_isbn.State = 1 Then rs_isbn.close

If txtBookCode = "" Then
   MsgBox "Enter Book Code ..", vbCritical
   txtBookCode.SetFocus
   Exit Sub
End If

If txtname = "" Then
   MsgBox "Enter Book Name ..", vbCritical
   txtname.SetFocus
   Exit Sub
End If

Set ref = New ADODB.Recordset

If ref.State = 1 Then ref.close
ref.Open "Select * from BookMaster where BookNo = '" & txtBookCode.text & "'", con, adOpenDynamic, adLockOptimistic
If ref.RecordCount <= 0 Then

If Len(txtISBN) > 2 Then
    rs_isbn.Open "Select top 2 isbn,bookno from BookMaster where isbn = '" & Trim(txtISBN) & "'", con, adOpenKeyset, adLockReadOnly
    If rs_isbn.EOF = False Then
    If rs_isbn!bookNo <> Trim(txtBookCode.text) Then
       MsgBox "Duplicate ISBN .." & vbCrLf & "This ISBN is already given to " & rs_isbn!bookNo, vbCritical
       txtISBN.SetFocus
       Exit Sub
    End If
    End If
End If

End If


If txtname.text <> "" Then

SaveBookDet

If ref.State = 1 Then ref.close
ref.Open "Select * from BookMaster where BookNo = '" & txtBookCode.text & "'", con, adOpenDynamic, adLockOptimistic
If ref.RecordCount <= 0 Then

ref.AddNew
End If

ref!remarks = txtRem.text
ref!Binding = cboBinding.text
ref!trimsize = Trim(txtTrimSize.text)
ref!isbn = txtISBN.text

ref!firmname = cbofirm.text

ref!titlepath = cd.filename
ref!InternalPrint = Val(txtInternalPrint.text)
ref!gpno = Val(txtgpNo.text)
ref!Price = Val(txtPrice.text)

ref!PriceLY = txtPriceLY.text

ref!Class = cboClass.text
ref!Book = txtname.text
ref!bookNo = txtBookCode.text
ref!book_info = txtdes.text
ref!book_size = bsize.text
If bfont <> "" Then
ref!bookfont = bfont.text
End If

ref!websheet = ftype.text
If bunit.text = "" Then
ref!book_unit = 0
Else
ref!book_unit = bunit.text
End If

If bwastage.text = "" Then
ref!Wastage = 0
Else
ref!Wastage = bwastage.text
End If

'If Val(txtHeadData1.Text) <> 0 Then
   ref!Head1 = txtHead1.text
   ref!HeadData1 = txtHeadData1.text
'End If

'If Val(txtHeadData2.Text) <> 0 Then
   ref!Head2 = txtHead2.text
   ref!HeadData2 = txtHeadData2.text
'End If

'If Val(txtHeadData3.Text) <> 0 Then
   ref!Head3 = txtHead3.text
   ref!HeadData3 = txtHeadData3.text
'End If

''If Val(txtHeadData4.Text) <> 0 Then
   ref!Head4 = txtHead4.text
   ref!HeadData4 = txtHeadData4.text
''End If

''If Val(txtHeadData5.Text) <> 0 Then
   ref!Head5 = txtHead5.text
   ref!HeadData5 = txtHeadData5.text
''End If

'If Val(txtDivide.Text) <> 0 Then

If txtDivide.text <> "" Then
   ref!DivideValue = txtDivide.text
End If

'End If

ref!Writer = txtWriter.text
ref!TypeSetter = txtTypeSetter.text
ref!NegativeBy = txtNegativeby.text

ref!Binder = cboBinder.text
ref!quality = cboQuality
ref!Color = cboColour.text
ref!Lamination = cboLemination.text

ref!Inn_Printer = cboInnerPrint.text
ref!text_Printer = cboTextPrint.text
ref!Exam_Printer = cboExamPrint.text
ref!Supp_Printer = cboSuppPrint.text
ref!Title_Printer = cboTitlePrint.text

ref!Inn_color = cboColour.text
ref!text_color = cboTextColour.text
ref!Exam_color = cboExamColour.text
ref!supp_color = cboSuppColour.text
ref!title_color = cboTitleColour.text


ref!Inn_pcode = txtInnerPCode.text
ref!text_pcode = txtTextPCode.text
ref!Exam_pcode = txtExamPCode.text
ref!supp_pcode = txtSuppPCode.text
ref!title_pcode = txtTitlePCode.text


'new code
ref!pcode1 = txtInnerPCode.text
ref!pcode2 = txtTextPCode.text
ref!pcode3 = txtExamPCode.text
ref!pcode4 = txtSuppPCode.text
ref!pcode5 = txtTitlePCode.text
ref!pcode6 = txtPCode6.text
ref!pcode7 = txtPCode7.text
ref!pcode8 = txtPCode8.text


ref!color1 = cboColour.text
ref!color2 = cboTextColour.text
ref!color3 = cboExamColour.text
ref!color4 = cboSuppColour.text
ref!color5 = cboTitleColour.text
ref!color6 = cboColour6.text
ref!color7 = cboColour7.text
ref!color8 = cboColour8.text





'=================================================


ref!edition = txtEdition.text

'If cboInner.Text <> "" Then
ref!Inn_DBy = Val(cboInner.text)
'End If

ref!text_DBy = IIf(cbotext.text = "", 0, cbotext.text)
ref!Exam_DBy = IIf(cboExam.text = "", 0, cboExam.text)
ref!Supp_DBy = IIf(cbosupp.text = "", 0, cbosupp.text)
ref!Title_DBy = IIf(cbotitle.text = "", 0, cbotitle.text)

If txtInForms.text <> "" Then
ref!Inn_Forms = txtInForms.text
End If

ref!text_Forms = IIf(txtTextForms.text = "", 0, txtTextForms.text)
ref!Exam_Forms = IIf(txtTextExam.text = "", 0, txtTextExam.text)
ref!Supp_Forms = IIf(txtTextSupp.text = "", 0, txtTextSupp.text)
ref!Title_Forms = IIf(txtTextTitle.text = "", 0, txtTextTitle.text)

ref!Inn_Bright = txtInnbright.text
ref!text_Bright = txttextbright.text
ref!Exam_Bright = txtExambright.text
ref!Supp_Bright = txtSuppbright.text
ref!Title_Bright = txtTitlebright.text
'==================================================

'=================================================


ref!txtHead6 = txtHead6.text
ref!txtHeadData6 = IIf(txtHeadData6.text = "", 0, txtHeadData6.text)
ref!cbosupp6 = IIf(cbosupp6.text = "", 0, cbosupp6.text)
ref!txtTextSupp6 = IIf(txtTextSupp6.text = "", 0, txtTextSupp6.text)
ref!cboPrinter6 = cboPrinter6.text
ref!cboColour6 = cboColour6.text
ref!txtPCode6 = txtPCode6.text



ref!txtHead7 = txtHead7.text
ref!txtHeadData7 = IIf(txtHeadData7.text = "", 0, txtHeadData7.text)
ref!cbosupp7 = IIf(cbosupp7.text = "", 0, cbosupp7.text)
ref!txtTextSupp7 = IIf(txtTextSupp7.text = "", 0, txtTextSupp7.text)
ref!cboPrinter7 = cboPrinter7.text
ref!cboColour7 = cboColour7.text
ref!txtPCode7 = txtPCode7.text



ref!txtHead8 = txtHead8.text
ref!txtHeadData8 = IIf(txtHeadData8.text = "", 0, txtHeadData8.text)
ref!cbosupp8 = IIf(cbosupp8.text = "", 0, cbosupp8.text)
ref!txtTextSupp8 = IIf(txtTextSupp8.text = "", 0, txtTextSupp8.text)
ref!cboPrinter8 = cboPrinter8.text
ref!cboColour8 = cboColour8.text
ref!txtPCode8 = txtPCode8.text



ref!txtHead9 = txtHead9.text
ref!txtHeadData9 = IIf(txtHeadData9.text = "", 0, txtHeadData9.text)
ref!cbosupp9 = IIf(cbosupp9.text = "", 0, cbosupp9.text)
ref!txtTextSupp9 = IIf(txtTextSupp9.text = "", 0, txtTextSupp9.text)
ref!cboPrinter9 = cboPrinter9.text
ref!Color9 = cboColour9.text
ref!txtPCode9 = txtPCode9.text


ref!txtHead10 = txtHead10.text
ref!txtHeadData10 = IIf(txtHeadData10.text = "", 0, txtHeadData10.text)
ref!cbosupp10 = IIf(cbosupp10.text = "", 0, cbosupp10.text)
ref!txtTextSupp10 = IIf(txtTextSupp10.text = "", 0, txtTextSupp10.text)
ref!cboPrinter10 = cboPrinter10.text
ref!Color10 = cboColour10.text
ref!txtPCode10 = txtPCode10.text

ref!txtHead11 = txtHead11.text
ref!txtHeadData11 = IIf(txtHeadData11.text = "", 0, txtHeadData11.text)
ref!cbosupp11 = IIf(cbosupp11.text = "", 0, cbosupp11.text)
ref!txtTextSupp11 = IIf(txtTextSupp11.text = "", 0, txtTextSupp11.text)
ref!cboPrinter11 = cboPrinter11.text
ref!Color11 = cboColour11.text
ref!txtPCode11 = txtPCode11.text


ref!txtHead12 = txtHead12.text
ref!txtHeadData12 = IIf(txtHeadData12.text = "", 0, txtHeadData12.text)
ref!cbosupp12 = IIf(cbosupp12.text = "", 0, cbosupp12.text)
ref!txtTextSupp12 = IIf(txtTextSupp12.text = "", 0, txtTextSupp12.text)
ref!cboPrinter12 = cboPrinter12.text
ref!Color12 = cboColour12.text
ref!txtPCode12 = txtPCode12.text




'==================================================
ref!setupid = setupid
ref!fyear = session
    
ref!uname = UserName

ref.update
MsgBox "Record saved..", vbInformation

Clearvalue
txtBookCode.SetFocus
'txtname.SetFocus
End If

Exit Sub
aa10:

Set ref = New ADODB.Recordset
Set rs11 = New ADODB.Recordset

MsgBox "" & err.DESCRIPTION

End Sub

Private Sub Text1_GotFocus()
Label5.Visible = True
If PopUpValue1 <> "" Then
   Text1.text = PopUpValue1
   'txtdes.Text = PopUpValue2
End If
PopUpValue1 = ""
PopUpValue2 = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
      popuplist1 "Select customer_id, Customer_name from CustomerMaster where " & stringyear & " and customer_id <> ''", con
      End If
End Sub

Private Sub TtxtHeadData12_Change()
 calForm
 totalValue
End Sub

Private Sub txtBookCode_GotFocus()

If PopUpValue1 = "" Then Exit Sub
   lstfocus
   

   PopUpValue1 = ""
   PopUpValue2 = ""

End Sub

Private Sub txtBookCode_KeyDown(KeyCode As Integer, Shift As Integer)

 Dim rr As New ADODB.Recordset
 
 rr.Open "select top 1 * from BookMaster where BookNo='" & txtBookCode.text & "' and  " & stringyear, con, adOpenKeyset, adLockReadOnly
 If rr.EOF = True Then
    
    If KeyCode = 114 Then
        txtname.Font = hindi
        bfont = "h"
    End If
    
    If KeyCode = 115 Then
        txtname.Font = english
        bfont = "e"
    End If
    
End If



If KeyCode = 113 Then
 'If cboClass <> "" Then
 
    searchType = "party"
    popuplist_client "Select BookNo as [Book Code],book as [Book Name] from BookMaster order by BookNo", con
    
 'End If
End If
   

End Sub

Private Sub txtBookCode_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

    Set ref = New ADODB.Recordset
    If ref.State = 1 Then ref.close
    ref.Open "Select * from bookMaster where bookno = '" + txtBookCode.text + "' and  " & stringyear, con, adOpenKeyset, adLockReadOnly
    If ref.EOF = False Then
       lstfocus
       
    End If



End If

End Sub

Private Sub txtdes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 114 Then
'txtname.Font = hindi
txtdes.Font = hindi
bfont = "h"
End If
If KeyCode = 115 Then
'txtname.Font = english
txtdes.Font = english
bfont = "e"
End If
End Sub

Private Sub txtdes_KeyPress(KeyAscii As Integer)
   b = val_int(txtdes, KeyAscii)
   If b = False Then KeyAscii = 0

End Sub

Private Sub txtDivide_Click()
On Error Resume Next
 cboInner.text = Val(txtDivide.text)
 cbotext.text = Val(txtDivide.text)
 cboExam.text = Val(txtDivide.text)
 cbosupp.text = Val(txtDivide.text)
 'txtHeadData5.Text = Val(txtDivide.Text)
 
 totalValue

End Sub

Private Sub txtDivide_GotFocus()
'   txtDivide.SelLength = 10
End Sub

Private Sub txtDivide_KeyPress(KeyAscii As Integer)
'   b = val_int(txtDivide, KeyAscii)
'   If b = False Then KeyAscii = 0
End Sub

Private Sub txtHeadData1_Change()
  '' calForm
  '' totalValue
End Sub
Private Sub txtHeadData1_GotFocus()
  txtHeadData1.SelLength = 10
End Sub

Private Sub txtHeadData1_KeyPress(KeyAscii As Integer)
   b = val_int(txtHeadData1, KeyAscii)
   If b = False Then KeyAscii = 0
End Sub
Sub totalValue()
    
    
    If Val(txtHeadData1.text) = 0 Then
       txtHeadData1.text = 0
    ElseIf Val(txtHeadData2.text) = 0 Then
       txtHeadData2.text = 0
    ElseIf Val(txtHeadData3.text) = 0 Then
       txtHeadData3.text = 0
    ElseIf Val(txtHeadData4.text) = 0 Then
       txtHeadData4.text = 0
    ElseIf Val(txtHeadData5.text) = 0 Then
       txtHeadData5.text = 0
    ElseIf Val(txtHeadData6.text) = 0 Then
       txtHeadData6.text = 0
    ElseIf Val(txtHeadData7.text) = 0 Then
       txtHeadData7.text = 0
    ElseIf Val(txtHeadData8.text) = 0 Then
       txtHeadData8.text = 0
    ElseIf Val(txtHeadData9.text) = 0 Then
       txtHeadData9.text = 0
    ElseIf Val(txtHeadData10.text) = 0 Then
       txtHeadData10.text = 0
    ElseIf Val(txtHeadData11.text) = 0 Then
       txtHeadData11.text = 0
    ElseIf Val(txtHeadData2.text) = 0 Then
       txtHeadData12.text = 0
 
    End If
    
      
    
    totalSum = (Val(txtHeadData1.text) + Val(txtHeadData2.text) + Val(txtHeadData3.text) + Val(txtHeadData4.text) + Val(txtHeadData6.text) + Val(txtHeadData7.text) + Val(txtHeadData8.text) + Val(txtHeadData9.text) + Val(txtHeadData10.text) + Val(txtHeadData11.text) + Val(txtHeadData11.text))
    
    txtTotal.text = totalSum
    
 If Val(txtDivide.text) > 0 Then
   bunit.text = Round((Val(txtTotal.text) / Val(txtDivide.text)), 2)
   tmpbunit = Val(bunit.text) - Int(Val(bunit.text))
   If tmpbunit > 0 And tmpbunit <= 0.25 Then
        bunit = Int(Val(bunit.text)) + 0.25
   Else
        If tmpbunit >= 0.26 And tmpbunit <= 0.5 Then
            bunit = Int(Val(bunit.text)) + 0.5
        Else
            If tmpbunit >= 0.51 And tmpbunit <= 0.75 Then
                bunit = Int(Val(bunit.text)) + 0.75
                Else
                If tmpbunit >= 0.76 And tmpbunit <= 0.999 Then
                    bunit = Int(Val(bunit.text)) + 1
                End If
            End If
        End If
   End If
   End If
   
   
End Sub
Private Sub txtHeadData1_LostFocus()
 calForm
 totalValue
End Sub

Private Sub txtHeadData10_Change()
 calForm
 totalValue
End Sub

Private Sub txtHeadData11_Change()
 calForm
 totalValue
End Sub

Private Sub txtHeadData2_Change()
'     calForm
 '  totalValue

End Sub
Private Sub txtHeadData2_GotFocus()
  txtHeadData2.SelLength = 10
End Sub
Private Sub txtHeadData2_KeyPress(KeyAscii As Integer)
   b = val_int(txtHeadData2, KeyAscii)
   If b = False Then KeyAscii = 0
End Sub


Private Sub txtHeadData2_LostFocus()
   calForm
   totalValue

End Sub

Private Sub txtHeadData3_Change()
  ' calForm
  ' totalValue

End Sub
Private Sub txtHeadData3_GotFocus()
   txtHeadData3.SelLength = 10
End Sub

Private Sub txtHeadData3_KeyPress(KeyAscii As Integer)
   b = val_int(txtHeadData3, KeyAscii)
   If b = False Then KeyAscii = 0
End Sub

Private Sub txtHeadData3_LostFocus()
   calForm
   totalValue

End Sub

Private Sub txtHeadData4_Change()
'   calForm
'   totalValue

End Sub
Private Sub txtHeadData4_GotFocus()
   txtHeadData4.SelLength = 10
End Sub

Private Sub txtHeadData4_KeyPress(KeyAscii As Integer)
   b = val_int(txtHeadData4, KeyAscii)
   If b = False Then KeyAscii = 0
End Sub

Private Sub txtHeadData4_LostFocus()
   calForm
   totalValue

End Sub

Private Sub txtHeadData5_Change()
  'totalValue
End Sub

Private Sub txtHeadData5_GotFocus()
 txtHeadData5.SelLength = 10
End Sub

Private Sub txtHeadData5_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then txtTextTitle.SetFocus
End Sub

Private Sub txtHeadData5_LostFocus()
   calForm
   totalValue

End Sub

Private Sub txtHeadData6_LostFocus()
  calForm
  totalValue
End Sub



Private Sub txtHeadData7_LostFocus()
   calForm
   totalValue

End Sub



Private Sub txtHeadData8_LostFocus()
   calForm
   totalValue

End Sub

Private Sub txtHeadData9_Change()
 calForm
 totalValue

End Sub

Private Sub txtname_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 114 Then
    txtname.Font = hindi
    bfont = "h"
End If

If KeyCode = 115 Then
    txtname.Font = english
    bfont = "e"
End If

End Sub
Sub lstfocus()

On Error Resume Next



If PopUpValue1 <> "" Then
   txtBookCode.text = PopUpValue1
End If

Label5.Visible = True



Set ref = New ADODB.Recordset
If ref.State = 1 Then ref.close
ref.Open "Select * from bookMaster where bookno = '" + txtBookCode.text + "' order by BookNo", con, adOpenKeyset, adLockReadOnly
If ref.RecordCount > 0 Then

Clearvalue

SearchBookDet


txtRem.text = ref!remarks & ""
cboBinding.text = ref!Binding & ""

txtTrimSize.text = ref!trimsize & ""

txtBookCode.text = ref!bookNo
'txtSpecimenCY = ref!Sepimen_CY & ""
txtRet = ref!return_cy & ""

txtPriceLY.text = ref!PriceLY & ""

'txtInternalPrint.Text = ref!InternalPrint

txtgpNo.text = ref!gpno
txtPrice.text = ref!Price
cboClass.text = ref!Class



If ref!bookfont = "e" Then
   txtname.FontName = english
   txtname = ref!Book
Else
   txtname.FontName = hindi
   txtname = ref!Book
End If
   
'cboClass.Text = ref!Class & ""
   
txtPrice.text = ref!Price & ""

txtgpNo.text = ref!gpno & ""

txtInternalPrint.text = ref!InternalPrint & ""

txtISBN.text = ref!isbn & ""

cbofirm.text = ref!firmname
   
txtRet.text = ref!return_cy & ""

If Not IsNull(ref!titlepath) Then
picOriginal.Picture = LoadPicture(ref!titlepath)

End If
   
txtdes = IIf(IsNull(ref!book_info), "", ref!book_info)
bsize = IIf(IsNull(ref!book_size), "", ref!book_size)
bunit = IIf(IsNull(ref!book_unit), "", ref!book_unit)
formrate = IIf(IsNull(ref!atrate), "", ref!atrate)
platerate = IIf(IsNull(ref!patrate), "", ref!patrate)
bwastage = IIf(IsNull(ref!Wastage), "", ref!Wastage)
ftype.text = IIf(IsNull(ref!websheet), "", ref!websheet)
'Text1.Text = IIf(IsNull(ref!party_id), "", ref!party_id)
bfont.text = ref!bookfont

'txtHead1.Text = IIf(IsNull(ref!Head1), "", ref!Head1)
'txtHead2.Text = IIf(IsNull(ref!Head2), "", ref!Head2)
'txtHead3.Text = IIf(IsNull(ref!Head3), "", ref!Head3)
'txtHead4.Text = IIf(IsNull(ref!Head4), "", ref!Head4)
'txtHead5.Text = IIf(IsNull(ref!Head5), "", ref!Head5)

txtHeadData1.text = IIf(IsNull(ref!HeadData1), "", ref!HeadData1)
txtHeadData2.text = IIf(IsNull(ref!HeadData2), "", ref!HeadData2)
txtHeadData3.text = IIf(IsNull(ref!HeadData3), "", ref!HeadData3)
txtHeadData4.text = IIf(IsNull(ref!HeadData4), "", ref!HeadData4)
txtHeadData5.text = IIf(IsNull(ref!HeadData5), "", ref!HeadData5)




'----------------------------------------------------------------------
txtHead1.text = ref!Head1 & ""
txtHead2.text = ref!Head2 & ""
txtHead3.text = ref!Head3 & ""
txtHead4.text = ref!Head4 & ""
txtHead5.text = ref!Head5 & ""
txtHead6.text = ref!txtHead6 & ""
txtHead7.text = ref!txtHead7 & ""
txtHead8.text = ref!txtHead8 & ""
txtHead9.text = ref!txtHead9 & ""
txtHead10.text = ref!txtHead10 & ""

txtHeadData6.text = ref!txtHeadData6 & ""

If ref!cbosupp6 <> "" Then
  cbosupp6.text = ref!cbosupp6
End If

If ref!txtTextSupp6 <> "" Then
   txtTextSupp6.text = ref!txtTextSupp6
End If

If ref!cboPrinter6 <> "" Then
   cboPrinter6.text = ref!cboPrinter6
End If

If ref!cboColour6 <> "" Then
cboColour6.text = ref!cboColour6
End If

If ref!txtPCode6 <> "" Then
   txtPCode6.text = ref!txtPCode6
End If

If ref!txtHead7 <> "" Then
   txtHead7.text = ref!txtHead7
End If

If ref!txtHeadData7 <> "" Then
   txtHeadData7.text = ref!txtHeadData7
End If

If ref!cbosupp7 <> "" Then
cbosupp7.text = ref!cbosupp7
End If



If ref!txtTextSupp7 <> "" Then
  txtTextSupp7.text = ref!txtTextSupp7
End If


If ref!txtTextSupp8 <> "" Then
  txtTextSupp8.text = ref!txtTextSupp8
End If


If ref!txtTextSupp9 <> "" Then
  txtTextSupp9.text = ref!txtTextSupp9
End If

If ref!txtTextSupp10 <> "" Then
  txtTextSupp10.text = ref!txtTextSupp10
End If


If ref!txtTextSupp11 <> "" Then
  txtTextSupp11.text = ref!txtTextSupp11
End If

If ref!txtTextSupp12 <> "" Then
  txtTextSupp12.text = ref!txtTextSupp12
End If


If ref!cboPrinter7 <> "" Then
   cboPrinter7.text = ref!cboPrinter7
End If


If ref!cboPrinter8 <> "" Then
   cboPrinter8.text = ref!cboPrinter8
End If


If ref!cboPrinter9 <> "" Then
   cboPrinter9.text = ref!cboPrinter9
End If


If ref!cboPrinter10 <> "" Then
   cboPrinter10.text = ref!cboPrinter10
End If

If ref!cboPrinter11 <> "" Then
   cboPrinter11.text = ref!cboPrinter11
End If


If ref!cboPrinter12 <> "" Then
   cboPrinter12.text = ref!cboPrinter12
End If

If ref!cboColour7 <> "" Then
cboColour7.text = ref!cboColour7
End If

If ref!cboColour8 <> "" Then
cboColour8.text = ref!cboColour8
End If


If ref!Color9 <> "" Then
cboColour9.text = ref!Color9
End If


If ref!Color10 <> "" Then
cboColour10.text = ref!Color10
End If

If ref!Color11 <> "" Then
cboColour11.text = ref!Color11
End If

If ref!Color12 <> "" Then
cboColour12.text = ref!Color12
End If



If ref!txtPCode7 <> "" Then
txtPCode7.text = ref!txtPCode7
End If


If ref!txtPCode8 <> "" Then
txtPCode8.text = ref!txtPCode8
End If



If ref!txtPCode9 <> "" Then
txtPCode9.text = ref!txtPCode9
End If

If ref!txtPCode10 <> "" Then
txtPCode10.text = ref!txtPCode10
End If

If ref!txtPCode11 <> "" Then
txtPCode11.text = ref!txtPCode11
End If

If ref!txtPCode12 <> "" Then
txtPCode12.text = ref!txtPCode12
End If




If ref!txtHead8 <> "" Then
txtHead8.text = ref!txtHead8
End If


If ref!txtHead9 <> "" Then
txtHead9.text = ref!txtHead9
End If


If ref!txtHead10 <> "" Then
txtHead10.text = ref!txtHead10
End If

If ref!txtHead11 <> "" Then
txtHead11.text = ref!txtHead11
End If

If ref!txtHead12 <> "" Then
txtHead12.text = ref!txtHead12
End If


If ref!txtHeadData8 <> "" Then
txtHeadData8.text = ref!txtHeadData8
End If


If ref!txtHeadData9 <> "" Then
txtHeadData9.text = ref!txtHeadData9
End If


If ref!txtHeadData10 <> "" Then
txtHeadData10.text = ref!txtHeadData10
End If


If ref!txtHeadData11 <> "" Then
txtHeadData11.text = ref!txtHeadData11
End If


If ref!txtHeadData12 <> "" Then
txtHeadData12.text = ref!txtHeadData12
End If


If ref!cbosupp8 <> "" Then
cbosupp8.text = ref!cbosupp8
End If

If ref!cbosupp9 <> "" Then
cbosupp9.text = ref!cbosupp9
End If

If ref!cbosupp10 <> "" Then
cbosupp10.text = ref!cbosupp10
End If



If ref!cbosupp11 <> "" Then
 cbosupp11.text = ref!cbosupp11
End If


If ref!cbosupp12 <> "" Then
 cbosupp12.text = ref!cbosupp12
End If



If ref!txtTextSupp8 <> "" Then
   txtTextSupp8.text = ref!txtTextSupp8
End If


If ref!txtTextSupp9 <> "" Then
   txtTextSupp9.text = ref!txtTextSupp9
End If


If ref!txtTextSupp10 <> "" Then
   txtTextSupp10.text = ref!txtTextSupp10
End If

If ref!txtTextSupp11 <> "" Then
   txtTextSupp11.text = ref!txtTextSupp11
End If

If ref!txtTextSupp12 <> "" Then
   txtTextSupp12.text = ref!txtTextSupp12
End If



If ref!cboPrinter8 <> "" Then
   cboPrinter8.text = ref!cboPrinter8
End If

If ref!cboPrinter9 <> "" Then
   cboPrinter9.text = ref!cboPrinter9
End If


If ref!cboPrinter10 <> "" Then
   cboPrinter10.text = ref!cboPrinter10
End If

If ref!cboPrinter11 <> "" Then
   cboPrinter11.text = ref!cboPrinter11
End If

If ref!cboPrinter12 <> "" Then
   cboPrinter12.text = ref!cboPrinter12
End If




If ref!cboColour8 <> "" Then
   cboColour8.text = ref!cboColour8
End If


If ref!Color9 <> "" Then
   cboColour9.text = ref!Color9
End If


If ref!Color10 <> "" Then
   cboColour10.text = ref!Color10
End If

If ref!Color11 <> "" Then
   cboColour11.text = ref!Color11
End If

If ref!Color12 <> "" Then
   cboColour12.text = ref!Color12
End If



If ref!txtPCode8 <> "" Then
   txtPCode8.text = ref!txtPCode8
End If

If ref!txtPCode9 <> "" Then
   txtPCode9.text = ref!txtPCode9
End If

If ref!txtPCode10 <> "" Then
   txtPCode10.text = ref!txtPCode10
End If

If ref!txtPCode11 <> "" Then
   txtPCode11.text = ref!txtPCode11
End If

If ref!txtPCode12 <> "" Then
   txtPCode12.text = ref!txtPCode12
End If




'----------------------------------------------------------------------

txtDivide.text = IIf(IsNull(ref!DivideValue), "", ref!DivideValue)
txtWriter.text = IIf(IsNull(ref!Writer), "", ref!Writer)
txtTypeSetter.text = IIf(IsNull(ref!TypeSetter), "", ref!TypeSetter)
txtNegativeby.text = IIf(IsNull(ref!NegativeBy), "", ref!NegativeBy)



'txtBrand.Text = ref!Brand & ""
cboQuality = ref!quality & ""
If ref!Color <> "" Then
cboColour.text = ref!Color
End If

cboBinder.text = ref!Binder & ""

cboInnerPrint.text = ref!Inn_Printer & ""

cboTextPrint.text = ref!text_Printer & ""
cboExamPrint.text = ref!Exam_Printer & ""
cboSuppPrint.text = ref!Supp_Printer & ""
cboTitlePrint.text = ref!Title_Printer & ""

cboColour.text = ref!Inn_color & ""
cboTextColour.text = ref!text_color & ""
cboExamColour.text = ref!Exam_color & ""
cboSuppColour.text = ref!supp_color & ""
cboTitleColour.text = ref!title_color & ""


txtInnerPCode.text = ref!Inn_pcode & ""
txtTextPCode.text = ref!text_pcode & ""
txtExamPCode.text = ref!Exam_pcode & ""
txtSuppPCode.text = ref!supp_pcode & ""
txtTitlePCode.text = ref!title_pcode & ""

txtEdition.text = ref!edition & ""

'====================================================

cboInner.text = ref!Inn_DBy & ""
cbotext.text = ref!text_DBy & ""
cboExam.text = ref!Exam_DBy & ""
cbosupp.text = ref!Supp_DBy & ""
cbotitle.text = ref!Title_DBy & ""

'If InStr(txtHead1.Text, "Tex") > 0 Then
'   txtTextForms.Text = ref!Inn_Forms & ""
'ElseIf InStr(txtHead2.Text, "Tit") > 0 Then
'   txtInForms.Text = ref!text_Forms & ""
'End If

txtInForms.text = ref!Inn_Forms & ""
txtTextForms.text = ref!text_Forms & ""

txtTextExam.text = ref!Exam_Forms & ""
txtTextSupp.text = ref!Supp_Forms & ""
txtTextTitle.text = ref!Title_Forms & ""


txtInnbright.text = ref!Inn_Bright & ""
txttextbright.text = ref!text_Bright & ""
txtExambright.text = ref!Exam_Bright & ""
txtSuppbright.text = ref!Supp_Bright & ""
txtTitlebright.text = ref!Title_Bright & ""

If ref!Lamination <> "" Then
cboLemination.text = ref!Lamination
End If






If RS.State = 1 Then RS.close
st1 = "Select papermaker_name,Eco,GSM,Size,papermaker_Id from papermakemaster where  " & stringyear
RS.Open st1, con, adOpenDynamic, adLockReadOnly, adCmdText
    If txtInnerPCode <> "" Then
     
     RS.MoveFirst
     RS.Find "papermaker_Id='" + txtInnerPCode + "'"
     If RS.EOF = False Then
        txtInnerPaper.text = RS(0) & " , " & RS(1) & " , " & RS(2) & " , " & RS(3)
     End If
    
    End If

'------------------------------------------
    If txtTextPCode <> "" Then
     
     RS.MoveFirst
     RS.Find "papermaker_Id='" & txtTextPCode & "'"
     If RS.EOF = False Then
        txtTextPaper.text = RS(0) & " , " & RS(1) & " , " & RS(2) & " , " & RS(3)
     End If
    
    End If

'------------------------------------------
    If txtExamPCode <> "" Then
     
     RS.MoveFirst
     RS.Find "papermaker_Id='" & txtExamPCode & "'"
     If RS.EOF = False Then
        txtExamPaper.text = RS(0) & " , " & RS(1) & " , " & RS(2) & " , " & RS(3)
     End If
    
    End If


    If txtSuppPCode <> "" Then
     
     RS.MoveFirst
     RS.Find "papermaker_Id='" & txtSuppPCode & "'"
     If RS.EOF = False Then
        txtSuppPaper.text = RS(0) & " , " & RS(1) & " , " & RS(2) & " , " & RS(3)
     End If
    
    End If


    If txtTitlePCode <> "" Then
     
     RS.MoveFirst
     RS.Find "papermaker_Id='" & txtTitlePCode & "'"
     If RS.EOF = False Then
        txtTitlePaper.text = RS(0) & " , " & RS(1) & " , " & RS(2) & " , " & RS(3)
     End If
    
    End If


'================================================

If txtPCode6.text <> "" Then
 
    RS.MoveFirst
    RS.Find "papermaker_Id='" & txtPCode6.text & "'"
    If RS.EOF = False Then
       txtPaper6.text = RS(0) & " , " & RS(1) & " , " & RS(2) & " , " & RS(3)
    End If

End If



If txtPCode7.text <> "" Then
 
    RS.MoveFirst
    RS.Find "papermaker_Id='" & txtPCode7.text & "'"
    If RS.EOF = False Then
       txtPaper7.text = RS(0) & " , " & RS(1) & " , " & RS(2) & " , " & RS(3)
    End If

End If

If txtPCode8.text <> "" Then
 
    RS.MoveFirst
    RS.Find "papermaker_Id='" & txtPCode8.text & "'"
    If RS.EOF = False Then
       txtPaper8.text = RS(0) & " , " & RS(1) & " , " & RS(2) & " , " & RS(3)
    End If

End If

If txtPCode9.text <> "" Then
 
    RS.MoveFirst
    RS.Find "papermaker_Id='" & txtPCode9.text & "'"
    If RS.EOF = False Then
       txtPaper9.text = RS(0) & " , " & RS(1) & " , " & RS(2) & " , " & RS(3)
    End If

End If





If txtPCode10.text <> "" Then
 
    RS.MoveFirst
    RS.Find "papermaker_Id='" & txtPCode10.text & "'"
    If RS.EOF = False Then
       txtPaper10.text = RS(0) & " , " & RS(1) & " , " & RS(2) & " , " & RS(3)
    End If

End If


If txtPCode11.text <> "" Then
 
    RS.MoveFirst
    RS.Find "papermaker_Id='" & txtPCode11.text & "'"
    If RS.EOF = False Then
       txtPaper11.text = RS(0) & " , " & RS(1) & " , " & RS(2) & " , " & RS(3)
    End If

End If


If txtPCode12.text <> "" Then
 
    RS.MoveFirst
    RS.Find "papermaker_Id='" & txtPCode12.text & "'"
    If RS.EOF = False Then
       txtPaper12.text = RS(0) & " , " & RS(1) & " , " & RS(2) & " , " & RS(3)
    End If

End If




End If

End Sub
Sub SearchBook()
On Error Resume Next
txtBookCode.text = RS!bookNo
If RS!bookfont = "e" Then
   txtname.FontName = "english"
   txtname = RS!Book
Else
   txtname.FontName = "Kundli"
   txtname = RS!Book
End If
   
txtdes = IIf(IsNull(RS!book_info), "", RS!book_info)
bsize = IIf(IsNull(RS!book_size), "", RS!book_size)
bunit = IIf(IsNull(RS!book_unit), "", RS!book_unit)
formrate = IIf(IsNull(RS!atrate), "", RS!atrate)
platerate = IIf(IsNull(RS!patrate), "", RS!patrate)
bwastage = IIf(IsNull(RS!Wastage), "", RS!Wastage)
ftype.text = IIf(IsNull(RS!websheet), "", RS!websheet)
Text1.text = IIf(IsNull(RS!party_id), "", RS!party_id)
bfont.text = RS!bookfont


txtHead5.text = IIf(IsNull(RS!Head5), "", RS!Head5)
txtHeadData1.text = IIf(IsNull(RS!HeadData1), "", RS!HeadData1)
txtHeadData2.text = IIf(IsNull(RS!HeadData2), "", RS!HeadData2)
txtHeadData3.text = IIf(IsNull(RS!HeadData3), "", RS!HeadData3)
txtHeadData4.text = IIf(IsNull(RS!HeadData4), "", RS!HeadData4)
txtHeadData5.text = IIf(IsNull(RS!HeadData4), "", RS!HeadData5)

txtDivide.text = IIf(IsNull(RS!DivideValue), "", RS!DivideValue)

txtWriter.text = IIf(IsNull(RS!Writer), "", RS!Writer)
txtTypeSetter.text = IIf(IsNull(RS!TypeSetter), "", RS!TypeSetter)
txtNegativeby.text = IIf(IsNull(RS!NegativeBy), "", RS!NegativeBy)

txtBrand.text = RS!Brand & ""
cboQuality.text = RS!quality & ""
cboColour.text = RS!Color & ""
cboLemination.text = RS!Lemination & ""

End Sub
Private Sub txtName_LostFocus()
Label5.Visible = False
End Sub
Sub Clearvalue()

Dim gpName, Bookcode As String
Dim o As Object

'Picture1.Picture = LoadPicture("")
Set picOriginal.Picture = Nothing

txtRem.text = ""
txtTrimSige = ""

txtISBN = ""
gpName = ""
Bookcode = ""

'txtBookCode = ""

On Error Resume Next

gpName = cboClass.text
Bookcode = txtBookCode

For Each o In Me

If TypeOf o Is ComboBox Then
  o.ListIndex = -1
  o.text = ""
End If

Next


cboLemination.ListIndex = -1
cbofirm.ListIndex = 0

txtBookCode = ""
'------------------------------
txtRet = ""

txtTextSupp6 = ""
txtTextSupp7 = ""
txtTextSupp8 = ""
txtTextSupp9 = ""
txtTextSupp10 = ""
txtTextSupp11 = ""
txtTextSupp12 = ""

txtPaper6 = ""
txtPaper7 = ""
txtPaper8 = ""
txtPaper9 = ""
txtPaper10 = ""
txtPaper11 = ""
txtPaper12 = ""

'-------------------------------
txtHead6 = ""
txtHead7 = ""
txtHead8 = ""
txtHead9 = ""
txtHead10 = ""
txtHead11 = ""
txtHead12 = ""

txtHeadData6 = ""
txtHeadData7 = ""
txtHeadData8 = ""
txtHeadData9 = ""
txtHeadData10 = ""
txtHeadData11 = ""
txtHeadData12 = ""



txtPCode6 = ""
txtPCode7 = ""
txtPCode8 = ""
txtPCode9 = ""
txtPCode10 = ""
txtPCode11 = ""
txtPCode12 = ""
'-----------------------------

txtPriceLY = ""

txtPrintedCY = ""
txtPrintedLY = ""
txtInternalPrint = ""


txtSpecimenCY = ""
txtSpecimenLY = ""

txtPrice.text = ""

txtTextForms.text = ""
txtInForms.text = ""
txtTextExam.text = ""
txtTextSupp.text = ""
txtTextTitle.text = ""


txtInnbright.text = ""
txttextbright.text = ""
txtSuppbright.text = ""
txtExambright.text = ""
txtTitlebright.text = ""



'--------------------------------
txtInnerPCode.text = ""
txtTextPCode.text = ""
txtSuppPCode.text = ""
txtExamPCode.text = ""
txtTitlePCode.text = ""

txtInnerPaper.text = ""
txtTextPaper.text = ""
txtSuppPaper.text = ""
txtExamPaper.text = ""
txtTitlePaper.text = ""



txtEdition.text = ""
cboLemination.ListIndex = -1
txtBrand = ""
'cboQuality.ListIndex = -1
cboColour.ListIndex = -1
cboLemination.ListIndex = -1


'============================================
cboColour.ListIndex = -1
cboTextColour.ListIndex = -1
cboExamColour.ListIndex = -1
cboSuppColour.ListIndex = -1
cboTitleColour.ListIndex = -1



cboInnerPrint.ListIndex = -1
cboTextPrint.ListIndex = -1
cboExamPrint.ListIndex = -1
cboSuppPrint.ListIndex = -1
cboTitlePrint.ListIndex = -1
'============================================

txtname.text = ""
txtdes.text = ""
bsize.text = ""
bunit.text = ""
formrate = ""
platerate = ""
bwastage = ""
ftype.ListIndex = -1
'txtBookCode.Text = " "
txtHeadData1.text = ""
txtHeadData2.text = ""
txtHeadData3.text = ""
txtHeadData4.text = ""
txtHeadData5.text = ""
txtDivide.ListIndex = -1

txtWriter.ListIndex = -1
txtTypeSetter.ListIndex = -1
txtNegativeby.ListIndex = -1
bunit.text = ""

cboClass = gpName
txtBookCode = Bookcode

'searchBookGp


End Sub
Private Sub txtPrice_KeyPress(KeyAscii As Integer)
   
   b = val_int(txtPrice, KeyAscii)
   If b = False Then KeyAscii = 0

End Sub
Private Sub VS_bk_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If VS_bk.Col = 2 Then
       sendkeys "{down}"
    End If
End If
End Sub

Private Sub vs_Click()
txtGpno1.text = vs.TextMatrix(vs.RowSel, 4)
End Sub
