VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmBookMaster 
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   14556
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   14556
   WindowState     =   2  'Maximized
   Begin VB.Frame panel 
      Caption         =   "Book Master"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   9600
      Left            =   75
      TabIndex        =   15
      Top             =   30
      Width           =   14370
      Begin VB.TextBox txtbagInbox 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6165
         MaxLength       =   6
         TabIndex        =   53
         Top             =   3312
         Width           =   675
      End
      Begin VB.TextBox txtClass 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2940
         MaxLength       =   6
         TabIndex        =   11
         Top             =   2925
         Width           =   675
      End
      Begin VB.ComboBox cboGName_sub 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4785
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   2145
         Width           =   2055
      End
      Begin VB.ComboBox cboGcode_sub 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4785
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1740
         Width           =   2055
      End
      Begin VB.TextBox txtupto10 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6165
         MaxLength       =   6
         TabIndex        =   12
         Top             =   2940
         Width           =   675
      End
      Begin VB.TextBox txthsncode 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1425
         MaxLength       =   12
         TabIndex        =   10
         Top             =   2940
         Width           =   1035
      End
      Begin VB.CheckBox Check1_editkit 
         Caption         =   "Edit Kit Qty"
         Height          =   315
         Left            =   12075
         TabIndex        =   38
         Top             =   3975
         Width           =   1095
      End
      Begin VB.Frame frmEditBookKit 
         Height          =   3165
         Left            =   8205
         TabIndex        =   36
         Top             =   525
         Visible         =   0   'False
         Width           =   4275
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Ex&it"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   2520
            Width           =   930
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&OK"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1680
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   2520
            Width           =   930
         End
         Begin VSFlex7Ctl.VSFlexGrid vsKit 
            Height          =   1785
            Left            =   45
            TabIndex        =   37
            Top             =   660
            Width           =   4200
            _cx             =   7408
            _cy             =   3149
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
            BackColor       =   16777215
            ForeColor       =   11162880
            BackColorFixed  =   12640511
            ForeColorFixed  =   8388608
            BackColorSel    =   16777153
            ForeColorSel    =   -2147483647
            BackColorBkg    =   -2147483636
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
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   2
            RowHeightMin    =   350
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   0
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
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "Free Qty"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   1
            Left            =   2100
            TabIndex        =   44
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "Free Qty Apply (y/n)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   1
            Left            =   3060
            TabIndex        =   43
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "KIT Code"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   0
            Left            =   60
            TabIndex        =   42
            Top             =   120
            Width           =   1035
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Caption         =   "Book into this KIT"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   0
            Left            =   1080
            TabIndex        =   41
            Top             =   120
            Width           =   1035
         End
      End
      Begin VB.Frame Frame1_kit 
         Caption         =   "Kit Book List"
         Height          =   3315
         Left            =   8190
         TabIndex        =   34
         Top             =   495
         Visible         =   0   'False
         Width           =   6075
         Begin VB.ListBox List1_Book 
            Appearance      =   0  'Flat
            Height          =   2832
            Left            =   60
            Style           =   1  'Checkbox
            TabIndex        =   35
            Top             =   300
            Width           =   5955
         End
      End
      Begin VB.CheckBox Check1_kit 
         Caption         =   "KIT BOOK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   7935
         TabIndex        =   33
         Top             =   180
         Width           =   1395
      End
      Begin VB.TextBox txtBookInGaddi 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   6165
         MaxLength       =   4
         TabIndex        =   9
         Top             =   2580
         Width           =   675
      End
      Begin VB.TextBox txtBname_binder 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   330
         Left            =   1410
         MaxLength       =   60
         TabIndex        =   2
         Top             =   1350
         Width           =   5430
      End
      Begin VB.TextBox txtBcode 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1410
         MaxLength       =   10
         TabIndex        =   0
         Top             =   585
         Width           =   2370
      End
      Begin VB.TextBox txtDis 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1425
         MaxLength       =   5
         TabIndex        =   7
         Top             =   2535
         Width           =   1035
      End
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2940
         MaxLength       =   6
         TabIndex        =   8
         Top             =   2520
         Width           =   675
      End
      Begin VB.TextBox txtBName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   330
         Left            =   1410
         MaxLength       =   60
         TabIndex        =   1
         Top             =   960
         Width           =   5430
      End
      Begin VB.ComboBox cboGcode 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1755
         Width           =   2235
      End
      Begin VB.ComboBox cboGName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1410
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   2160
         Width           =   2235
      End
      Begin VB.Frame buttonFrame 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   135
         TabIndex        =   16
         Top             =   3975
         Width           =   8745
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Add Series Master"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   6615
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   45
            Width           =   1005
         End
         Begin VB.CommandButton Command2SerEdit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Edit Series (in book)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   4980
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   480
            Width           =   1590
         End
         Begin VB.CommandButton cmdAddSer 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Add Series (in book)"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   4980
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   30
            Width           =   1590
         End
         Begin VB.CommandButton cmdSearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Search"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   4005
            Picture         =   "frmBookMaster.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   45
            Width           =   960
         End
         Begin VB.CommandButton cmdExit_12 
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
            Height          =   855
            Left            =   7665
            Picture         =   "frmBookMaster.frx":0BE4
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   30
            Width           =   1020
         End
         Begin VB.CommandButton cmdEdit_4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   3015
            Picture         =   "frmBookMaster.frx":17C8
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   45
            Width           =   975
         End
         Begin VB.CommandButton cmdDelete_3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1980
            Picture         =   "frmBookMaster.frx":1BD5
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   45
            Width           =   1035
         End
         Begin VB.CommandButton cmdSave_2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.4
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   990
            Picture         =   "frmBookMaster.frx":27B9
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   45
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
            Height          =   855
            Left            =   45
            Picture         =   "frmBookMaster.frx":339D
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   45
            Width           =   930
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid vs 
         Height          =   4440
         Left            =   45
         TabIndex        =   21
         Top             =   5070
         Width           =   14280
         _cx             =   25188
         _cy             =   7832
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
         BackColor       =   16777215
         ForeColor       =   11162880
         BackColorFixed  =   7917545
         ForeColorFixed  =   8388608
         BackColorSel    =   16777153
         ForeColorSel    =   -2147483647
         BackColorBkg    =   -2147483636
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
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
         ExplorerBar     =   1
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bag in Box"
         Height          =   192
         Left            =   5148
         TabIndex        =   54
         Top             =   3360
         Width           =   768
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         Height          =   195
         Left            =   2520
         TabIndex        =   51
         Top             =   2985
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Gp. Code"
         Height          =   195
         Left            =   3720
         TabIndex        =   50
         Top             =   1785
         Width           =   1005
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Gp. Name"
         Height          =   195
         Left            =   3720
         TabIndex        =   49
         Top             =   2190
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Up to 10"
         Height          =   192
         Left            =   5136
         TabIndex        =   48
         Top             =   2988
         Width           =   588
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HSN Code "
         Height          =   195
         Left            =   240
         TabIndex        =   47
         Top             =   3030
         Width           =   810
      End
      Begin VB.Label lblTotal 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   60
         TabIndex        =   32
         Top             =   7500
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Book in Gaddi"
         Height          =   315
         Left            =   5100
         TabIndex        =   31
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "{"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   1080
         TabIndex        =   30
         Top             =   855
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For Binder"
         Height          =   195
         Left            =   195
         TabIndex        =   29
         Top             =   1395
         Width           =   720
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Name"
         Height          =   195
         Left            =   195
         TabIndex        =   28
         Top             =   2205
         Width           =   900
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Code"
         Height          =   195
         Left            =   195
         TabIndex        =   27
         Top             =   630
         Width           =   1020
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book Name"
         Height          =   195
         Left            =   195
         TabIndex        =   26
         Top             =   1035
         Width           =   840
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Code"
         Height          =   240
         Left            =   195
         TabIndex        =   25
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount "
         Height          =   240
         Left            =   180
         TabIndex        =   24
         Top             =   2580
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   195
         Left            =   2520
         TabIndex        =   23
         Top             =   2580
         Width           =   660
      End
      Begin VB.Image imgFile 
         Height          =   192
         Left            =   13500
         Picture         =   "frmBookMaster.frx":3F81
         Top             =   9180
         Visible         =   0   'False
         Width           =   192
      End
      Begin VB.Label lblId 
         Height          =   285
         Left            =   3735
         TabIndex        =   22
         Top             =   630
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   1065
         Left            =   90
         Top             =   3930
         Width           =   8835
      End
   End
End
Attribute VB_Name = "frmBookMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit1 As Boolean
Private Sub cboGcode_Click()
   
   If cboGcode = "" Then Exit Sub
   If RS.State = 1 Then RS.Close
   RS.Open "select groupname from groups where " & stringyear & " and groupcode='" & Trim(cboGcode) & "'", con
   If RS.EOF = False Then
      cboGName.text = RS(0)
      fillGrid
   End If


End Sub
Private Sub cboGcode_sub_Click()

If cboGcode_sub = "" Then Exit Sub
If RS.State = 1 Then RS.Close
RS.Open "select groupname from groups where " & stringyear & " and groupcode='" & Trim(cboGcode_sub) & "'", con
If RS.EOF = False Then
   cboGName_sub.text = RS(0)
End If

End Sub

Private Sub cboGName_Click()
  If cboGName = "" Then Exit Sub
   
   If RS.State = 1 Then RS.Close
   RS.Open "select groupcode from groups where " & stringyear & " and groupname='" & Trim(cboGName) & "'", con
   If RS.EOF = False Then
      cboGcode.text = RS(0)
   End If
End Sub

Private Sub cboGName_sub_Click()
  If cboGName_sub = "" Then Exit Sub
   
   If RS.State = 1 Then RS.Close
   RS.Open "select groupcode from groups where " & stringyear & " and groupname='" & Trim(cboGName_sub) & "'", con
   If RS.EOF = False Then
      cboGcode.text = RS(0)
   End If
End Sub

Private Sub Check1_editkit_Click()
 If Check1_editkit.value = 0 Then
    frmEditBookKit.Visible = False
 Else
    frmEditBookKit.Visible = True
 End If
End Sub

Private Sub Check1_kit_Click()

If txtBcode = "" Then
   MsgBox "Select Book Name .... ", vbCritical
   Exit Sub
End If

If Check1_kit.value = 1 Then
   Frame1_kit.Visible = True
Else
   Frame1_kit.Visible = False
End If

End Sub

Private Sub cmdAdd_1_Click()

frmEditBookKit.Visible = False
txtBcode = ""
txtBName = ""
cboGcode = ""
cboGName = ""
txtbagInbox.text = ""
cboGcode_sub = ""
cboGName_sub = ""
txtDis = ""
txtRate = ""
txtBname_binder = ""
txthsncode.text = ""
txtupto10.text = ""
txtBookInGaddi.text = ""

txtClass.text = ""

edit1 = False

cmdSave_2.Enabled = True
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False

fillGrid


txtBcode.SetFocus

   
   
End Sub

Private Sub cmdAddSer_Click()
frmAddSeries.Show
End Sub

Private Sub cmdDelete_3_Click()

If MsgBox("want To delete ?", vbInformation + vbYesNo) = vbYes Then

 con.BeginTrans
 con.Execute "delete from  books where bookcode='" & txtBcode & "' and " & stringyear
 con.CommitTrans
 
End If

cmdAdd_1_Click
End Sub
Private Sub cmdEdit_4_Click()
cmdEdit_4.Enabled = False
cmdSave_2.Enabled = True
cmdDelete_3.Enabled = True
cmdSave_2.SetFocus
edit1 = True
End Sub
Private Sub cmdExit_12_Click()
  Unload Me
End Sub
Sub upDateKitBook()

'---------------------------------------------------------------------------
'con.Execute "delete from BOOKS_KIT where " & stringyear & " and kitcode='" & txtBcode.Text & "'"

For k1 = 0 To List1_Book.ListCount - 1
If List1_Book.Selected(k1) = True Then
   sss = Mid(UCase(List1_Book.List(k1)), 1, InStr(UCase(List1_Book.List(k1)), "=>") - 1)
   If rs1.State = 1 Then rs1.Close
   rs1.Open "SELECT * from BOOKS_KIT where " & stringyear & " and kitcode='" & txtBcode.text & "' AND BOOKCODE='" & sss & "'", con
   If rs1.EOF = True Then
      con.Execute "insert into BOOKS_KIT(BOOKCODE,KITCODE,Fyear,setupid) values('" & sss & "','" & txtBcode.text & "','" & session & "','" & setupid & "')"
   End If
   
End If
Next
'---------------------------------------------------------------------------

'For k1 = 1 To vsKit.Rows - 1
'If (vsKit.TextMatrix(k1, 0) <> "") Then
'    con.Execute "insert into BOOKS_KIT(KITCODE,BOOKCODE,Qty,Apply,Fyear,setupid) values('" & vsKit.TextMatrix(k1, 0) & "','" & vsKit.TextMatrix(k1, 1) & "','" & vsKit.TextMatrix(k1, 2) & "','" & vsKit.TextMatrix(k1, 3) & "','" & session & "','" & setupid & "')"
'End If
'Next


End Sub

Private Sub cmdExit_Click()
frmEditBookKit.Visible = False
End Sub

Private Sub cmdSave_2_Click()

If txtBcode = "" Then
   MsgBox "Enter Book Code. ...", vbInformation
   txtBcode.SetFocus
   Exit Sub
End If
   
   
If txtBName = "" Then
   MsgBox "Enter Book Name. ...", vbInformation
   txtBName.SetFocus
   Exit Sub
End If
   
If cboGcode = "" Then
   MsgBox "Select Group Code ...", vbInformation
   cboGcode.SetFocus
   Exit Sub
End If

If txtClass.text = "" Then
   
   MsgBox "Enter Class Name ...", vbInformation
   txtClass.SetFocus
   Exit Sub
End If


If edit1 = True Then
   
   con.Execute "update  [books] set BOOKNAME='" & UCase(txtBName) & "',BookDes='" & UCase(txtBname_binder) & "'" & _
   ",GROUPCODE='" & UCase(cboGcode) & "',GROUPCODE_sub='" & UCase(cboGcode_sub) & "',rate=" & txtRate & ",discount=" & txtDis & ",BooksInGaddi=" & Val(txtBookInGaddi.text) & ",hsncode='" & txthsncode.text & "'" & _
   ",bkclass='" & txtClass.text & "',noofbox='" & txtbagInbox.text & "' where " & stringyear & " and Bookcode='" & UCase(Trim(txtBcode)) & "'"
 
   ' code for Binder
   
   If RS.State = 1 Then RS.Close
   RS.Open "select bookno from BookMaster where bookno='" & UCase(Trim(txtBcode)) & "'", con
   If RS.EOF = False Then
        con.Execute "update  BookMaster set book='" & UCase(txtBname_binder.text) & "',bookfont='e'" & _
        " where " & stringyear & " and Bookno='" & UCase(Trim(txtBcode)) & "'"
   Else
        con.BeginTrans
        con.Execute "INSERT INTO  [BookMaster]" & _
                  "(bookno" & _
                  ",Book" & _
                  ",bookfont" & _
                  ",class" & _
                  ",fyear" & _
                  ",[setupid]" & _
            ") Values" & _
                  "('" & UCase(txtBcode) & "'" & _
                  ",'" & UCase(txtBName) & "'" & _
                  ",'e'" & _
                  ",'" & UCase(cboGName) & "'" & _
                  ",'" & main.session & "'" & _
                  ",'" & main.setupid & "')"
        con.CommitTrans
   End If
   
   
   
   upDateKitBook
 
Else
   
   
If RS.State = 1 Then RS.Close
RS.Open "SELECT TOP 1 * FROM books WHERE Bookcode='" & UCase(Trim(txtBcode)) & "'", con
If RS.EOF = False Then
   MsgBox "Duplicate Record Code...", vbCritical
   txtBcode.SetFocus
   Exit Sub
End If
   
  
   
con.BeginTrans
con.Execute "INSERT INTO  [books]" & _
           "(Bookcode" & _
           ",BOOKNAME" & _
           ",BookDes" & _
           ",GROUPCODE" & _
           ",GROUPCODE_sub" & _
           ",DISCOUNT" & _
           ",Rate" & _
           ",fyear" & _
           ",[setupid],HSNCode,bkclass,noofbox" & _
     ") Values" & _
           "('" & UCase(txtBcode) & "'" & _
           ",'" & UCase(txtBName) & "'" & _
           ",'" & UCase(txtBname_binder) & "'" & _
           ",'" & UCase(cboGcode) & "'" & _
           ",'" & UCase(cboGcode_sub) & "'" & _
           "," & IIf(txtDis = "", 0, txtDis) & "" & _
           "," & IIf(txtRate = "", 0, txtRate) & "" & _
           ",'" & main.session & "'" & _
           ",'" & main.setupid & "','" & txthsncode.text & "','" & txtClass.text & "','" & txtbagInbox.text & "')"
con.CommitTrans

'code for binder

con.BeginTrans
con.Execute "INSERT INTO  [BookMaster]" & _
           "(bookno" & _
           ",Book" & _
           ",bookfont" & _
           ",class" & _
           ",fyear" & _
           ",[setupid]" & _
     ") Values" & _
           "('" & UCase(txtBcode) & "'" & _
           ",'" & UCase(txtBName) & "'" & _
           ",'e'" & _
           ",'" & UCase(cboGName) & "'" & _
           ",'" & main.session & "'" & _
           ",'" & main.setupid & "')"
con.CommitTrans

upDateKitBook


End If


addmaster_addSingleData "Books", txtBcode

MsgBox "Date Saved ....", vbInformation
cmdSave_2.Enabled = False

Call cmdAdd_1_Click

End Sub

Private Sub cmdSearch_Click()
   searchType = "books"
   sqlQry = "select BOOKCODE,BOOKNAME from BOOKS where BOOKCODE"
   orderby = "order by BOOKCODE"
   popuplist10 "select BOOKCODE,BOOKNAME from BOOKS where " & stringyear & "  order by BOOKCODE", con

End Sub

Private Sub cmdSearch_GotFocus()
  
  
  If PopUpValue1 <> "" Then
     
    cmdSave_2.Enabled = False
    cmdEdit_4.Enabled = True
    If RS.State = 1 Then RS.Close
    RS.Open "select * from books where bookcode='" & PopUpValue1 & "' and " & stringyear & "", con
    If RS.EOF = False Then
      
      txtClass.text = RS!bkclass & ""
      txtBcode = RS!Bookcode
      txtBName = RS!Bookname
      txtBname_binder = RS!BookDes & ""
      cboGcode = RS!groupcode
      txtDis = RS!discount
      txtRate = RS!rate
      txtBookInGaddi = RS!BooksInGaddi & ""
      txthsncode.text = RS!hsncode & ""
      
      txtbagInbox.text = RS!noofbox & ""
      
      If RS.State = 1 Then RS.Close
      RS.Open "select groupname from GROUPS where groupcode='" & cboGcode & "' and " & stringyear & "", con
      If RS.EOF = False Then
           cboGName = RS(0)
      End If
      End If
    
    viewKit
    
  End If
  
  
  PopUpValue1 = ""
  PopUpValue2 = ""
  
  cmdEdit_4.Enabled = True
  

End Sub
Sub fillGrid()

Dim f_grid As New ADODB.Recordset
Dim k1 As Integer


If cboGcode.text <> "" Then
        Set f_grid = con.Execute("exec bookSearch '" & session & "'," & main.setupid & ",'" & cboGcode.text & "'")
Else
        Set f_grid = con.Execute("exec bookSearch '" & session & "'," & main.setupid & ",'" & "" & "'")
End If

Set vs.DataSource = f_grid


'===========================================================
 'If f_grid.State = 1 Then f_grid.close
 'f_grid.Open "select BookCode,BookName from BOOKS"
 List1_Book.Clear
 While f_grid.EOF = False
   List1_Book.AddItem f_grid("BookCode") & "=>" & f_grid("BookName")
   f_grid.MoveNext
 Wend
'===========================================================

vs.ColWidth(1) = 1000
k1 = 0
For I = 1 To vs.rows - 1
    vs.Cell(flexcpPicture, I, 1) = imgFile
    k1 = k1 + 1
Next

vs.ColWidth(2) = 3000
vs.ColWidth(3) = 3100

lblTotal.Caption = "Total Record : " & k1

End Sub

Private Sub Command1_Click()
 For kk1 = 0 To vsKit.rows - 1
     If vsKit.TextMatrix(kk1, 2) <> "" Then
        con.Execute "update BOOKS_KIT set Qty=" & vsKit.TextMatrix(kk1, 2) & ",apply='" & vsKit.TextMatrix(kk1, 3) & "' where (KITCODE='" & vsKit.TextMatrix(kk1, 0) & "' and BOOKCODE='" & vsKit.TextMatrix(kk1, 1) & "')"
     End If
 Next
 
 MsgBox "Data Updated....", vbInformation
End Sub

Private Sub Command2_Click()
frmSerMaster.Show
End Sub

Private Sub Command2SerEdit_Click()
frmSeriesEdit.Show

End Sub

Private Sub Form_Activate()
cmdAdd_1_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then sendkeys "{tab}"
End Sub

Private Sub Form_Load()

 Screen.MousePointer = vbHourglass

 Me.top = 1800
 Me.Left = 1500
 
 Me.Width = s
 
 
 fillGrid
 
 fillcombo cboGcode, "groupcode", "groups", con
 fillcombo cboGName, "groupname", "groups", con

 fillcombo cboGcode_sub, "groupcode", "groups", con
 fillcombo cboGName_sub, "groupname", "groups", con

 
 
 
 BackColorFrom Me
 
 Screen.MousePointer = vbDefault
 
End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub txtBcode_GotFocus()

If PopUpValue1 <> "" Then


   If RS.State = 1 Then RS.Close
   RS.Open "select BOOKCODE,BOOKNAME,GROUPCODE,RATE,DISCOUNT" & _
   ",RetailPrice,RetailDis,BooksInGaddi,BookDes,noofbox from books where " & stringyear & " and bookcode='" & PopUpValue1 & "'", con
   If RS.EOF = False Then
      txtBcode = PopUpValue1
      txtBName.text = RS!Bookname
      cboGcode = RS!groupcode
      txtBname_binder.text = RS!BookDes & ""
      txtDis = RS!discount
      txtRate = RS!rate
      txtBookInGaddi = RS!BooksInGaddi & ""
      txtbagInbox.text = RS!noofbox & ""
   End If

   
   If RS.State = 1 Then RS.Close
   RS.Open "select groupname from groups where " & stringyear & " and groupcode='" & Trim(cboGcode) & "'", con
   If RS.EOF = False Then
      cboGName.text = RS(0)
   End If
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   
   cmdSave_2.Enabled = False
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = True
   edit1 = True

End If


End Sub

Private Sub txtBcode_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    value = "select BookCode,BookName from books where " & stringyear & "  order by BookCode"
    popuplist_client value, CCON

End If

End Sub

Private Sub txtBcode_LostFocus()
txtBcode = UCase(txtBcode)
End Sub


Private Sub txtBname_binder_LostFocus()
txtBname_binder = UCase(txtBname_binder)
End Sub

Private Sub txtBName_LostFocus()
txtBName = UCase(txtBName)
End Sub

Private Sub txtBookInGaddi_GotFocus()
txtBookInGaddi.SelLength = 10
End Sub

Private Sub txtClass_KeyDown(KeyCode As Integer, Shift As Integer)
  If Len(txtClass.text) = 0 Then
     txtClass.SetFocus
  End If
End Sub

Private Sub txthsncode_KeyDown(KeyCode As Integer, Shift As Integer)
   
If (KeyCode = 13) Then

    txthsncode.SetFocus
    
End If
   
End Sub
Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then txtBookInGaddi.SetFocus
End Sub

Private Sub vs_Click()
  
  cmdSave_2.Enabled = False
  cmdEdit_4.Enabled = True
  cboGName_sub.text = ""
  
  txtBcode = vs.TextMatrix(vs.RowSel, 1)
  txtBName = vs.TextMatrix(vs.RowSel, 2)
  txtBname_binder = vs.TextMatrix(vs.RowSel, 3)
  cboGcode = vs.TextMatrix(vs.RowSel, 4)
  
  If RS.State = 1 Then RS.Close
  RS.Open "select groupname from GROUPS where " & stringyear & " and groupcode='" & vs.TextMatrix(vs.RowSel, 4) & "'", con
  If RS.EOF = False Then
     cboGName = RS(0)
  End If
  
  vs.Editable = flexEDNone
  If (vs.Col = 8 Or vs.Col = 9 Or vs.Col = 10) Then
     vs.Editable = flexEDKbdMouse
     con.Execute "update BOOKS set NoPrintDesc=" & vs.TextMatrix(vs.RowSel, 8) & " where " & stringyear & " and bookcode='" & vs.TextMatrix(vs.RowSel, 1) & "'"
  End If
  
  
  txtRate = vs.TextMatrix(vs.RowSel, 5)
  txtDis = vs.TextMatrix(vs.RowSel, 6)
  txtBookInGaddi = vs.TextMatrix(vs.RowSel, 7)
  txthsncode.text = vs.TextMatrix(vs.RowSel, 11)
  txtupto10.text = vs.TextMatrix(vs.RowSel, 12)
  
  cboGcode_sub.text = vs.TextMatrix(vs.RowSel, 13)
  
  txtbagInbox.text = ""
  If RS.State = 1 Then RS.Close
  RS.Open "select bkclass,noofbox from bookS where " & stringyear & " and bookcode='" & txtBcode.text & "'", con
  If RS.EOF = False Then
     txtClass.text = RS(0) & ""
     txtbagInbox.text = RS(1) & ""
  End If
  
  
  If RS.State = 1 Then RS.Close
  RS.Open "select groupname from GROUPS where " & stringyear & " and groupcode='" & vs.TextMatrix(vs.RowSel, 13) & "'", con
  If RS.EOF = False Then
     cboGName_sub = RS(0)
  End If
  
  viewKit
  



End Sub
Sub viewKit()
   
  Dim kk1 As Integer
  kk1 = 0
  
  vsKit.Clear
   
  For k1 = 0 To List1_Book.ListCount - 1
    List1_Book.Selected(k1) = False
  Next
  
  If RS.State = 1 Then RS.Close
  RS.Open "select bookcode,kitcode,Qty,Apply from BOOKS_KIT where " & stringyear & " and kitcode='" & txtBcode & "'", con
  If RS.EOF = True Then
     Frame1_kit.Visible = False
     Check1_kit.value = 0
  End If
  
  If RS.EOF = False Then
     If Len(RS!Kitcode) > 1 Then
        Check1_kit.value = 1
        
        For J = 0 To RS.RecordCount - 1
        For k1 = 0 To List1_Book.ListCount - 1
              aaa = Mid(UCase(List1_Book.List(k1)), 1, InStr(UCase(List1_Book.List(k1)), "=>") - 1)
              If Mid(UCase(List1_Book.List(k1)), 1, InStr(UCase(List1_Book.List(k1)), "=>") - 1) = RS!Bookcode Then
                 List1_Book.Selected(k1) = True
                 vsKit.TextMatrix(kk1, 0) = RS!Kitcode
                 vsKit.TextMatrix(kk1, 1) = RS!Bookcode
                 vsKit.TextMatrix(kk1, 3) = RS!Apply & ""
                 
                 If IsNull(RS!qty) Then
                    vsKit.TextMatrix(kk1, 2) = ""
                    kk1 = kk1 + 1
                 Else
                    vsKit.TextMatrix(kk1, 2) = RS!qty
                    kk1 = kk1 + 1
                    
                 End If
                 
              End If
           Next
           RS.MoveNext
        Next
        
     Else
        Check1_kit.value = 0
     End If
     
  End If
  
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then
 If (vs.Col = 9 Or vs.Col = 10) Then
    sendkeys "{down}"
 End If
End If
 
End Sub

Private Sub vs_SelChange()
   If (vs.Col = 9 Or vs.Col = 8 Or vs.Col = 10) Then
      
      DoEvents
      DoEvents
      DoEvents
      
      vs.Editable = flexEDKbdMouse
      If (vs.Col = 8 Or vs.Col = 9 Or vs.Col = 10) Then
      con.Execute "update BOOKS set KITCODE='" & vs.TextMatrix(vs.RowSel, 9) & "',sername='" & vs.TextMatrix(vs.RowSel, 10) & "' where " & stringyear & " and bookcode='" & vs.TextMatrix(vs.RowSel, 1) & "'"
      End If
      DoEvents
      DoEvents
      DoEvents


   Else
      vs.Editable = flexEDNone
   End If
End Sub
