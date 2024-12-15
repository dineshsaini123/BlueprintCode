VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmINVOrder 
   Caption         =   "Order Creation"
   ClientHeight    =   10596
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   14172
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmINVOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10596
   ScaleWidth      =   14172
   Begin VB.ComboBox cboDisType 
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
      ForeColor       =   &H000000FF&
      Height          =   312
      ItemData        =   "frmINVOrder.frx":0442
      Left            =   1296
      List            =   "frmINVOrder.frx":045B
      Style           =   2  'Dropdown List
      TabIndex        =   136
      Top             =   3456
      Width           =   4032
   End
   Begin VB.CommandButton cmdDis 
      BackColor       =   &H00C0FFFF&
      Caption         =   "S.Discount"
      Height          =   336
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   135
      Top             =   648
      Visible         =   0   'False
      Width           =   1596
   End
   Begin VB.TextBox txtbagInBox 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7056
      TabIndex        =   133
      Top             =   8604
      Width           =   1056
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Next Trans"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   516
      Left            =   9648
      Picture         =   "frmINVOrder.frx":04C9
      Style           =   1  'Graphical
      TabIndex        =   131
      Top             =   8676
      Visible         =   0   'False
      Width           =   1344
   End
   Begin VB.CheckBox Check1_onlySp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Only Pending Sp. Books"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Left            =   7308
      TabIndex        =   130
      Top             =   10152
      Width           =   4032
   End
   Begin VB.TextBox txtRem 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   840
      Left            =   72
      MaxLength       =   250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   125
      Top             =   8280
      Width           =   5172
   End
   Begin VB.ComboBox cbofrt 
      Height          =   300
      ItemData        =   "frmINVOrder.frx":10AD
      Left            =   9516
      List            =   "frmINVOrder.frx":10BA
      TabIndex        =   5
      Top             =   348
      Width           =   1104
   End
   Begin VB.TextBox txtschool 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8352
      MaxLength       =   150
      TabIndex        =   24
      Top             =   2880
      Width           =   4536
   End
   Begin VB.TextBox txtPIN_Ship 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   10764
      MaxLength       =   100
      TabIndex        =   16
      Top             =   1980
      Width           =   900
   End
   Begin VB.ListBox List_emptyList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2136
      Left            =   252
      TabIndex        =   121
      Top             =   1008
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton cmdListBlankOrd 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Empty Ord.No"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   252
      Style           =   1  'Graphical
      TabIndex        =   120
      Top             =   648
      Width           =   948
   End
   Begin VB.CheckBox Check1_notefull 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Only Pending Order (Not Complete Order)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Left            =   3240
      TabIndex        =   119
      Top             =   10152
      Width           =   4032
   End
   Begin VB.Frame frmMail 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pending Mail..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4752
      Left            =   36
      TabIndex        =   101
      Top             =   3852
      Visible         =   0   'False
      Width           =   13788
      Begin VB.TextBox txt_NetAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11115
         Locked          =   -1  'True
         TabIndex        =   111
         Top             =   4380
         Width           =   1185
      End
      Begin VB.CommandButton cmdHide 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remove From List"
         Height          =   615
         Left            =   12552
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   864
         Width           =   1020
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Exit"
         Height          =   570
         Left            =   12552
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   1536
         Width           =   1020
      End
      Begin VB.CommandButton cmdMail 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Send Mail"
         Height          =   570
         Left            =   12552
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   240
         Width           =   1020
      End
      Begin VSFlex7Ctl.VSFlexGrid vs_Mail 
         Height          =   4068
         Left            =   96
         TabIndex        =   104
         Top             =   240
         Width           =   12300
         _cx             =   21696
         _cy             =   7175
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   16761024
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16761992
         ForeColorSel    =   0
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   200
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmINVOrder.frx":10CE
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
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Net Amt. :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9768
         TabIndex        =   112
         Top             =   4416
         Width           =   1320
      End
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   11808
      Top             =   10044
   End
   Begin VB.ComboBox cbobiltyrem 
      Height          =   300
      ItemData        =   "frmINVOrder.frx":1157
      Left            =   8325
      List            =   "frmINVOrder.frx":1164
      TabIndex        =   6
      Top             =   756
      Width           =   3480
   End
   Begin VB.TextBox txtNoOfGaddi 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   7056
      TabIndex        =   109
      Top             =   8280
      Width           =   1068
   End
   Begin VB.TextBox TXTBAL 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9684
      TabIndex        =   107
      Top             =   8304
      Width           =   1284
   End
   Begin VB.CommandButton cmdImportExcel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "U&pdate Qty"
      Height          =   465
      Left            =   6075
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   3156
      Width           =   735
   End
   Begin VB.TextBox txtpath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2070
      MaxLength       =   80
      TabIndex        =   99
      Top             =   3156
      Width           =   3225
   End
   Begin VB.CommandButton cmdAddFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add &File"
      Height          =   465
      Left            =   5310
      Style           =   1  'Graphical
      TabIndex        =   98
      Top             =   3156
      Width           =   735
   End
   Begin VB.ComboBox cboBilty 
      Height          =   300
      ItemData        =   "frmINVOrder.frx":1190
      Left            =   8325
      List            =   "frmINVOrder.frx":119D
      TabIndex        =   4
      Top             =   348
      Width           =   1164
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   11376
      Top             =   10044
   End
   Begin VB.TextBox txtPin 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3690
      MaxLength       =   100
      TabIndex        =   87
      Top             =   1920
      Width           =   900
   End
   Begin VB.CheckBox Check1_AddOrderNo 
      Caption         =   "Add order no manually"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   264
      TabIndex        =   84
      Top             =   12
      Width           =   1695
   End
   Begin VB.CheckBox Check1_filter 
      Caption         =   "Not Filter Again"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   12450
      TabIndex        =   81
      Top             =   10176
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Frame frmPendingClear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pending Clear"
      Height          =   4512
      Left            =   324
      TabIndex        =   77
      Top             =   3888
      Visible         =   0   'False
      Width           =   13515
      Begin VSFlex7Ctl.VSFlexGrid vs1 
         Height          =   3756
         Left            =   -816
         TabIndex        =   78
         Top             =   492
         Width           =   13308
         _cx             =   23469
         _cy             =   6615
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   16761024
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16761992
         ForeColorSel    =   0
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   200
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmINVOrder.frx":11B0
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
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   435
         Left            =   11220
         TabIndex        =   80
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   435
         Left            =   10080
         TabIndex        =   79
         Top             =   120
         Width           =   1095
      End
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   312
      Left            =   120
      TabIndex        =   75
      Top             =   10440
      Visible         =   0   'False
      Width           =   10872
      _ExtentX        =   19177
      _ExtentY        =   550
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   444
      Left            =   10656
      TabIndex        =   72
      Top             =   288
      Width           =   3432
      Begin VB.OptionButton Option2_sp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sp. Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1764
         TabIndex        =   74
         Top             =   132
         Width           =   1608
      End
      Begin VB.OptionButton Option1_sale 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sale Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   276
         Left            =   48
         TabIndex        =   73
         Top             =   132
         Value           =   -1  'True
         Width           =   1740
      End
   End
   Begin VB.Frame frmbk 
      Height          =   1008
      Left            =   11016
      TabIndex        =   69
      Top             =   8172
      Width           =   2892
      Begin VB.OptionButton Option2_seriseWise 
         Caption         =   "Serise Wise"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   1404
         TabIndex        =   129
         Top             =   756
         Width           =   1380
      End
      Begin VB.OptionButton Option1_bookwise 
         Caption         =   "Book Wise"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   36
         TabIndex        =   128
         Top             =   756
         Value           =   -1  'True
         Width           =   1308
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   936
         TabIndex        =   70
         Top             =   108
         Width           =   1116
      End
      Begin VB.ComboBox cbogd1 
         Appearance      =   0  'Flat
         Height          =   300
         ItemData        =   "frmINVOrder.frx":1271
         Left            =   2124
         List            =   "frmINVOrder.frx":127E
         Style           =   2  'Dropdown List
         TabIndex        =   85
         Top             =   108
         Width           =   696
      End
      Begin VB.TextBox txtbk 
         Height          =   300
         Left            =   948
         TabIndex        =   71
         Top             =   408
         Width           =   1152
      End
      Begin VB.Label Label10 
         Caption         =   "B.Code"
         Height          =   228
         Left            =   72
         TabIndex        =   127
         Top             =   468
         Width           =   876
      End
      Begin VB.Label Label9 
         Caption         =   "Ser.Name"
         Height          =   228
         Left            =   72
         TabIndex        =   126
         Top             =   144
         Width           =   876
      End
   End
   Begin VB.ComboBox txtMark 
      Appearance      =   0  'Flat
      Height          =   300
      ItemData        =   "frmINVOrder.frx":128B
      Left            =   1305
      List            =   "frmINVOrder.frx":1298
      TabIndex        =   22
      Top             =   3156
      Width           =   750
   End
   Begin VB.TextBox txtContectNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8340
      MaxLength       =   50
      TabIndex        =   18
      Top             =   2280
      Width           =   5400
   End
   Begin VB.TextBox txtPartyName 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   6840
      MaxLength       =   100
      TabIndex        =   64
      Top             =   1032
      Visible         =   0   'False
      Width           =   168
   End
   Begin VB.TextBox txtPartySt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5070
      MaxLength       =   100
      TabIndex        =   11
      Top             =   1932
      Width           =   1755
   End
   Begin VB.TextBox txtShipState 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12192
      MaxLength       =   50
      TabIndex        =   17
      Top             =   1980
      Width           =   1548
   End
   Begin VB.TextBox txtBankName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8340
      MaxLength       =   80
      TabIndex        =   13
      Top             =   1380
      Width           =   5400
   End
   Begin VB.TextBox txtDist_ship 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8340
      MaxLength       =   50
      TabIndex        =   15
      Top             =   1980
      Width           =   1968
   End
   Begin VB.TextBox txtDist 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1305
      MaxLength       =   100
      TabIndex        =   10
      Top             =   1932
      Width           =   1890
   End
   Begin VB.TextBox txtShip 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8340
      MaxLength       =   100
      TabIndex        =   12
      Top             =   1056
      Width           =   5400
   End
   Begin VB.CheckBox Check1_school 
      Caption         =   "Select School"
      Height          =   240
      Left            =   11844
      TabIndex        =   54
      Top             =   816
      Width           =   1515
   End
   Begin VB.ComboBox cmbAgentName 
      Appearance      =   0  'Flat
      Height          =   300
      ItemData        =   "frmINVOrder.frx":12A5
      Left            =   8340
      List            =   "frmINVOrder.frx":12A7
      TabIndex        =   23
      Top             =   2568
      Width           =   5370
   End
   Begin VB.TextBox txtScId 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   12960
      TabIndex        =   49
      Top             =   2880
      Width           =   735
   End
   Begin VB.TextBox txtBillQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7080
      TabIndex        =   48
      Text            =   "0"
      Top             =   7800
      Width           =   1080
   End
   Begin VB.TextBox txtTotQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5688
      TabIndex        =   47
      Text            =   "0"
      Top             =   7800
      Width           =   1212
   End
   Begin VB.ComboBox cboOrderBy 
      Height          =   300
      ItemData        =   "frmINVOrder.frx":12A9
      Left            =   5205
      List            =   "frmINVOrder.frx":12B9
      TabIndex        =   2
      Top             =   336
      Width           =   1635
   End
   Begin VB.TextBox txtBankAdd 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8340
      MaxLength       =   80
      TabIndex        =   14
      Top             =   1680
      Width           =   5400
   End
   Begin VB.TextBox txtBookingStn 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8340
      MaxLength       =   50
      TabIndex        =   25
      Top             =   3480
      Width           =   5355
   End
   Begin VB.TextBox txtTransAdd 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1305
      MaxLength       =   50
      TabIndex        =   20
      Top             =   2550
      Width           =   5520
   End
   Begin VB.ComboBox cboTrans 
      Height          =   300
      Left            =   1305
      TabIndex        =   19
      Text            =   "cboTrans"
      Top             =   2232
      Width           =   5550
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   75
      TabIndex        =   40
      Top             =   9348
      Width           =   13956
      Begin VB.CommandButton cmd_update 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Modify"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   624
         Left            =   1800
         Picture         =   "frmINVOrder.frx":12E7
         Style           =   1  'Graphical
         TabIndex        =   113
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdSendMail 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Send Mail Option"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   624
         Left            =   11868
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   120
         Width           =   1044
      End
      Begin VB.CommandButton cmdOnlinePrint 
         BackColor       =   &H00FFFFC0&
         Caption         =   "O&nline Bill Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   624
         Left            =   10890
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   120
         Width           =   960
      End
      Begin VB.CommandButton cmdPrintLabel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Print &Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   624
         Left            =   9990
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdPrintPending 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Print Pending Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   624
         Left            =   9045
         Picture         =   "frmINVOrder.frx":1729
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   120
         Width           =   945
      End
      Begin VB.CommandButton cmdPendingSpOrder 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pending &Sp. Order List"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   6135
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   120
         Width           =   1065
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pending &Clear"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   624
         Left            =   8145
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdpendingbook 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pending &Book wise"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   624
         Left            =   7215
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdPending 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pending &Order List"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   624
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFC0&
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
         Height          =   624
         Left            =   12912
         Picture         =   "frmINVOrder.frx":230D
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   984
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFC0&
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
         Height          =   624
         Left            =   3528
         Picture         =   "frmINVOrder.frx":2EF1
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete_3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   624
         Left            =   2712
         Picture         =   "frmINVOrder.frx":32FE
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Width           =   825
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFC0&
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
         Height          =   624
         Left            =   876
         Picture         =   "frmINVOrder.frx":3EE2
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   900
      End
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFC0&
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
         Height          =   624
         Left            =   45
         Picture         =   "frmINVOrder.frx":4AC6
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   810
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   624
         Left            =   4392
         Picture         =   "frmINVOrder.frx":56AA
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   120
         Width           =   816
      End
   End
   Begin Crystal.CrystalReport CR 
      Left            =   12900
      Top             =   9960
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComCtl2.DTPicker txtOrderDate 
      Height          =   312
      Left            =   2892
      TabIndex        =   1
      Top             =   336
      Width           =   1392
      _ExtentX        =   2455
      _ExtentY        =   550
      _Version        =   393216
      Format          =   139264001
      CurrentDate     =   39500
   End
   Begin VB.TextBox txtPartyAdd2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1305
      MaxLength       =   100
      TabIndex        =   9
      Top             =   1632
      Width           =   5520
   End
   Begin VB.TextBox txtPartyAdd1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1305
      MaxLength       =   100
      TabIndex        =   8
      Top             =   1332
      Width           =   5520
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9945
      TabIndex        =   33
      Text            =   "0"
      Top             =   7836
      Width           =   1284
   End
   Begin VB.TextBox txtNarration 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1305
      MaxLength       =   80
      TabIndex        =   21
      Top             =   2856
      Width           =   5520
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1305
      MaxLength       =   40
      TabIndex        =   7
      Top             =   1020
      Width           =   5520
   End
   Begin VB.TextBox txtOrderNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1305
      TabIndex        =   0
      Top             =   300
      Width           =   1080
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3948
      Left            =   96
      TabIndex        =   26
      Top             =   3888
      Width           =   13956
      _cx             =   24617
      _cy             =   6964
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777215
      ForeColorSel    =   0
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   450
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmINVOrder.frx":628E
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
   Begin MSMask.MaskEdBox txtschool1 
      Height          =   288
      Left            =   13788
      TabIndex        =   123
      Top             =   2880
      Visible         =   0   'False
      Width           =   252
      _ExtentX        =   445
      _ExtentY        =   508
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   50
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker txtOrderDate1 
      Height          =   312
      Left            =   6888
      TabIndex        =   3
      Top             =   336
      Width           =   1428
      _ExtentX        =   2519
      _ExtentY        =   550
      _Version        =   393216
      Format          =   139264001
      CurrentDate     =   39500
   End
   Begin VB.TextBox txtGAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12285
      TabIndex        =   96
      Text            =   "0"
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox txtOnlineAmt 
      Height          =   285
      Left            =   4596
      TabIndex        =   93
      Top             =   7800
      Visible         =   0   'False
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   13410
      Top             =   5895
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount"
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   288
      TabIndex        =   137
      Top             =   3492
      Width           =   972
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Bag in Box :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   5544
      TabIndex        =   134
      Top             =   8640
      Width           =   1272
   End
   Begin VB.Label lblCAF 
      BackStyle       =   0  'Transparent
      Caption         =   "---"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   420
      Index           =   1
      Left            =   10656
      TabIndex        =   132
      Top             =   36
      Width           =   3612
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight"
      Height          =   288
      Left            =   9684
      TabIndex        =   124
      Top             =   72
      Width           =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN :"
      Height          =   300
      Index           =   27
      Left            =   10332
      TabIndex        =   122
      Top             =   1980
      Width           =   516
   End
   Begin VB.Label lblgpSchool 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   27
      Left            =   8352
      TabIndex        =   118
      Top             =   3204
      Width           =   5316
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Series Wise Discount"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   240
      Index           =   26
      Left            =   1368
      TabIndex        =   117
      Top             =   720
      Visible         =   0   'False
      Width           =   3684
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Series Wise Discount"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   276
      Index           =   25
      Left            =   1404
      MousePointer    =   1  'Arrow
      TabIndex        =   116
      Top             =   720
      Visible         =   0   'False
      Width           =   3612
   End
   Begin VB.Label lblOrder_reminder 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   276
      Left            =   5544
      TabIndex        =   115
      Top             =   9000
      Width           =   8280
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bilty Remarks :"
      Height          =   252
      Index           =   24
      Left            =   6972
      TabIndex        =   114
      Top             =   768
      Width           =   1368
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "No Of Gaddi :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   5544
      TabIndex        =   110
      Top             =   8316
      Width           =   1272
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Outstanding :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   8244
      TabIndex        =   108
      Top             =   8352
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "G.Amount :"
      Height          =   192
      Index           =   23
      Left            =   11256
      TabIndex        =   97
      Top             =   7836
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bilty :"
      Height          =   252
      Index           =   22
      Left            =   8616
      TabIndex        =   95
      Top             =   60
      Width           =   1092
   End
   Begin VB.Label lblQtySpBalance 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   288
      Left            =   180
      TabIndex        =   94
      Top             =   8664
      Width           =   4608
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CASH PARTY"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   312
      Index           =   21
      Left            =   180
      TabIndex        =   91
      Top             =   8148
      Visible         =   0   'False
      Width           =   3288
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CASH PARTY"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   492
      Index           =   20
      Left            =   228
      TabIndex        =   90
      Top             =   8196
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN :"
      Height          =   300
      Index           =   19
      Left            =   3240
      TabIndex        =   88
      Top             =   1920
      Width           =   516
   End
   Begin VB.Label lblSave 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   312
      Left            =   168
      TabIndex        =   83
      Top             =   9840
      Width           =   2952
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Godown"
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   300
      TabIndex        =   67
      Top             =   3156
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No:"
      Height          =   300
      Index           =   18
      Left            =   6996
      TabIndex        =   66
      Top             =   2244
      Width           =   1188
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "F1 For Search Party"
      Height          =   288
      Left            =   4080
      TabIndex        =   65
      Top             =   12
      Width           =   2088
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order Dt."
      Height          =   276
      Index           =   17
      Left            =   7164
      TabIndex        =   63
      Top             =   72
      Width           =   912
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   300
      Index           =   16
      Left            =   4608
      TabIndex        =   62
      Top             =   1932
      Width           =   432
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   300
      Index           =   13
      Left            =   252
      TabIndex        =   61
      Top             =   1932
      Width           =   1188
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   300
      Index           =   15
      Left            =   11676
      TabIndex        =   60
      Top             =   1980
      Width           =   528
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   300
      Index           =   14
      Left            =   6996
      TabIndex        =   59
      Top             =   1944
      Width           =   1188
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address 1"
      Height          =   300
      Index           =   12
      Left            =   276
      TabIndex        =   58
      Top             =   1332
      Width           =   1128
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address 2"
      Height          =   300
      Index           =   11
      Left            =   276
      TabIndex        =   57
      Top             =   1632
      Width           =   1068
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ship to: "
      ForeColor       =   &H80000008&
      Height          =   288
      Index           =   2
      Left            =   6996
      TabIndex        =   56
      Top             =   1092
      Width           =   960
   End
   Begin VB.Label lblBookSId 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   13440
      TabIndex        =   55
      Top             =   732
      Visible         =   0   'False
      Width           =   468
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Representative :"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   6996
      TabIndex        =   53
      Top             =   2592
      Width           =   1296
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F4 Key To Delete A Invoive Item"
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
      Left            =   216
      TabIndex        =   51
      Top             =   7836
      Width           =   2952
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "School : "
      ForeColor       =   &H80000008&
      Height          =   288
      Index           =   1
      Left            =   6996
      TabIndex        =   50
      Top             =   2892
      Width           =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order By :"
      Height          =   300
      Index           =   10
      Left            =   4320
      TabIndex        =   46
      Top             =   336
      Width           =   912
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address 2"
      Height          =   300
      Index           =   9
      Left            =   6996
      TabIndex        =   45
      Top             =   1680
      Width           =   1068
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address 1"
      Height          =   300
      Index           =   8
      Left            =   6996
      TabIndex        =   44
      Top             =   1380
      Width           =   1128
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Stn."
      Height          =   300
      Index           =   7
      Left            =   6996
      TabIndex        =   43
      Top             =   3480
      Width           =   1188
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   300
      Index           =   5
      Left            =   270
      TabIndex        =   42
      Top             =   2550
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Transport"
      Height          =   300
      Index           =   3
      Left            =   276
      TabIndex        =   41
      Top             =   2232
      Width           =   1008
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 For Search Order"
      Height          =   288
      Left            =   2112
      TabIndex        =   39
      Top             =   12
      Width           =   1848
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NetAmount :"
      Height          =   192
      Index           =   6
      Left            =   8952
      TabIndex        =   38
      Top             =   7836
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      Height          =   300
      Index           =   4
      Left            =   270
      TabIndex        =   37
      Top             =   2850
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date "
      Height          =   276
      Index           =   1
      Left            =   2436
      TabIndex        =   36
      Top             =   396
      Width           =   492
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name"
      Height          =   240
      Index           =   2
      Left            =   276
      TabIndex        =   35
      Top             =   1020
      Width           =   1248
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order No :"
      Height          =   276
      Index           =   0
      Left            =   276
      TabIndex        =   34
      Top             =   348
      Width           =   1056
   End
End
Attribute VB_Name = "frmINVOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit As Boolean
Dim sale_sp_ As String
Dim search_ As String
Dim bb As Boolean
Dim emptyInv_bool As Boolean
Dim seriesWiseDis_ As String
Dim add_ As Boolean
Dim party_type As String
Private Sub cboBilty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   cbofrt.SetFocus
End If
End Sub
Private Sub cboDisType_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   'cmbAgentName.SetFocus
   'If cboDisType.text <> "" Then
      txtContectNo.SetFocus
   'End If
End If

End Sub
Private Sub cbofrt_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   txtParty.SetFocus
End If

End Sub
Private Sub cboOrderBy_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
 If cboOrderBy <> "" Then
    txtOrderDate1.SetFocus
 End If
End If

End Sub
Private Sub cboOrderBy_LostFocus()
If cboOrderBy = "" Then
  cboOrderBy.SetFocus
End If
End Sub
Private Sub cboTrans_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If cboTrans <> "" Then
   txtTransAdd.SetFocus
End If
End If
End Sub

Private Sub cmbAgentName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  txtSchool.SetFocus
End If
End Sub
Private Sub cmbAgentName_LostFocus()

If (cmbAgentName.text <> "") Then

    If RS.State = 1 Then RS.close
    RS.Open "select Rep as Representative from SalesRepQry where rep='" & cmbAgentName.text & "'", CON_blue
    If RS.EOF = True Then
       MsgBox "Not Valid Agent Name...", vbCritical
       cmbAgentName.SetFocus
    End If

End If

End Sub

Private Sub cmd_update_Click()

   'PartyLedgerNew txtPartyName.Text
   'seriesWiseMessage
   
   'seriesWiseDis_ = " " & seriesWiseDis_
   
   
   
   
   
   If MsgBox("Want to Modify ?", vbQuestion + vbYesNo) = vbNo Then
      Exit Sub
   End If
   
   TXTBAL.text = 0
   If Option1_sale.value = True Then
    
   PartyLedgerNew txtPartyName.text
    
   End If
   
   reserial
   
      
   createLog UserName, txtOrderNo.text, "SaleOrder", " Edit : Qty " & txtTotQty.text, Date

   Set RS = New ADODB.Recordset
   RS.Open "select * from ordera where invoiceno=" & txtOrderNo & "", con, adOpenDynamic, adLockOptimistic
   If RS.EOF = False Then
            RS!through_ = cboDisType.text
            RS!pin_ship = txtPIN_Ship.text
            RS!GroupOfSchool = lblgpSchool(27).Caption
            RS!ccattach = cbobiltyrem.text
            RS!noofgaddi = IIf(txtNoOfGaddi.text = "", 0, txtNoOfGaddi.text)
            RS!bal = IIf(TXTBAL.text = "", 0, TXTBAL.text)
            RS!bilty = cboBilty.text
            RS!sale_sp = sale_sp_
            RS!ContactNo = Trim(txtContectNo)
            RS!Godown = txtMark.text
            RS!partyname = txtPartyName.text
            RS!RepName = Trim(cmbAgentName.text)
            RS!scname = Trim(txtSchool)
            RS!scid = Trim(txtScId.text)
            RS!Shipto = Trim(txtShip.text)
            RS!invoiceNo = Val(txtOrderNo)
            RS!invoiceDate = txtOrderDate
            RS!subledger = Trim(txtParty)
            RS!address1 = Trim(txtPartyAdd1)
            RS!address2 = Trim(txtPartyAdd2)
            RS!orderby = cboOrderBy
            RS!ORDERDATE = txtOrderDate1.value
            RS!transport = Trim(cboTrans)
            RS!TransAdd = Trim(txtTransAdd)
            RS!narration = Trim(txtNarration)
            RS!Shipto = Trim(txtShip.text)
            RS!Shipto_Add1 = Trim(UCase(txtBankName))
            RS!Shipto_Add2 = Trim(UCase(txtBankAdd))
            RS!shipto_dist = Trim(UCase(txtDist_ship))
            RS!Shipto_States = UCase(txtShipState.text)
          
            RS!BookingStn = Trim(txtBookingStn)
            RS!netamount = Round(txtTotal, 0)
            RS!gamount = Round(txtGAmt, 0)
            
            RS!party_dist = Trim(txtDist)
            RS!fyear = session
            RS!setupid = setupid
   
            RS!party_state = UCase(txtPartySt)
            RS!ContactNo = Trim(txtContectNo)
            
            RS!Frt_Yes = cbofrt.text

            RS.update
   End If
   
   
   '''' OrderB , Modify
   ''con.Execute "delete orderb where invoiceno=" & txtOrderNo & ""
   
'   If RS.State = 1 Then RS.close
'   RS.Open "select * from orderb where invoiceno=" & txtOrderNo & "", con, adOpenDynamic, adLockOptimistic

   
   Dim bagInBox, noofgaddi_ As Double
   
   noofgaddi_ = 0
   bagInBox = 0
   
   
   
   For I = 1 To vs.rows - 1
            If vs.TextMatrix(I, 1) <> "" Then
                  
                ''========================
                   qty = 0
                   net = 0
                   qty = IIf(vs.TextMatrix(I, 3) = "", 0, vs.TextMatrix(I, 3))
                   
                   If Option1_sale.value = False Then
                   qty = Val(qty) + Val(IIf(vs.TextMatrix(I, 4) = "", 0, vs.TextMatrix(I, 4)))
                   End If
                   
                   net = (qty * Val(vs.TextMatrix(I, 6)))
                   
                   Set rs1 = con.Execute("exec fatchDiscount_partywise '" & Trim(Mid(txtPartyName.text, 1, 5)) & "', '" & vs.TextMatrix(I, 1) & "'")
                   If rs1.EOF = False Then
                   vs.TextMatrix(I, 9) = rs1(0)
                   End If
                   
                   
                   
                   If (DateDiff("d", Now, SessionLastDate) <= 0) Then
                        
                        Set rs1 = con.Execute("exec fatchDiscount_SeriresAndpartywisefinal_test '" & Trim(Mid(txtPartyName.text, 1, 5)) & "', '" & vs.TextMatrix(I, 1) & "','" & txtScId.text & "'")
                        If rs1.EOF = False Then
                           vs.TextMatrix(I, 9) = rs1(0)
                           vs.TextMatrix(I, 12) = "y"
                        Else
                           vs.TextMatrix(I, 12) = "n"
                        End If
                        
                   Else
                        Set rs1 = con_LAST.Execute("exec fatchDiscount_SeriresAndpartywisefinal_test '" & Trim(Mid(txtPartyName.text, 1, 5)) & "', '" & vs.TextMatrix(I, 1) & "','" & txtScId.text & "'")
                        If rs1.EOF = False Then
                           vs.TextMatrix(I, 9) = rs1(0)
                           vs.TextMatrix(I, 12) = "y"
                        Else
                           vs.TextMatrix(I, 12) = "n"
                        End If

                   End If

                   
                   
                   If vs.TextMatrix(I, 9) <> "" Then
                   vs.TextMatrix(I, 7) = net - Format(Round(net * (vs.TextMatrix(I, 9) / 100), 2), "0.00")
                   End If
                   vs.TextMatrix(I, 11) = (qty * IIf(vs.TextMatrix(I, 6) = "", 0, vs.TextMatrix(I, 6)))
                   
                ''========================
                
                 If RS.State = 1 Then RS.close
                 RS.Open "select * from orderb where (invoiceno=" & txtOrderNo & " and bookcode='" & vs.TextMatrix(I, 1) & "')", con, adOpenDynamic, adLockOptimistic
                 If RS.EOF = True Then
                     RS.AddNew
                 End If
                 
                 RS!invoiceNo = Val(txtOrderNo)
                 RS!invoiceDate = txtOrderDate
                 RS!PRINTORDER = I
                 RS!Bookcode = vs.TextMatrix(I, 1)
                 RS!QUANTITY = Val(vs.TextMatrix(I, 3))
                 If vs.TextMatrix(I, 4) <> "" Then
                 RS!Spqty = vs.TextMatrix(I, 4)
                 End If
                 RS!unit = vs.TextMatrix(I, 5)
                 RS!rate = Val(vs.TextMatrix(I, 6))
                 RS!amount = vs.TextMatrix(I, 7)
                 RS!billno = vs.TextMatrix(I, 8)
                 RS!discount = Val(vs.TextMatrix(I, 9))
                 RS!pending = vs.TextMatrix(I, 10)
                 RS!gamount = IIf(vs.TextMatrix(I, 11) = "", 0, vs.TextMatrix(I, 11))
                 
                 gaddi = 0
                 gaddi = noofgaddi(Val((IIf(vs.TextMatrix(I, 3) = "", 0, vs.TextMatrix(I, 3))) + Val(IIf(vs.TextMatrix(I, 4) = "", 0, vs.TextMatrix(I, 4)))), vs.TextMatrix(I, 1))
                 RS!noofgaddi = gaddi
                 noofgaddi_ = noofgaddi_ + gaddi


                 bagNo = 0
                 bagNo = noofbox(Val((IIf(vs.TextMatrix(I, 3) = "", 0, vs.TextMatrix(I, 3))) + Val(IIf(vs.TextMatrix(I, 4) = "", 0, vs.TextMatrix(I, 4)))), vs.TextMatrix(I, 1))
                 RS!noofbox = bagNo
               
                 bagInBox = bagInBox + bagNo


'                 bagNo = 0
'                 bagNo = noofbox(Val((IIf(vs.TextMatrix(I, 3) = "", 0, vs.TextMatrix(I, 3))) + Val(IIf(vs.TextMatrix(I, 4) = "", 0, vs.TextMatrix(I, 4)))), vs.TextMatrix(I, 1))
'                 RS!noofbox = bagNo
'
'                 bagInBox = bagInBox + bagNo



                 RS.update
                 
                 
                 
                 
                 
                 
            End If
   Next

    
    
    txtNoOfGaddi = Round(noofgaddi_, 2)
    txtbagInbox.text = Round(bagInBox, 2)
    
    
    
    Total
    
    
    con.Execute "UPDATE a SET a.pan = b.pan,a.noofgaddi=" & noofgaddi_ & ",NETAMOUNT=" & txtTotal.text & ",GAmount=" & txtGAmt.text & "  FROM ORDERA AS a " & _
    " INNER JOIN SLEDGER AS b ON (a.partyname = b.subledger " & _
    " and  a.invoiceno = " & txtOrderNo & ")"
    
    
    txtbagInbox.text = IIf(txtbagInbox.text = "", 0, txtbagInbox.text)
    
    con.Execute "UPDATE ORDERA set noofgaddi=" & noofgaddi_ & ",BagIn_Box=" & txtbagInbox.text & " where  invoiceno = " & txtOrderNo & ""
    
    

    Set rs1 = New ADODB.Recordset
    
''old code
'''        If txtScId.text = "" Then
'''            rs1.Open "select top 1 * from SeriesWiseDiscount where substring(Party,1,5)='" & Mid(txtPartyName.text, 1, 5) & "'", con
'''        Else
'''            rs1.Open "select top 1 * from SeriesWiseDiscountQry where substring(Party,1,5)='" & Mid(txtPartyName.text, 1, 5) & "' and scid='" & txtScId.text & "'", con
'''        End If
'''
    
    'PartyWiseDis_Con
    
    Set rs1 = New ADODB.Recordset
    
    If (DateDiff("d", Now, SessionLastDate) <= 0) Then
    
        If txtScId.text = "" Then
            rs1.Open "select top 1 * from SeriesWiseDiscount where substring(Party,1,5)='" & Mid(txtPartyName.text, 1, 5) & "'", con
        Else
            rs1.Open "select top 1 * from SeriesWiseDiscountQry where substring(Party,1,5)='" & Mid(txtPartyName.text, 1, 5) & "' and scid='" & txtScId.text & "'", con
        End If
    
    Else

        If txtScId.text = "" Then
            rs1.Open "select top 1 * from SeriesWiseDiscount where substring(Party,1,5)='" & Mid(txtPartyName.text, 1, 5) & "'", con_LAST
        Else
            rs1.Open "select top 1 * from SeriesWiseDiscountQry where substring(Party,1,5)='" & Mid(txtPartyName.text, 1, 5) & "' and scid='" & txtScId.text & "'", con_LAST
        End If

    End If
    
    
    
    
    
    
    
    If rs1.EOF = False Then
    
        seriesWiseMessage
        seriesWiseDis_ = "" & seriesWiseDis_
        
        If Len(seriesWiseDis_) > 2 Then
           MsgBox seriesWiseDis_, vbInformation
        End If
        
    End If
        
    
   
    
    cmd_update.Enabled = False


End Sub

Private Sub cmdAdd_1_Click()
   
   lblCAF(1).Caption = ""
   'cmdSDis.Visible = False
   txtbagInbox.text = ""
   txtrem.text = ""
   cboDisType.ListIndex = 0
   
   add_ = True
   maxOrder
   Edit = False
   refreshFld
   sale_sp
    cmdSave_2.Enabled = True
    cmdDelete_3.Enabled = False
    cmdEdit_4.Enabled = False
    cmd_update.Enabled = False
    
    cbofrt.text = ""
    
    Check1_school.value = 0
    txtOrderNo.SetFocus
    frmPendingClear.Visible = False
    Timer1.Enabled = False
    lblSave.Caption = ""
    Check1_AddOrderNo.value = 0
    
    Label1(20).Visible = False
    Label1(21).Visible = False
    Timer1.Enabled = False
    lblQtySpBalance.Caption = ""
    
    
    'SetButton cmdEdit_4, cmdEdit_4, cmdSave_2, cmdDelete_3

End Sub
Sub refreshFld()
   lblOrder_reminder.Caption = ""
   
   lblgpSchool(27).Caption = ""
   cboDisType.ListIndex = 0
   cbobiltyrem.text = ""
   
   party_type = ""
   TXTBAL.text = 0
   txtNoOfGaddi.text = 0
   txtbagInbox.text = 0
   txtContectNo = ""
   txtPartyName = ""
   txtMark = ""
   txtPin = ""
   txtParty = ""
   txtPartyAdd1 = ""
   txtPartyAdd2 = ""
   cboOrderBy.ListIndex = 0
   txtGAmt.text = 0
   cboTrans = ""
   txtTransAdd = ""
   txtNarration = ""
   txtBankName = ""
   txtBankAdd = ""
   txtBookingStn = ""
   txtTotal = ""
   
   txtTotQty = 0
   
   txtSchool = ""
   txtScId = ""
   cmbAgentName.text = ""
   txtShip = ""
   txtDist = ""
   txtDist_ship = ""
   txtShipState = ""
   txtPartySt = ""
   
   Timer2.Enabled = False
   Label1(25).Visible = False
   Label1(26).Visible = False
   
   vs.Clear
   setVSWidth

End Sub
Sub maxOrder()
   
'If rs1.State = 1 Then rs1.close
'rs1.Open "select  max(INVOICENO) from ORDERA_Tmp", con, adOpenDynamic, adLockOptimistic
'If Not IsNull(rs1(0)) Then
'   txtOrderNo = rs1(0) + 1
'   con.Execute "insert into ORDERA_Tmp(INVOICENO,SaveInMainTbl) values('" & txtOrderNo & "','n')"
'Else

    If RS.State = 1 Then RS.close
    RS.Open "select  max(INVOICENO) from ordera", con, adOpenKeyset, adLockOptimistic
    If IsNull(RS(0)) Then
        txtOrderNo = 1
      Else
        txtOrderNo = RS(0) + 1
    End If
    
'    con.Execute "insert into ORDERA_Tmp(INVOICENO,SaveInMainTbl) values('" & txtOrderNo & "','n')"
'End If
   
End Sub

Private Sub cmdAddFile_Click()
cd.ShowOpen
txtpath.text = cd.filename
End Sub

Private Sub cmdDelete_3_Click()

If txtOrderNo = "" Then Exit Sub
If RS.State = 1 Then RS.close
RS.Open "select OrderNo from INVOICEA where OrderNo=" & txtOrderNo & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
   MsgBox "Bill Already Created, you ca'nt delete ", vbCritical
Else
  
If MsgBox("Are you Sure ", vbQuestion + vbYesNo) = vbYes Then
   createLog UserName, txtOrderNo.text, "SaleOrder", " Delete : Qty " & txtTotQty.text, Date
   con.Execute "delete from ORDERA where INVOICENO=" & txtOrderNo & ""
   con.Execute "delete from ORDERB where INVOICENO=" & txtOrderNo & ""
   Check1_AddOrderNo.value = 0
End If

End If

End Sub

Private Sub cmdDis_Click()
If txtPartyName.text <> "" Then
    PopUpValue6 = txtPartyName.text
    
   If (LCase(UserName) = "admin") Then
    
        frmSeriesWiseDis.cmdSave_2.Enabled = True
        frmSeriesWiseDis.cmdDelete_3.Enabled = True
        frmSeriesWiseDis.cmdAdd_1.Enabled = True
        frmSeriesWiseDis.cmdEdit_4.Enabled = True
   
    End If
    
    frmSeriesWiseDis.Show 1
Else
   MsgBox "Plz Search Party ..", vbInformation
End If

End Sub

Private Sub cmdEdit_4_Click()
Edit = True
cmdSave_2.Enabled = False
'cmdDelete_3.Enabled = True
'cmdEdit_4.Enabled = False
vs.Enabled = True

mnuMenu_ = "mnuSaleOrder"
SetButton cmdEdit_4, cmd_update, cmdEdit_4, cmdDelete_3
If cmd_update.Enabled = True Then
   cmd_update.SetFocus
End If

End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdExit_Click()
frmPendingClear.Visible = False
End Sub
Private Sub cmdHide_Click()


If MsgBox("Want to Remove From List .. ?", vbQuestion + vbYesNo) = vbYes Then
    
    For J = 1 To vs_Mail.rows - 1
    If vs_Mail.TextMatrix(J, 5) = "-1" Then
     
     If vs_Mail.TextMatrix(vs_Mail.RowSel, 0) <> "" Then
       con.Execute "update ordera set MailSended='yes' where INVOICENO ='" & vs_Mail.TextMatrix(J, 0) & "'"
     End If
     
    End If
    Next
    
    cmdSendMail_Click
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

If InStr(txtParty, "(EM)") > 0 Then
   party_type = "EM"
Else
   party_type = "BP"
End If

rs_em.Open "select groupcode,bookcode from books", con, adOpenDynamic, adLockOptimistic

Set rs_fatch = con.Execute("exec BookSearch_withPartydis")

If RS.State = 1 Then RS.close
RS.Open "SELECT * FROM [sheet1$]", sconn
While RS.EOF = False

  

 




v1 = IIf(IsNull(RS(2)), 0, RS(2))
v2 = IIf(IsNull(RS(3)), 0, RS(3))

 If (v1 > 0 Or v2 > 0) Then
    
    
'========================================
If Option1_sale.value = True Then
          
If InStr(txtParty, "(SUNDRY)") = 0 Then

    rs_em.MoveFirst
    rs_em.Find "bookcode='" & UCase(RS(0)) & "'"
    If rs_em.EOF = False Then
                 
    If (rs_em(0) = "EM" Or rs_em(0) = "SMD" Or rs_em(0) = "GI" Or rs_em(0) = "QB" Or rs_em(0) = "SMD1" Or rs_em(0) = "SMD2" Or rs_em(0) = "QT1" Or rs_em(0) = "QT2" Or rs_em(0) = "Q-CUET") Then
       gp_ = "EM"
    End If
                 
    If party_type = gp_ Then
        If party_type <> gp_ Then
             MsgBox "Book List is not valid for this customer ...", vbCritical
             Exit Sub
        End If
    Else
        If (rs_em(0) = "EM" Or rs_em(0) = "SMD" Or rs_em(0) = "GI" Or rs_em(0) = "QB" Or rs_em(0) = "SMD1" Or rs_em(0) = "SMD2" Or rs_em(0) = "QT1" Or rs_em(0) = "QT2" Or rs_em(0) = "Q-CUET") Then
             MsgBox "Book List is not valid for this customer ...", vbCritical
             Exit Sub
        End If
    End If
 End If

End If
End If
'=======================================
  
    
    vs.TextMatrix(k1, 0) = k1
    vs.TextMatrix(k1, 1) = UCase(RS(0))
    vs.TextMatrix(k1, 2) = UCase(RS(1))
    If Option1_sale.value = False Then
       vs.TextMatrix(k1, 4) = v2
       txtTotQty = Val(txtTotQty) + v2
    Else
       
       vs.TextMatrix(k1, 3) = v1
       vs.TextMatrix(k1, 4) = v2
       txtTotQty = Val(txtTotQty) + v1
       txtBillQty = Val(txtBillQty) + v2
    End If
    '------------------------------------------------------------
      If rs1.State = 1 Then rs1.close
      rs1.Open "select Bookcode,bookname,rate,DISCOUNT from BOOKS where bookcode='" & vs.TextMatrix(k1, 1) & "'", con
      If rs1.EOF = False Then
         vs.TextMatrix(k1, 6) = rs1!rate
         
         qty = 0
         qty = IIf(vs.TextMatrix(k1, 3) = "", 0, vs.TextMatrix(k1, 3))
         
         If Option1_sale.value = False Then
            qty = Val(qty) + Val(IIf(vs.TextMatrix(k1, 4) = "", 0, vs.TextMatrix(k1, 4)))
         End If

         
         net = (qty * Val(vs.TextMatrix(k1, 6)))
         
         vs.TextMatrix(k1, 9) = rs1!discount
         
         dis = rs1!discount
         If Option1_sale.value = True Then
            rs_fatch.MoveFirst
            rs_fatch.Find "bookcode='" & RS(0) & "'"
            If rs_fatch.EOF = False Then
               vs.TextMatrix(k1, 9) = rs_fatch!discount
               dis = rs_fatch!discount
            End If
         End If
         
         vs.TextMatrix(k1, 7) = net - Format(Round(net * (dis / 100), 2), "0.00")
         
         vs.TextMatrix(k1, 11) = net
      End If
          
    '------------------------------------------------------------


    k1 = k1 + 1
 End If


'abc:


RS.MoveNext
Wend

Total

MsgBox "Data import Successfully", vbInformation

End Sub
Private Sub cmdListBlankOrd_Click()

List_emptyList.Visible = True
Dim k1 As Integer
Dim b1 As Boolean
b1 = False


If emptyInv_bool = False Then
   List_emptyList.Visible = True
   emptyInv_bool = True
Else
   List_emptyList.Visible = False
   emptyInv_bool = False
End If


List_emptyList.Clear


Set RS = New ADODB.Recordset
Set RS = con.Execute("exec searchList 'ORDER'")

While RS.EOF = False

    If b1 = False Then
       k1 = RS(0)
       b1 = True
    End If
    
    If k1 <> RS(0) Then
       List_emptyList.AddItem (RS(0) - 1)
       b1 = False
    End If

k1 = k1 + 1
RS.MoveNext

Wend


End Sub
Private Sub cmdMail_Click()


Dim address1 As String
Dim address2 As String
Dim address3 As String
Dim address4 As String
Dim transport As String
Dim through As String
Dim party_mail As String
Dim rep_mail As String
Dim head_mail As String
Dim head_mail2 As String
Dim rss As New ADODB.Recordset



If MsgBox("Want to Send Mail .. ?", vbQuestion + vbYesNo) = vbYes Then

For J = 1 To vs_Mail.rows - 1
 
If (vs_Mail.TextMatrix(J, 0) <> "" And vs_Mail.TextMatrix(J, 5) = "-1") Then
  
 If RS.State = 1 Then RS.close
 
    RS.Open "select INVOICENO,INVOICEDATE,SUBLEDGER,Address1,Address2,transport,BookingStn,party_state,ORDERBY,Sale_sp,ContactNo  from ORDERA where invoiceno=" & vs_Mail.TextMatrix(J, 0) & "", con
    If RS.EOF = False Then
       address1 = RS!subledger
       address2 = RS!address1
       address3 = RS!address2
       address4 = RS!BookingStn + ",(" + RS!party_state + ")"
       transport = RS!transport & ""
       through = RS!orderby & ""
       

 
       party_mail = ""
       rep_mail = ""
       head_mail = ""
       head_mail2 = ""
       

       If rss.State = 1 Then rss.close
       rss.Open "select Email,RepName1,RepName2,mobile from SLEDGER  where SUBLEDGER='" & vs_Mail.TextMatrix(J, 2) & "'", con
       If rss.EOF = False Then
             
'          If (Not IsNull(rss!mobile) Or rss!mobile <> "") Then
'             mobile_ = ""
'
'             mobile_ = Mid(rss!mobile, 1, 10)
'             mobile_ = "9997314681"
'
'             con.Execute "insert into whatsapp_SMS(PartyName,Mobile,types) values('" & address1 & "','" & mobile_ & "','ORD')"
'          End If

             
             
             party_mail = rss!email & ""
             
             If rss!RepName2 <> "" Then
             If rs1.State = 1 Then rs1.close
             rs1.Open "select email,HeadEmail,HeadEmail_2 from Rep where rep='" & rss!RepName1 & "'", CON_blue
             If rs1.EOF = False Then
                    
                    head_mail = rs1!HeadEmail & ""     'head
                    rep_mail = rs1!email & ""          'Rep. Mail
                    
                    If (head_mail = rep_mail) Then
                        rep_mail = ""
                    End If
                    
                    If Not IsNull(HeadEmail2) Then
                       head_mail2 = rs1!HeadEmail_2 & ""
                    End If
                    
                    
             End If
             End If
             
             
             
             
             
             
          End If
       
       End If


    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from MailDetails where (Bill=" & vs_Mail.TextMatrix(J, 0) & " and BillType='order')", con
    If rs1.EOF = True Then
    
    If party_mail <> "" Then
       party_ = Trim(Mid(vs_Mail.TextMatrix(J, 2), 6))
       con.Execute "insert into MailDetails(Bill,BillType,Mail,MailSended,OrderDis,address1,address2,address3,address4,through,Transport,RepEmail,HeadEmail,HeadEmail_2)" & _
       " values(" & vs_Mail.TextMatrix(J, 0) & ",'order','" & party_mail & "','n','y','" & party_ & "','" & address2 & "','" & address3 & "','" & address4 & "','" & through & "','" & transport & "','" & rep_mail & "','" & head_mail & "','" & head_mail2 & "')"
    End If
       
    Else
       con.Execute "update MailDetails set MailSended='n',OrderDis='y',Mail='" & party_mail & "',RepEmail='" & rep_mail & "',HeadEmail='" & head_mail & "',address1='" & address1 & "',address2='" & address2 & "',address3='" & address3 & "',address4='" & address4 & "',HeadEmail_2='" & head_mail & "' where (Bill=" & vs_Mail.TextMatrix(J, 0) & " and BillType='order')"
    End If
    

    con.Execute "update ordera set MailSended='yes' where INVOICENO ='" & vs_Mail.TextMatrix(J, 0) & "'"

End If


Next

End If


End Sub

Private Sub cmdNext_Click()

'
'frmBookStock.Show
'
'Exit Sub

'On Error GoTo err1
    
''   txtRem.text = ""
''
''   lblSave.Caption = ""
''   If RS.State = 1 Then RS.close
''   RS.Open "select * from ordera where invoiceno=" & txtOrderNo & "", con, adOpenDynamic, adLockOptimistic
''   If RS.EOF = False Then
''
''    add_ = False
''    vs.Enabled = True
''    cmdSave_2.Enabled = False
''    cmdDelete_3.Enabled = False
''    cmdEdit_4.Enabled = True
''    cmd_update.Enabled = False
''
''    cbofrt.text = RS!Frt_Yes & ""
''
''    txtPIN_Ship.text = RS!pin_ship & ""
''
''     lblgpSchool(27).Caption = RS!GroupOfSchool & ""
''
''    cbobiltyrem.text = RS!ccattach & ""
''    cboBilty.text = RS!bilty & ""
''    If RS!sale_sp = "sale" Then
''       Option1_sale.value = True
''     Else
''       Option2_sp.value = True
''    End If
''
''
''    If Not IsNull(RS!bal) Then
''       TXTBAL.text = RS!bal
''    Else
''       TXTBAL.text = 0
''    End If
''    txtMark = RS!Godown & ""
''    txtPartyName.text = RS!Godown & ""
''    txtContectNo = RS!ContactNo & ""
''    txtShip = RS!Shipto & ""
''    txtPartyName.text = RS!partyname & ""
''    txtschool.text = RS!scname & ""
''    txtScId.text = RS!scid & ""
''    cmbAgentName.text = RS!RepName & ""
''    txtOrderNo = RS!invoiceNo
''    txtOrderDate = RS!invoiceDate
''    txtParty = RS!subledger
''    txtPartyAdd1 = RS!address1 & ""
''    txtPartyAdd2 = RS!address2 & ""
''    cboOrderBy = RS!orderby & ""
''    txtOrderDate1.value = RS!ORDERDATE
''    cboTrans = RS!transport & ""
''    txtTransAdd = RS!TransAdd & ""
''    txtNarration = RS!narration & ""
''    txtBankName = RS!Shipto_Add1 & ""
''    txtBankAdd = RS!Shipto_Add2 & ""
''    txtDist_ship = RS!shipto_dist & ""
''    txtShipState.text = RS!Shipto_States & ""
''    txtBookingStn = RS!BookingStn & ""
''    txtTotal = RS!netamount
''    txtGAmt.text = RS!gamount & ""
''
''    txtDist = RS!party_dist & ""
''    txtDist_ship = RS!shipto_dist & ""
''    txtPartySt = RS!party_state & ""
''    txtPin = ""
''    If Option2_sp.value = True Then
''        txtPin = RS!pin & ""
''    Else
''
''        txtRem.text = ""
''        If rs1.State = 1 Then rs1.close
''        rs1.Open "select pin,PartyRemarks from sledger where SUBLEDGER ='" & txtPartyName & "'", con
''        If rs1.EOF = False Then
''           txtPin = rs1!pin
''           txtRem.text = rs1!PartyRemarks & ""
''        End If
''
''    End If
''
''
''
''
''
''   End If
''
   
   
   
Dim Q1 As Integer
Dim inv As String
Dim noofgaddi As Double



noofgaddi = 0
   
   
Dim rss As New ADODB.Recordset
   
   
If rss.State = 1 Then rss.close
rss.Open "select invoiceNo,sale_sp from ordera where invoiceNo>=13000 order by invoiceNo", con, adOpenDynamic, adLockOptimistic
While rss.EOF = False
   
txtOrderNo.text = rss!invoiceNo
   
If rss!sale_sp = "sale" Then
   Option1_sale.value = True
 Else
   Option2_sp.value = True
End If
   
   
str1 = "SELECT  PRINTORDER,ORDERB.BOOKCODE, BOOKNAME, ORDERB.RATE,ORDERB.discount,ORDERB.billno, QUANTITY,unit,AMOUNT,orderb.pending,orderb.spQty,orderb.onlineAmt,orderb.gamount,orderb.noofgaddi,orderb.sno " & _
"FROM ORDERB INNER JOIN BOOKS ON ORDERB.BOOKCODE = BOOKS.BOOKCODE where invoiceno=" & txtOrderNo & " order by printorder"
If RS.State = 1 Then RS.close
RS.Open str1, con, adOpenForwardOnly, adLockReadOnly
For I = 1 To RS.RecordCount


       inv = ""
       Q1 = 0
       
       
       If Option1_sale.value = True Then
       
       'If rs1.State = 1 Then rs1.close
       'rs1.Open "select sum(QUANTITY),invoiceno from InvoicebAndInvoicebSP_Qry where OrderNo=" & txtOrderNo & " and bookcode='" & RS!Bookcode & "' group by invoiceno", con
       Set rs1 = New ADODB.Recordset
       Set rs1 = con.Execute("exec SP_PendingOrder_saleAndSp '" & txtOrderNo & "','" & RS!Bookcode & "'")
       While rs1.EOF = False
           Q1 = Q1 + rs1(0)
           If inv = "" Then
              inv = rs1!invoiceNo
             Else
              inv = inv & "," & rs1!invoiceNo
           End If
       rs1.MoveNext
       Wend
       
       'If rs1.EOF = False Then
       If Q1 > 0 Then
            'vs.TextMatrix(I, 5) = Q1
            'vs.TextMatrix(I, 8) = inv     'rs1!invoiceno
            'vv = (IIf(vs.TextMatrix(I, 3) = "", 0, Val(vs.TextMatrix(I, 3))) + IIf(vs.TextMatrix(I, 4) = "", 0, Val(vs.TextMatrix(I, 4))))
            'vv1 = IIf(vs.TextMatrix(I, 5) = "", 0, Val(vs.TextMatrix(I, 5)))
            'If vv = vv1 Then
            '    For k1 = 0 To 11
            '        vs.Cell(flexcpBackColor, I, k1) = vbGreen
            '        DoEvents
            '    Next
            '    vs.TextMatrix(I, 10) = "n"
            'End If
        
        Else
        
            con.Execute "insert into  ORDERA_ForNext SELECT * FROM ORDERA where invoiceno=" & txtOrderNo & ""
            con.Execute "insert into  ORDERB_ForNext([INVOICENO],[INVOICEDATE],[BOOKCODE],[QUANTITY],[Unit],[RATE],[AMOUNT],[PRINTORDER],[Fyear],[setupid],[Bquantity] ,[DISCOUNT],[billNo] ,[Pending],[SpQty],[PQty],[onlineAmt],[onlineDis],[onlineDisAmt],[gamount] ,[noofgaddi]) SELECT [INVOICENO],[INVOICEDATE],[BOOKCODE],[QUANTITY],[Unit],[RATE],[AMOUNT],[PRINTORDER],[Fyear],[setupid],[Bquantity] ,[DISCOUNT],[billNo] ,[Pending],[SpQty],[PQty],[onlineAmt],[onlineDis],[onlineDisAmt],[gamount] ,[noofgaddi] FROM ORDERB where invoiceno=" & txtOrderNo & " and BOOKCODE='" & RS!Bookcode & "' and sno=" & RS!sno & ""

        
        End If
        
Else
     
       inv = ""
       Q1 = 0
     
       If rs1.State = 1 Then rs1.close
       rs1.Open "select sum(QUANTITY),invoiceno from invoicespBQry where OrderNo=" & txtOrderNo & " and bookcode='" & RS!Bookcode & "' group by invoiceno", con
       While rs1.EOF = False
           Q1 = Q1 + rs1(0)
           If inv = "" Then
              inv = rs1!invoiceNo
             Else
              inv = inv & "," & rs1!invoiceNo
           End If
       rs1.MoveNext
       Wend
     
     
         If Q1 > 0 Then
            'vs.TextMatrix(I, 5) = Q1   'rs1(0)
            'vs.TextMatrix(I, 8) = inv     'rs1!invoiceno
            'If vs.TextMatrix(I, 4) = vs.TextMatrix(I, 5) Then
            '    For k1 = 0 To 11
                    'vs.Cell(flexcpBackColor, I, k1) = vbGreen
            '        DoEvents
            '    Next
            '    vs.TextMatrix(I, 10) = "n"
            'End If
        Else
           
        On Error Resume Next
        con.Execute "insert into  ORDERA_ForNext SELECT * FROM ORDERA where invoiceno=" & txtOrderNo & ""
        
        
        'con.Execute "insert into  ORDERB_ForNext SELECT * FROM ORDERB where invoiceno=" & txtOrderNo & " where BOOKCODE='" & RS!Bookcode & "' and sno=" & RS!sno & ""

        con.Execute "insert into  ORDERB_ForNext([INVOICENO],[INVOICEDATE],[BOOKCODE],[QUANTITY],[Unit],[RATE],[AMOUNT],[PRINTORDER],[Fyear],[setupid],[Bquantity] ,[DISCOUNT],[billNo] ,[Pending],[SpQty],[PQty],[onlineAmt],[onlineDis],[onlineDisAmt],[gamount] ,[noofgaddi]) SELECT [INVOICENO],[INVOICEDATE],[BOOKCODE],[QUANTITY],[Unit],[RATE],[AMOUNT],[PRINTORDER],[Fyear],[setupid],[Bquantity] ,[DISCOUNT],[billNo] ,[Pending],[SpQty],[PQty],[onlineAmt],[onlineDis],[onlineDisAmt],[gamount] ,[noofgaddi] FROM ORDERB where invoiceno=" & txtOrderNo & " and BOOKCODE='" & RS!Bookcode & "' and sno=" & RS!sno & ""

          
       
       End If
       End If
       
       
       
       
 
       RS.MoveNext
Next
   

rss.MoveNext

Wend


'
'Exit Sub
'
'err1:
'
'MsgBox "" & err.DESCRIPTION


End Sub

Private Sub cmdOnlinePrint_Click()
If (session = "2015-16" Or session = "2016-17" Or session = "2017-18") Then
Else
  frmOnlineBill.Show 1
End If

End Sub

Private Sub cmdPending_Click()

On Error GoTo err1:

Dim s1 As String
Dim ii As Long

'login.DSN
DSNNew


If Check1_notefull.value = 0 Then

    If cbogd1.text = "" Then
     con.Execute "exec sale_pendingSpList '" & "ALL" & "'"
    Else
     con.Execute "exec sale_pendingSpList '" & cbogd1.text & "'"
    End If
Else

    If cbogd1.text = "" Then
     con.Execute "exec sale_pendingSpListFinal '" & "ALL" & "'"
    Else
     con.Execute "exec sale_pendingSpListFinal '" & cbogd1.text & "'"
    End If

End If

'----------------------------------------------------------------
If MsgBox("Want to Print?", vbQuestion + vbYesNo) = vbYes Then

    
    CR.Reset
    CR.ReportFileName = rptPath & "/PendingorderListSale.rpt"
    CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    CR.Formulas(0) = "header='Sales Pending Order'"
    CR.WindowShowPrintSetupBtn = True
    CR.WindowShowRefreshBtn = True
    CR.WindowMaxButton = True
    CR.WindowShowSearchBtn = True
    CR.WindowState = crptMaximized
    CR.Action = 1

End If


Exit Sub

err1:

MsgBox "" & err.DESCRIPTION

End Sub
Private Sub cmdpendingbook_Click()


On Error GoTo err1:

Dim s1 As String
Dim ii As Long
Dim gp As String

DSNNew



If (Text1.text <> "") Then
  
    If RS.State = 1 Then RS.close
    RS.Open "select groupcode from books where sername='" & Text1.text & "'", con
    If RS.EOF = False Then
        gp = RS(0)
    End If

Else
    
    If RS.State = 1 Then RS.close
    RS.Open "select groupcode from books where bookcode='" & txtbk.text & "'", con
    If RS.EOF = False Then
        gp = RS(0)
    End If

End If


'con.Execute "exec sale_pendingSpList"

If cbogd1.text = "" Then
  con.Execute ("exec PendingOrder_Sale_Specimen_New1 '" & cbogd1.text & "','" & gp & "'")
  
ElseIf Text1.text <> "" Then
  con.Execute ("exec PendingOrder_Sale_Specimen_New1 '" & cbogd1.text & "','" & gp & "'")
Else
  con.Execute ("exec PendingOrder_Sale_Specimen_New1 '" & cbogd1.text & "','" & "" & "'")
End If

'----------------------------------------------------------------
If MsgBox("Want to Print?", vbQuestion + vbYesNo) = vbYes Then

CR.Reset
If (Text1.text <> "" And txtbk.text = "") Then
    CR.ReportFileName = rptPath & "/PendingorderBkList _New.rpt"
    CR.ReplaceSelectionFormula "{books.sername}='" & Text1.text & "'"
ElseIf txtbk <> "" Then
    CR.ReportFileName = rptPath & "/PendingorderBkList _New.rpt"
    CR.ReplaceSelectionFormula "{pendingSplist.bookcode}='" & txtbk.text & "'"
Else
    CR.ReportFileName = rptPath & "/PendingorderBkList _New.rpt"
End If


CR.Connect = "filedsn=chitradsn;uid=" & sql_user & ";pwd=" & sql_pass
CR.Formulas(0) = "header='Book Pending List'"

CR.WindowShowPrintSetupBtn = True
CR.WindowShowRefreshBtn = True
CR.WindowMaxButton = True
CR.WindowState = crptMaximized
CR.Action = 1

End If


Exit Sub

err1:

MsgBox "" & err.DESCRIPTION



''On Error GoTo err1:
''
''Dim s1 As String
''Dim ii As Long
''
''DSNNew
''
''If Option1_bookwise.value = True Then
''
''        If cbogd1.text = "" Then
''         con.Execute "exec sale_pendingSpList '" & "ALL" & "'"
''        Else
''         con.Execute "exec sale_pendingSpList '" & cbogd1.text & "'"
''        End If
''
''        '----------------------------------------------------------------
''        If MsgBox("Want to Print?", vbQuestion + vbYesNo) = vbYes Then
''
''        CR.Reset
''        If txtbk <> "" Then
''            CR.ReportFileName = rptPath & "/PendingorderBkList.rpt"
''            CR.ReplaceSelectionFormula "{pendingSpList.bookcode}='" & txtbk.text & "' and {pendingSpList.balance}<>0"
''        End If
''
''
''        CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
''        CR.Formulas(0) = "header='Book Pending List'"
''
''        CR.WindowShowPrintSetupBtn = True
''        CR.WindowShowRefreshBtn = True
''        CR.WindowMaxButton = True
''        CR.WindowState = crptMaximized
''        CR.Action = 1
''
''        End If
''
''Else
''
''        If (Check1_onlySp.value = 0) Then
''
''            If cbogd1.text = "" Then
''             con.Execute "exec SerWise_sale_pendingSpList '" & "ALL" & "'"
''            Else
''             con.Execute "exec SerWise_sale_pendingSpList '" & cbogd1.text & "','" & Text1.text & "'"
''            End If
''
''        Else
''
''            If cbogd1.text = "" Then
''             con.Execute "exec SerWise_sale_pendingList_ForSp '" & "ALL" & "'"
''            Else
''             con.Execute "exec SerWise_sale_pendingList_ForSp '" & cbogd1.text & "','" & Text1.text & "'"
''            End If
''
''         End If
''
''
''
''        '----------------------------------------------------------------
''        If MsgBox("Want to Print?", vbQuestion + vbYesNo) = vbYes Then
''
''        CR.Reset
''        If txtbk <> "" Then
''            CR.ReportFileName = rptPath & "/PendingorderBkSeriseWise.rpt"
''            CR.ReplaceSelectionFormula "{PendingSpList.totOrdQty}-{PendingSpList.totBillQty}>0 and {PendingSpList.bookcode}='" & txtbk.text & "'"
''        Else
''            CR.ReportFileName = rptPath & "/PendingorderBkSeriseWise.rpt"
''            CR.ReplaceSelectionFormula "{PendingSpList.totOrdQty}-{PendingSpList.totBillQty}>0"
''
''        End If
''
''
''        CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
''        CR.Formulas(0) = "header='Book Pending List'"
''
''        CR.WindowShowPrintSetupBtn = True
''        CR.WindowShowRefreshBtn = True
''        CR.WindowMaxButton = True
''        CR.WindowState = crptMaximized
''        CR.Action = 1
''
''        End If
''
''
''
''
''End If
''
''
''
''Exit Sub
''
''err1:
''
''MsgBox "" & err.DESCRIPTION
''



End Sub

Private Sub cmdPendingSpOrder_Click()

On Error GoTo err1:

Dim s1 As String
Dim ii As Long

s1 = ""
ii = 0

DSNNew


If Check1_notefull.value = 0 Then

    If cbogd1.text = "" Then
    con.Execute "exec SP_pendingSpList '" & "ALL" & "'"
    Else
    con.Execute "exec SP_pendingSpList '" & cbogd1.text & "'"
    End If
Else

  If cbogd1.text = "" Then
    con.Execute "exec SP_pendingSpListfinal '" & "ALL" & "'"
    Else
    con.Execute "exec SP_pendingSpListfinal '" & cbogd1.text & "'"
    End If

End If


''//con.Execute "exec SP_pendingSpList"
'----------------------------------------------------------------
If MsgBox("Want to Print?", vbQuestion + vbYesNo) = vbYes Then
    CR.Reset
    CR.ReportFileName = rptPath & "/PendingorderListNew.rpt"
    CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    CR.ReplaceSelectionFormula "({pendingSpList.qty}-{pendingSpList.bqty})>0"
    
    
    CR.Formulas(0) = "header='Specimen Pending Order'"
    CR.WindowShowPrintSetupBtn = True
    CR.WindowShowRefreshBtn = True
    CR.WindowMaxButton = True
    CR.WindowState = crptMaximized
    CR.Action = 1
End If


Exit Sub
err1:
MsgBox "" & err.DESCRIPTION

End Sub
Sub printPRemarks()

Dim f As New ADODB.Recordset
Set f = New ADODB.Recordset
st10 = ""
Dim ss_, ss1_, pp_ As String
Dim tdis As Double
ss_ = ""
ss1_ = ""
pp_ = ""
tdis = 0

con.Execute "delete from AppPrintTmp1"


If (txtScId.text <> "" And txtPartyName <> "") Then
   str1 = "SELECT SUBLEDGER,PartyRemarks,appNo,Promo,Net_Gross,PartyRemarks,adj,discount,appPer,tod,cd FROM PartyRemarksQrynew where (scid='" & txtScId.text & "' and SUBLEDGER='" & txtPartyName.text & "') group by SUBLEDGER,PartyRemarks,appNo,Promo,Adj,Net_Gross,PartyRemarks,adj,discount,appPer,tod,cd"
Else
   If txtPartyName.text <> "" Then
      con.Execute "insert into AppPrintTmp1(Party,Remarks) select subledger,partyremarks from sledger where subledger='" & txtPartyName.text & "'"
   End If
   Exit Sub
End If


ss_ = ""
If f.State = 1 Then f.close

If (DateValue(txtOrderDate.value) >= DateValue(financialyear_Fdate) And DateValue(txtOrderDate.value) <= DateValue(financialyear_Tdate)) Then
   f.Open str1, con
Else
   f.Open str1, con_LAST
End If

If f.RecordCount = 0 Then
   If txtPartyName.text <> "" Then
      con.Execute "insert into AppPrintTmp1(Party,Remarks) select subledger,partyremarks from sledger where subledger='" & txtPartyName.text & "'"
   End If
End If


While f.EOF = False
      
       ss_ = ""
       remarks1 = ""
       If rs1.State = 1 Then rs1.close
       If (DateValue(txtOrderDate.value) >= DateValue(financialyear_Fdate) And DateValue(txtOrderDate.value) <= DateValue(financialyear_Tdate)) Then
          rs1.Open "select sername,Promo,PartyRemarks as remarks,Adj,discount,appPer from PartyRemarksQryNew where (appNo=" & f(2) & " and Promo=" & f!Promo & " and adj=" & f.Fields("adj").value & " and discount=" & f.Fields("discount").value & " and appPer=" & f.Fields("appPer").value & ") group by sername,Promo,PartyRemarks,Adj,discount,appPer", con
       Else
         rs1.Open "select sername,Promo,PartyRemarks as remarks,Adj,discount,appPer from PartyRemarksQryNew where (appNo=" & f(2) & " and Promo=" & f!Promo & " and adj=" & f.Fields("adj").value & " and discount=" & f.Fields("discount").value & " and appPer=" & f.Fields("appPer").value & ") group by sername,Promo,PartyRemarks,Adj,discount,appPer", con_LAST
       End If
       
       If rs1.EOF = False Then
          remarks1 = rs1!remarks & ""
       End If
       
       While rs1.EOF = False
          If ss_ = "" Then
           ss_ = rs1(0)
          Else
           ss_ = ss_ & "," & rs1(0)
          End If
          rs1.MoveNext
       Wend
     
      If (f(2) > 0) Then
         tdis = (f.Fields("appPer").value + f.Fields("adj").value + f.Fields("discount").value + f.Fields("Promo").value)
         con.Execute "insert into AppPrintTmp1(party,Remarks,appno,PromPer,adjper,sername,remarks1,gross_,upto_5,TDis,tod,cd) values('" & f(0) & "','" & f(5) & "','" & f(2) & "','" & f(3) & "','" & f.Fields("adj").value & "','" & ss_ & "','" & f.Fields("discount").value & "','" & f(4) & "','" & f.Fields("appper").value & "','" & tdis & "','" & f.Fields("tod").value & "','" & f.Fields("cd").value & "')"
      End If
     
     
     
     ss_ = ""

       
    f.MoveNext
Wend


If RS.State = 1 Then RS.close
RS.Open "SELECT SUBLEDGER,PartyRemarks,appNo FROM invoiceaQry where (scid='" & txtScId.text & "' and SUBLEDGER='" & txtPartyName.text & "') group by SUBLEDGER,PartyRemarks,appNo", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   If rs1.State = 1 Then rs1.close
   rs1.Open "SELECT party,Remarks,appNo FROM AppPrintTmp1 where party='" & RS!subledger & "'", con, adOpenDynamic, adLockOptimistic
   If rs1.RecordCount > 0 Then
      con.Execute "delete from AppPrintTmp1 where party='" & rs1!party & "' and (AppNo ='' or AppNo =0) and (remarks='NA' or remarks='')"
   End If
   RS.MoveNext
Wend

End Sub
Sub PartyLedgerNew(pname As String)

dates = to_date
user_id = Trim((Sys_user_ + Str(UId)))
con.Execute "delete from templedger6 where userid='" & user_id & "'"

con.Execute "INSERT INTO templedger6 (Balance,drcr,party,billtype,rptid,rptype,setupid,fyear,district,userid,states,Party1)  SELECT op,drcr,subledger,'Opening',1,'',setupid,fyear,ADDRESS3,'" & user_id & "',states,DESCFORINVOICE from sledger where subledger='" & pname & "'  group by op,subledger,drcr,setupid,Fyear,ADDRESS3,states,DESCFORINVOICE  HAVING  op <> 0"
con.Execute "INSERT INTO templedger6 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,scname,billtype1,sdiscount,todno,toddate,scid)  SELECT INVOICEDATE,'I',INVOICENO,'Invoice Sales Bilty No-' + BILTYNO + ',Bundle-' + bundles ,netamount,BAA,SUBLEDGER,fyear,setupid,'" & user_id & "',district,'1','',states,Party,'',scname,'I',sdiscount,todid,toddate,scid  from invoiceaQry where Subledger='" & pname & "' and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & dates & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger6 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,scname,billtype1,todno,toddate,scid) Select INVOICEDATE,'CI',INVOICENO,'Credit Note Item Bilty No-' + BILTYNO + ',Bundle-' + bundles,BAA,netamount,SUBLEDGER,fyear,setupid,'" & user_id & "',district,'1','',states,Party,'',scname,'C',todid,toddate,scid from CREDITAQry where Subledger='" & pname & "' and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & dates & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"
con.Execute "INSERT INTO templedger6 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,sdiscount)" & _
"SELECT  CASHA.INVOICEDATE,'C/M',CASHA.INVOICENO,'Cash Memo',CASHA.NETAMOUNT,CASHA.BAA,CASHA.cashpartyname,CASHA.Fyear," & _
"CASHA.setupid,'" & user_id & "',SLEDGER.ADDRESS3,'1','',SLEDGER.states,SLEDGER.DESCFORINVOICE,'',casha.sdiscount " & _
"FROM CASHA INNER JOIN SLEDGER ON CASHA.SUBLEDGER = SLEDGER.SUBLEDGER where CASHA.SUBLEDGER='" & pname & "' and convert(smalldatetime,INVOICEDATE,103)<=convert(smalldatetime,'" & dates & "',103) and ((netamount-BAA)>0 or (netamount-BAA)<0)"

'-
con.Execute "INSERT INTO templedger6 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,Billtype1,todno,toddate)" & _
" SELECT CNF1A.CND,'CN',CNF1A.cnn,'Credit Note ' + desc_ ,0,CNF1A.NA,CNF1A.psld,CNF1A.Fyear," & _
"CNF1A.setupid,'" & user_id & "',SLEDGER.ADDRESS3,'1','',SLEDGER.states,SLEDGER.DESCFORINVOICE,'','CN',todid,toddate " & _
"FROM  dbo.CNF1A INNER JOIN SLEDGER ON CNF1A.psld = dbo.SLEDGER.SUBLEDGER where  SLEDGER.SUBLEDGER ='" & pname & "' and convert(smalldatetime,cnd,103)<=convert(smalldatetime,'" & dates & "',103) "

con.Execute "INSERT INTO templedger6 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName,Billtype1)" & _
"SELECT DNFA.DND,'DN',DNFA.Dnn,'Debit Note' + desc_,DNFA.NA,0,DNFA.psld,DNFA.Fyear," & _
"DNFA.setupid,'" & user_id & "',SLEDGER.ADDRESS3,'1','',SLEDGER.states,SLEDGER.DESCFORINVOICE,'','DN' " & _
"FROM DNFA INNER JOIN SLEDGER ON DNFA.psld = SLEDGER.SUBLEDGER where SLEDGER.SUBLEDGER ='" & pname & "' and convert(smalldatetime,dnd,103)<=convert(smalldatetime,'" & dates & "',103)"
'-
con.Execute "INSERT INTO templedger6 (dates,Billtype,bill,des,dr,cr,Party,fyear,setupid,userid,district,rptid,rptype,states,Party1,RepName) " & _
" SELECT a.Dates,'J',a.RecNo, a.Particullar, a.Dr, a.Cr,a.PartyName,a.fyear,a.setupid,'" & user_id & "'," & _
" b.ADDRESS3,'1','',b.states,b.DESCFORINVOICE,'' FROM ReceiveIssueParty as a INNER JOIN " & _
" SLEDGER as b ON a.PartyName = b.SUBLEDGER where a.PartyName ='" & pname & "' and convert(smalldatetime,DATEs,103)<=convert(smalldatetime,'" & dates & "',103) order by dates,recno"

Dim bal
TXTBAL.text = 0
bal = 0
Set RS = New ADODB.Recordset

RS.Open "SELECT SUM(Dr) AS DR,SUM(Cr) AS CR FROM tempLedger6 WHERE (Party ='" & pname & "' and UserId='" & user_id & "')", con
If RS.EOF = False Then
   bal = RS(0) - RS(1)
   
   If IsNull(bal) Then
      bal = 0
   End If
End If


DoEvents
DoEvents
DoEvents

If IsNull(bal) Then
   bal = 0
End If

If RS.State = 1 Then RS.close
RS.Open "SELECT SUM(Balance),drcr FROM tempLedger6 WHERE (billtype ='Opening' and Party ='" & pname & "' and UserId='" & user_id & "') group by drcr", con
If RS.EOF = False Then
   If RS!drcr = "Dr" Then
      bal = bal + RS(0)
   Else
      bal = bal - RS(0)
   End If
End If

DoEvents
DoEvents
DoEvents

TXTBAL.text = Round(bal, 2) & ""

End Sub
Private Sub cmdPrint_Click()

'==========================================================================================




Screen.MousePointer = vbHourglass

Dim series_discount As String
series_discount = ""
series_discount1 = ""
series_discount2 = ""

pcode_ = Mid(txtPartyName.text, 1, 5)

kk = 1



Set rs1 = New ADODB.Recordset

'old code
'rs1.Open "select Id,SeriesName,DISCOUNT,ScName,GroupOfSchool,ScId from SeriesWiseDiscount where substring(Party,1,6)='" & pcode_ & "' order by SeriesName", con



If (DateDiff("d", Now, SessionLastDate) <= 0) Then
    rs1.Open "select Id,SeriesName,DISCOUNT,ScName,GroupOfSchool,ScId from SeriesWiseDiscount where substring(Party,1,6)='" & pcode_ & "' order by SeriesName", con
Else
    rs1.Open "select Id,SeriesName,DISCOUNT,ScName,GroupOfSchool,ScId from SeriesWiseDiscount where substring(Party,1,6)='" & pcode_ & "' order by SeriesName", con_LAST
End If






While rs1.EOF = False
   
   
  If kk <= 18 Then
   If series_discount = "" Then
      series_discount = rs1!SeriesName & "-" & rs1!discount
     Else
      series_discount = series_discount & "," & rs1!SeriesName & "-" & rs1!discount
   End If
  End If
  
   If (kk >= 19 Or kk >= 33) Then
   If series_discount1 = "" Then
      series_discount1 = rs1!SeriesName & "-" & rs1!discount
     Else
      series_discount1 = series_discount & "," & rs1!SeriesName & "-" & rs1!discount
   End If
  End If
  
    If (kk >= 34 Or kk >= 48) Then
   If series_discount1 = "" Then
      series_discount1 = rs1!SeriesName & "-" & rs1!discount
     Else
      series_discount1 = series_discount & "," & rs1!SeriesName & "-" & rs1!discount
   End If
  End If
   
  kk = kk + 1
   
   
rs1.MoveNext
Wend





'==========================================================================================


DSNNew

'------------------------------------------------------------------------------------------
If Option1_sale.value = True Then
    If RS.State = 1 Then RS.close
    RS.Open "select partyName from ordera where invoiceNo='" & txtOrderNo.text & "'", con
    If RS.EOF = False Then
       printPRemarks
    End If
End If
'------------------------------------------------------------------------------------------


ss_ = ""
Dim frt As String


If RS.State = 1 Then RS.close
RS.Open "SELECT top 1 freight from SLEDGER" & _
" where SUBLEDGER='" & txtPartyName.text & "'", con
If RS.EOF = False Then
   'frt = "Freight:" & RS!freight & ""
   frt = "Freight:" & cbofrt.text & ""
   
End If


If RS.State = 1 Then RS.close
RS.Open "SELECT a.SUBLEDGER,b.PartyRemarks,a.orderno,Appno FROM INVOICEA as a " & _
" inner join SLEDGER as b on (a.SUBLEDGER = b.SUBLEDGER) " & _
" where a.orderno='" & txtOrderNo.text & "'", con
If RS.EOF = False Then
   con.Execute "update ordera set Bankaddress='" & RS(1) & "',bankname='" & RS(3) & "',pin='" & txtPin & "' where invoiceno=" & txtOrderNo.text & ""
Else
   con.Execute "update ordera set pin='" & txtPin & "' where invoiceno=" & txtOrderNo.text & ""
End If
  
  
con.Execute "UPDATE a SET a.pan = b.pan  FROM ORDERA AS a " & _
" INNER JOIN SLEDGER AS b ON (a.partyname = b.subledger " & _
" and  a.invoiceno = " & txtOrderNo & ")"





CR.Reset
CR.ReportFileName = rptPath & "/orderNew.rpt"
CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
CR.ReplaceSelectionFormula "{ordera.invoiceno}=" & txtOrderNo & ""
If Option1_sale.value = True Then
   CR.Formulas(0) = "sales_salesret='Sales Order'"
Else
   CR.Formulas(0) = "sales_salesret='Specimen Order'"
End If
CR.Formulas(1) = "frt='" & frt & "'"
CR.Formulas(2) = "seriesWiseDis='" & series_discount & "'"
CR.Formulas(3) = "seriesWiseDis1='" & series_discount1 & "'"
CR.Formulas(4) = "seriesWiseDis2='" & series_discount2 & "'"




If (Option2_sp.value = False) Then

If InStr(txtParty, "(EM)") > 0 Then
    
    If Len(lblCAF(1).Caption) > 5 Then
    CR.Formulas(5) = "mou_lbl='" & "" & "'"
    End If

Else

    If Len(lblCAF(1).Caption) > 5 Then
       CR.Formulas(5) = "mou_lbl='" & lblCAF(1).Caption & "'"
    End If

End If



End If


If rs1.State = 1 Then rs1.close
rs1.Open "select top 1 GroupOfSchool from collegeView_ind where CollegeID='" & txtScId.text & "'", CON_blue
If rs1.BOF = False Then
   If Not IsNull(rs1(0)) Then
      CR.Formulas(10) = "gpschool='" & "(" & rs1(0) & ")" & "'"
   End If

End If


CR.WindowShowPrintSetupBtn = True
CR.WindowShowRefreshBtn = True
CR.WindowMaxButton = True
CR.WindowState = crptMaximized
CR.Action = 1


Screen.MousePointer = vbDefault

End Sub

Private Sub cmdPrintLabel_Click()

Dim ph As String

ph = ""

If Option1_sale.value = True Then

 If rs1.State = 1 Then rs1.close
 rs1.Open "select phone,mobile from sledger where SUBLEDGER ='" & txtPartyName & "'", con
 If rs1.EOF = False Then
    'ph = rs1!phone
    If ph = "" Then
       ph = "Phone : " & rs1!phone
    End If
     
    If ph = "" Then
       ph = "Phone : " & rs1!mobile
    Else
       ph = ph & " " & rs1!mobile
    End If
 End If

Else


 If rs1.State = 1 Then rs1.close
 rs1.Open "select phone from rep where rep ='" & txtParty & "'", CON_blue
 If rs1.EOF = False Then
       ph = "Phone : " & rs1!phone
  End If


End If





If txtPin = "" Then
  pin = ""
Else
  pin = "- " & txtPin.text
End If


If Len(txtShip) > 0 And Len(txtBankName) > 0 Then
If txtContectNo <> "" Then
  ph = "Phone : " & txtContectNo
  pin = "- " & txtPIN_Ship.text
End If


con.Execute "exec printAddress '" & txtShip & "','" & txtBankName & "' ,'" & txtBankAdd & "','" & txtDist_ship & "','" & pin & "','" & txtShipState & "','" & ph & "'"
Else
con.Execute "exec printAddress '" & txtParty & "','" & txtPartyAdd1 & "' ,'" & txtPartyAdd2 & "','" & txtDist & "','" & pin & "','" & txtPartySt & "','" & ph & "'"
End If

DSNNew

CR.Reset
CR.ReportFileName = rptPath & "/AddressLabel.rpt"
CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
CR.WindowShowPrintSetupBtn = True
CR.WindowShowRefreshBtn = True
CR.WindowMaxButton = True
CR.WindowState = crptMaximized
CR.Action = 1

End Sub

Private Sub cmdPrintPending_Click()

Dim Q1, q2
Dim b1 As Boolean
Dim user_id As String
b1 = False

user_id = Trim((Sys_user_ + Str(UId)))

If txtOrderNo.text <> "" Then
'con.Execute "exec OrderWise_saleAndSpPending '" & txtOrderNo.Text & "'"
End If

con.Execute "delete from TmpBook1 where len(bcode)>0 and orderno='" & txtOrderNo.text & "'"
DoEvents
DoEvents
con.Execute "insert into  TmpBook1(orderno,BCode,BName,Qty,spQty) SELECT orderno,BOOKCODE,Bookname,Qty,SpQty FROM pendingOrderWise_summary where orderno='" & txtOrderNo.text & "'"
DoEvents
DoEvents

For I = 1 To vs.rows - 1
  
If vs.TextMatrix(I, 1) <> "" Then
   
   If Option1_sale.value = True Then
   
   Q1_sp = 0
   Q1_sp = IIf(vs.TextMatrix(I, 4) = "", 0, Val(vs.TextMatrix(I, 4)))
   Q1 = (IIf(vs.TextMatrix(I, 3) = "", 0, Val(vs.TextMatrix(I, 3))) + IIf(vs.TextMatrix(I, 4) = "", 0, Val(vs.TextMatrix(I, 4))))
   
   q2 = IIf(vs.TextMatrix(I, 5) = "", 0, vs.TextMatrix(I, 5))
   Q1 = Q1 - q2
   
   q2_sale = IIf(vs.TextMatrix(I, 3) = "", 0, vs.TextMatrix(I, 3))
   
   If Q1 > 0 Then
      'con.Execute "insert into TmpBook1(bcode,bname,Qty,issueQty,orderno) values('" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "'," & Q1 & "," & q1_sp & "," & txtOrderNo.Text & ")"
      b1 = True
   End If
   
Else
   
      Q1 = IIf(vs.TextMatrix(I, 4) = "", 0, vs.TextMatrix(I, 4))
      q2 = IIf(vs.TextMatrix(I, 5) = "", 0, vs.TextMatrix(I, 5))
      q2 = Q1 - q2
      Q1 = 0  ' IIf(vs.TextMatrix(I, 4) = "", 0, vs.TextMatrix(I, 4))

   
        If q2 > 0 Then
           'con.Execute "insert into TmpBook1(bcode,bname,Qty,issueQty,orderno) values('" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "'," & Q1 & "," & q2 & "," & txtOrderNo.Text & ")"
           b1 = True
        End If
   
   End If
   
   
   
   
    
End If


Next


'If Option1_sale.value = False Then
   'con.Execute "delete from TmpBook1 where issueQty=0 and orderno=" & txtOrderNo.Text & ""
   'con.Execute "delete from pendingSpList where Qty+SpQty=0 and INVOICENO=" & txtOrderNo.Text & ""
'End If

ss_ = ""
Dim frt As String


If RS.State = 1 Then RS.close
RS.Open "SELECT top 1 freight from SLEDGER" & _
" where SUBLEDGER='" & txtPartyName.text & "'", con
If RS.EOF = False Then
   frt = "Freight:" & RS!freight & ""
End If



If b1 = False Then
   MsgBox "No Pending ......", vbCritical
   Exit Sub
End If

DSNNew

MsgBox "want to view ?", vbInformation


CR.Reset
CR.ReportFileName = rptPath & "/Pending_order.rpt"
CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
CR.ReplaceSelectionFormula "{ordera.invoiceno}=" & txtOrderNo & " and ({TmpBook1.qty}>0 or {TmpBook1.spqty}>0)"
If Option1_sale.value = True Then
   CR.Formulas(0) = "sales_salesret='Pending Sales Order'"
Else
   CR.Formulas(0) = "sales_salesret='Pending Specimen Order'"
End If
CR.Formulas(1) = "frt='" & frt & "'"
CR.WindowShowPrintSetupBtn = True
CR.WindowShowRefreshBtn = True
CR.WindowMaxButton = True
CR.WindowState = crptMaximized
CR.Action = 1


End Sub
Function noofbox(qty As String, bcode As String) As Double
Dim rs_gaddi As New ADODB.Recordset
Set rs_gaddi = New ADODB.Recordset
Set rs_gaddi = con.Execute("exec BookSearch_bycode '" & session & "'," & main.setupid & ",'" & bcode & "'")
If rs_gaddi.EOF = False Then
   If Not IsNull(rs_gaddi!noofbox) Then
   If rs_gaddi!noofbox > 0 Then
      noofbox = Round(qty / rs_gaddi!noofbox, 2)
   End If
   End If
End If

End Function


Function noofgaddi(qty As String, bcode As String) As Double

Dim rs_gaddi As New ADODB.Recordset
Set rs_gaddi = New ADODB.Recordset
Set rs_gaddi = con.Execute("exec BookSearch_bycode '" & session & "'," & main.setupid & ",'" & bcode & "'")
If rs_gaddi.EOF = False Then
   If Not IsNull(rs_gaddi!BooksInGaddi) Then
   If rs_gaddi!BooksInGaddi > 0 Then
      noofgaddi = Round(qty / rs_gaddi!BooksInGaddi, 2)
   End If
   End If
End If

End Function
Sub seriesWiseMessage()

seriesWiseDis_ = ""

For I = 1 To vs.rows - 1
If vs.TextMatrix(I, 6) <> "" Then

    If vs.TextMatrix(I, 12) = "n" Then
        If seriesWiseDis_ = "" Then
           seriesWiseDis_ = vs.TextMatrix(I, 1)
        Else
           seriesWiseDis_ = seriesWiseDis_ & "," & vs.TextMatrix(I, 1)
        End If
        DoEvents
    End If
    

End If

Next

If seriesWiseDis_ <> "" Then
   seriesWiseDis_ = "Series Wise Discount is Not Set : " & seriesWiseDis_
End If


End Sub

Sub reserial()
   For I = 1 To vs.rows - 1
      
      If vs.TextMatrix(I, 1) <> "" Then
         vs.TextMatrix(I, 0) = I
      End If
      
'      If RS.State = 1 Then RS.close
'      RS.Open "select bookcode from books where bookname='" & vs.TextMatrix(I, 2) & "'", con
'      If RS.EOF = False Then
'         vs.TextMatrix(I, 1) = RS(0)
'      End If
   
   Next
End Sub
Function noDealing(party As String) As Boolean

Set RS = New ADODB.Recordset
RS.Open "select Profile_ from SLEDGER where SUBLEDGER='" & party & "'", con
If RS.EOF = False Then
       
    If RS!profile_ = "NO DEALING" Then
          noDealing = False
     Else
       noDealing = True
    End If
    
Else

noDealing = True
 
End If
 
End Function
Private Sub cmdSave_2_Click()





If (noDealing(txtPartyName.text) = False) Then
     MsgBox "NO DEALING Party....", vbCritical
     Exit Sub
End If




If txtShip.text <> "" Then

If (txtContectNo = "" And txtPIN_Ship = "") Then

    MsgBox "Enter Contact No....", vbCritical
    txtContectNo.SetFocus
    
    Exit Sub
End If

End If


If cboOrderBy = "" Then
   MsgBox "Select Order By....", vbCritical
   cboOrderBy.SetFocus
   Exit Sub
End If

If txtParty = "" Then
   MsgBox "Enter The Party....", vbCritical
   txtParty.SetFocus
   Exit Sub
End If


If txtMark.text = "" Then
   MsgBox "Select Godown....", vbCritical
   txtMark.SetFocus
   Exit Sub
End If



TXTBAL.text = 0
If Option1_sale.value = True Then

   PartyLedgerNew txtPartyName.text
   
End If


Dim noofgaddi_, bagInBox As Double
noofgaddi_ = 0
bagInBox = 0



reserial

If Edit = False Then

   
   If Check1_AddOrderNo.value = 0 Then
      maxOrder
   End If
   
   
   Set rs1 = con.Execute("exec HowtoCheckorderNo '" & txtOrderNo.text & "'")
   If rs1.EOF = False Then
         MsgBox "This Order No Already Found ...", vbCritical
         txtOrderNo.SetFocus
         Exit Sub
  End If
  
   If RS.State = 1 Then RS.close
   RS.Open "select top 1 * from ordera where INVOICENO=" & txtOrderNo.text & "", con, adOpenDynamic, adLockOptimistic
   
   If RS.EOF = True Then
   
    RS.AddNew
    
    RS!Frt_Yes = cbofrt.text
    
    RS!through_ = cboDisType.text
    RS!pin_ship = txtPIN_Ship.text
    RS!ccattach = cbobiltyrem.text
    RS!GroupOfSchool = lblgpSchool(27).Caption
    RS!noofgaddi = IIf(txtNoOfGaddi.text = "", 0, txtNoOfGaddi.text)
    
    RS!BagIn_Box = IIf(txtbagInbox.text = "", 0, txtbagInbox.text)
    
    RS!bal = IIf(TXTBAL.text = "", 0, TXTBAL.text)
    RS!bilty = cboBilty.text
    RS!sale_sp = sale_sp_
    RS!Godown = txtMark.text
    RS!partyname = txtPartyName.text
    RS!RepName = Trim(cmbAgentName.text)
    RS!scname = Trim(txtSchool)
    RS!scid = Trim(txtScId.text)
    RS!invoiceNo = Val(txtOrderNo)
    RS!invoiceDate = txtOrderDate
    RS!subledger = Trim(txtParty)
    RS!address1 = Trim(UCase(txtPartyAdd1))
    RS!address2 = Trim(UCase(txtPartyAdd2))
    RS!orderby = cboOrderBy
    RS!ORDERDATE = txtOrderDate1.value
    RS!transport = Trim(cboTrans)
    RS!TransAdd = Trim(txtTransAdd)
    RS!narration = Trim(txtNarration)
    RS!Shipto_Add1 = Trim(UCase(txtBankName))
    RS!Shipto_Add2 = Trim(UCase(txtBankAdd))
    RS!shipto_dist = Trim(UCase(txtDist_ship))
    RS!Shipto_States = UCase(txtShipState.text)
    RS!BookingStn = Trim(txtBookingStn)
    
    RS!netamount = Round(txtTotal, 0)
    RS!gamount = Round(txtGAmt, 0)
    
    RS!party_dist = Trim(txtDist)
    RS!fyear = session
    RS!setupid = setupid
    RS!Shipto = Trim(txtShip.text)
    RS!party_state = UCase(txtPartySt)
    RS!ContactNo = Trim(txtContectNo)
    RS.update
   
   End If
   
   
   createLog UserName, txtOrderNo.text, "SaleOrder", " Add : Qty " & txtTotQty.text, Date
   
   '''' OrderB
   If RS.State = 1 Then RS.close
   RS.Open "select * from orderb where invoiceno=" & txtOrderNo & "", con, adOpenDynamic, adLockOptimistic
   If RS.EOF = True Then
   
   For I = 1 To vs.rows - 1
            If vs.TextMatrix(I, 1) <> "" Then
                 
                ''========================
                   qty = 0
                   net = 0
                   qty = IIf(vs.TextMatrix(I, 3) = "", 0, vs.TextMatrix(I, 3))
                   
                   If Option1_sale.value = False Then
                      qty = Val(qty) + Val(IIf(vs.TextMatrix(I, 4) = "", 0, vs.TextMatrix(I, 4)))
                   End If
                   
                   net = (qty * Val(vs.TextMatrix(I, 6)))
                   
                   Set rs1 = con.Execute("exec fatchDiscount_partywise '" & Trim(Mid(txtPartyName.text, 1, 5)) & "', '" & vs.TextMatrix(I, 1) & "'")
                   If rs1.EOF = False Then
                      vs.TextMatrix(I, 9) = rs1(0)
                   End If
                   
                  
                   If (DateDiff("d", Now, SessionLastDate) <= 0) Then
                    
                        Set rs1 = con.Execute("exec fatchDiscount_SeriresAndpartywisefinal_test '" & Trim(Mid(txtPartyName.text, 1, 5)) & "', '" & vs.TextMatrix(I, 1) & "','" & txtScId.text & "'")
                        If rs1.EOF = False Then
                           vs.TextMatrix(I, 9) = rs1(0)
                           vs.TextMatrix(I, 12) = "y"
                        Else
                           vs.TextMatrix(I, 12) = "n"
                        End If
                     
                    Else

                        
                        Set rs1 = con_LAST.Execute("exec fatchDiscount_SeriresAndpartywisefinal_test '" & Trim(Mid(txtPartyName.text, 1, 5)) & "', '" & vs.TextMatrix(I, 1) & "','" & txtScId.text & "'")
                        If rs1.EOF = False Then
                           vs.TextMatrix(I, 9) = rs1(0)
                           vs.TextMatrix(I, 12) = "y"
                        Else
                           vs.TextMatrix(I, 12) = "n"
                        End If


                    End If
                  
                   
                   
                   If vs.TextMatrix(I, 9) <> "" Then
                   vs.TextMatrix(I, 7) = net - Format(Round(net * (vs.TextMatrix(I, 9) / 100), 2), "0.00")
                   End If
                   vs.TextMatrix(I, 11) = (qty * IIf(vs.TextMatrix(I, 6) = "", 0, vs.TextMatrix(I, 6)))
                
                ''========================
            
                 RS.AddNew
                 RS!invoiceNo = Val(txtOrderNo)
                 RS!invoiceDate = txtOrderDate
                 RS!Bookcode = vs.TextMatrix(I, 1)
                 RS!QUANTITY = Val(vs.TextMatrix(I, 3))
                 RS!Spqty = Val(vs.TextMatrix(I, 4))
                 RS!unit = vs.TextMatrix(I, 5)
                 RS!rate = Val(vs.TextMatrix(I, 6))
                 RS!amount = Val(vs.TextMatrix(I, 7))
                 RS!PRINTORDER = Val(vs.TextMatrix(I, 0))
                 RS!billno = vs.TextMatrix(I, 8)
                 RS!discount = Val(vs.TextMatrix(I, 9))
                 RS!pending = vs.TextMatrix(I, 10)
                 RS!gamount = IIf(vs.TextMatrix(I, 11) = "", 0, vs.TextMatrix(I, 11))
                 
                 'RS!noofgaddi = noofgaddi(Val((IIf(vs.TextMatrix(I, 3) = "", 0, vs.TextMatrix(I, 3))) + Val(IIf(vs.TextMatrix(I, 4) = "", 0, vs.TextMatrix(I, 4)))), vs.TextMatrix(I, 1))
                 
                 gaddi = 0
                 gaddi = noofgaddi(Val((IIf(vs.TextMatrix(I, 3) = "", 0, vs.TextMatrix(I, 3))) + Val(IIf(vs.TextMatrix(I, 4) = "", 0, vs.TextMatrix(I, 4)))), vs.TextMatrix(I, 1))
                 RS!noofgaddi = gaddi
                 noofgaddi_ = noofgaddi_ + gaddi


                 bagNo = 0
                 bagNo = noofbox(Val((IIf(vs.TextMatrix(I, 3) = "", 0, vs.TextMatrix(I, 3))) + Val(IIf(vs.TextMatrix(I, 4) = "", 0, vs.TextMatrix(I, 4)))), vs.TextMatrix(I, 1))
                 RS!noofbox = bagNo
               
                 bagInBox = bagInBox + bagNo

                 
                 
                 RS.update
            End If
   Next
   End If

End If


txtNoOfGaddi = Round(noofgaddi_, 2)
txtbagInbox.text = Round(bagInBox, 2)




Total

 
con.Execute "UPDATE a SET a.pan = b.pan,NETAMOUNT=" & txtTotal.text & ",GAmount=" & txtGAmt.text & "  FROM ORDERA AS a " & _
" INNER JOIN SLEDGER AS b ON (a.partyname = b.subledger " & _
" and  a.invoiceno = " & txtOrderNo & ")"

con.Execute "UPDATE ORDERA set noofgaddi=" & noofgaddi_ & ",BagIn_Box=" & txtbagInbox.text & " where  invoiceno = " & txtOrderNo & ""
 




DoEvents
DoEvents
DoEvents
DoEvents



    Set rs1 = New ADODB.Recordset
    
    'old code
''    If txtScId.text = "" Then
''     rs1.Open "select top 1 * from SeriesWiseDiscount where substring(Party,1,5)='" & Mid(txtPartyName.text, 1, 5) & "'", con
''    Else
''     rs1.Open "select top 1 * from SeriesWiseDiscountQry where substring(Party,1,5)='" & Mid(txtPartyName.text, 1, 5) & "' and scid='" & txtScId.text & "'", con
''    End If

    
    
    If (DateDiff("d", Now, SessionLastDate) <= 0) Then
    
        If txtScId.text = "" Then
            rs1.Open "select top 1 * from SeriesWiseDiscount where substring(Party,1,5)='" & Mid(txtPartyName.text, 1, 5) & "'", con
        Else
            rs1.Open "select top 1 * from SeriesWiseDiscountQry where substring(Party,1,5)='" & Mid(txtPartyName.text, 1, 5) & "' and scid='" & txtScId.text & "'", con
        End If
    
    Else
        If txtScId.text = "" Then
            rs1.Open "select top 1 * from SeriesWiseDiscount where substring(Party,1,5)='" & Mid(txtPartyName.text, 1, 5) & "'", con_LAST
        Else
            rs1.Open "select top 1 * from SeriesWiseDiscountQry where substring(Party,1,5)='" & Mid(txtPartyName.text, 1, 5) & "' and scid='" & txtScId.text & "'", con_LAST
        End If

    End If
    
    
    
    
    
    
    
    
    
    
    If rs1.EOF = False Then
    
        seriesWiseMessage
        seriesWiseDis_ = "" & seriesWiseDis_
        
        If Len(seriesWiseDis_) > 2 Then
           MsgBox seriesWiseDis_, vbInformation
        End If
        
    End If
        





lblSave.Caption = "Data saved....."

DoEvents
DoEvents

Check1_AddOrderNo.value = 0

cmdSave_2.Enabled = False
cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = True
cmdEdit_4.SetFocus

End Sub

'Sub updateSchooll(scid As String, scname As String)
'con.Execute "update INVOICEA set ScName='" & scname & "'  where ScID ='" & scid & "'"
'con.Execute "update CreditA set ScName='" & scname & "'  where ScID ='" & scid & "'"
'con.Execute "update INVOICEA_sp set ScName='" & scname & "' where ScID ='" & scid & "'"
'con.Execute "update ORDERA set ScName='" & scname & "'  where ScID ='" & scid & "'"
'con.Execute "update AppForm  set School_PartyName='" & scname & "'  where (ID ='" & scid & "' and School_Party='School')"
'con.Execute "update AppForm set PName='" & scname & "'  where (code ='" & scid & "' and School_Party='party')"
'
'End Sub

Sub cmdDisVisible()

If Option1_sale.value = True Then

   If (LCase(UserName) = "admin") Then
     If (txtPartyName.text <> "") Then
      cmdDis.Visible = True
     End If
   End If
   
End If
   
End Sub

Sub searchData()
    
    
On Error GoTo err1
    
   txtrem.text = ""
    
   lblSave.Caption = ""
   If RS.State = 1 Then RS.close
   RS.Open "select * from ordera where invoiceno=" & txtOrderNo & "", con, adOpenDynamic, adLockOptimistic
   If RS.EOF = False Then
   
   
   
   
    add_ = False
    vs.Enabled = True
    cmdSave_2.Enabled = False
    cmdDelete_3.Enabled = False
    cmdEdit_4.Enabled = True
    cmd_update.Enabled = False
    
    If Not IsNull(RS!through_) Then
      If Len(RS!through_) > 0 Then
       cboDisType.text = RS!through_
      End If
    End If
    
    cbofrt.text = RS!Frt_Yes & ""
    
    txtbagInbox.text = RS!BagIn_Box & ""
    
    txtPIN_Ship.text = RS!pin_ship & ""
    
     lblgpSchool(27).Caption = RS!GroupOfSchool & ""
    
    cbobiltyrem.text = RS!ccattach & ""
    cboBilty.text = RS!bilty & ""
    If RS!sale_sp = "sale" Then
       Option1_sale.value = True
     Else
       Option2_sp.value = True
    End If
    
    
    If Not IsNull(RS!bal) Then
       TXTBAL.text = RS!bal
    Else
       TXTBAL.text = 0
    End If
    txtMark = RS!Godown & ""
    txtPartyName.text = RS!partyname & ""
    
    
    cmdDisVisible
    
    txtContectNo = RS!ContactNo & ""
    txtShip = RS!Shipto & ""
    txtPartyName.text = RS!partyname & ""
    txtSchool.text = RS!scname & ""
    txtScId.text = RS!scid & ""
    cmbAgentName.text = RS!RepName & ""
    txtOrderNo = RS!invoiceNo
    txtOrderDate = RS!invoiceDate
    txtParty = RS!subledger
    txtPartyAdd1 = RS!address1 & ""
    txtPartyAdd2 = RS!address2 & ""
    cboOrderBy = RS!orderby & ""
    txtOrderDate1.value = RS!ORDERDATE
    cboTrans = RS!transport & ""
    txtTransAdd = RS!TransAdd & ""
    txtNarration = RS!narration & ""
    txtBankName = RS!Shipto_Add1 & ""
    txtBankAdd = RS!Shipto_Add2 & ""
    txtDist_ship = RS!shipto_dist & ""
    txtShipState.text = RS!Shipto_States & ""
    txtBookingStn = RS!BookingStn & ""
    txtTotal = RS!netamount
    txtGAmt.text = RS!gamount & ""
    
    txtDist = RS!party_dist & ""
    txtDist_ship = RS!shipto_dist & ""
    txtPartySt = RS!party_state & ""
    txtPin = ""
    If Option2_sp.value = True Then
        txtPin = RS!pin & ""
    Else
        
        txtrem.text = ""
        If rs1.State = 1 Then rs1.close
        rs1.Open "select pin,PartyRemarks from sledger where SUBLEDGER ='" & txtPartyName & "'", con
        If rs1.EOF = False Then
           txtPin = rs1!pin
           txtrem.text = rs1!PartyRemarks & ""
        End If
    
    End If
    
   
   
   
   
   End If
   
   
   
   
Dim Q1 As Integer
Dim inv As String
Dim noofgaddi As Double
noofgaddi = 0
   
   
str1 = "SELECT  PRINTORDER,ORDERB.BOOKCODE, BOOKNAME, ORDERB.RATE,ORDERB.discount,ORDERB.billno, QUANTITY,unit,AMOUNT,orderb.pending,orderb.spQty,orderb.onlineAmt,orderb.gamount,orderb.noofgaddi,orderb.sno " & _
"FROM ORDERB INNER JOIN BOOKS ON ORDERB.BOOKCODE = BOOKS.BOOKCODE where invoiceno=" & txtOrderNo & " order by printorder"
If RS.State = 1 Then RS.close
RS.Open str1, con, adOpenForwardOnly, adLockReadOnly
For I = 1 To RS.RecordCount
       vs.TextMatrix(I, 0) = I
       vs.TextMatrix(I, 1) = RS!Bookcode
       vs.TextMatrix(I, 2) = RS!Bookname
       vs.TextMatrix(I, 3) = RS!QUANTITY
       vs.TextMatrix(I, 4) = RS!Spqty & ""
       vs.TextMatrix(I, 10) = RS!pending
       vs.TextMatrix(I, 11) = Round(RS!gamount, 2) & ""
       If Not IsNull(RS!noofgaddi) Then
          noofgaddi = noofgaddi + RS!noofgaddi
       End If
       
       inv = ""
       Q1 = 0
       
       
       If Option1_sale.value = True Then
       
       'If rs1.State = 1 Then rs1.close
       'rs1.Open "select sum(QUANTITY),invoiceno from InvoicebAndInvoicebSP_Qry where OrderNo=" & txtOrderNo & " and bookcode='" & RS!Bookcode & "' group by invoiceno", con
       Set rs1 = New ADODB.Recordset
       Set rs1 = con.Execute("exec SP_PendingOrder_saleAndSp '" & txtOrderNo & "','" & RS!Bookcode & "'")
       While rs1.EOF = False
           Q1 = Q1 + rs1(0)
           If inv = "" Then
              inv = rs1!invoiceNo
             Else
              inv = inv & "," & rs1!invoiceNo
           End If
       rs1.MoveNext
       Wend
       
       'If rs1.EOF = False Then
       If Q1 > 0 Then
            vs.TextMatrix(I, 5) = Q1
            vs.TextMatrix(I, 8) = inv     'rs1!invoiceno
            vv = (IIf(vs.TextMatrix(I, 3) = "", 0, Val(vs.TextMatrix(I, 3))) + IIf(vs.TextMatrix(I, 4) = "", 0, Val(vs.TextMatrix(I, 4))))
            vv1 = IIf(vs.TextMatrix(I, 5) = "", 0, Val(vs.TextMatrix(I, 5)))
            If vv = vv1 Then
                For k1 = 0 To 11
                    vs.Cell(flexcpBackColor, I, k1) = vbGreen
                    DoEvents
                Next
                vs.TextMatrix(I, 10) = "n"
            End If
        End If
        
Else
     
       inv = ""
       Q1 = 0
     
       If rs1.State = 1 Then rs1.close
       rs1.Open "select sum(QUANTITY),invoiceno from invoicespBQry where OrderNo=" & txtOrderNo & " and bookcode='" & RS!Bookcode & "' group by invoiceno", con
       While rs1.EOF = False
           Q1 = Q1 + rs1(0)
           If inv = "" Then
              inv = rs1!invoiceNo
             Else
              inv = inv & "," & rs1!invoiceNo
           End If
       rs1.MoveNext
       Wend
     
     
     
       'If rs1.State = 1 Then rs1.close
       'rs1.Open "select sum(QUANTITY),invoiceno from invoicespBQry where OrderNo=" & txtOrderNo & " and bookcode='" & RS!Bookcode & "' group by invoiceno", con
       'If rs1.EOF = False Then
       If Q1 > 0 Then
            vs.TextMatrix(I, 5) = Q1   'rs1(0)
            vs.TextMatrix(I, 8) = inv     'rs1!invoiceno
            If vs.TextMatrix(I, 4) = vs.TextMatrix(I, 5) Then
                For k1 = 0 To 11
                    vs.Cell(flexcpBackColor, I, k1) = vbGreen
                    DoEvents
                Next
                vs.TextMatrix(I, 10) = "n"
            End If
       Else
           
          
          
       
       End If
       End If
       
       
       vs.TextMatrix(I, 6) = RS!rate
       vs.TextMatrix(I, 7) = Round(RS!amount, 2)
       vs.TextMatrix(I, 9) = RS!discount & ""
       vs.TextMatrix(I, 13) = RS!sno
       RS.MoveNext
Next
   
If txtParty.text <> "" Then
   checkOrder
End If
   
   
txtNoOfGaddi.text = noofgaddi

Total

seriesWiseDiscount

lblCAF(1).Caption = ""
code_ = Trim(Mid(txtPartyName.text, 1, 5))

If (Option2_sp.value = False) Then

    If (code_ <> "") Then
    lblCAF(1).Caption = fillDocument("" & code_)
    End If

End If
 

mnuMenu_ = "mnuSaleOrder"

'SetButton cmdEdit_4, cmdEdit_4, cmdSave_2, cmdDelete_3

Exit Sub

err1:

MsgBox "" & err.DESCRIPTION

End Sub

Private Sub cmdSave_Click()

For I = 1 To vs1.rows - 1

If vs1.TextMatrix(I, 0) <> "" Then
  con.Execute "update ORDERB set Pending='" & vs1.TextMatrix(I, 5) & "' where INVOICENO=" & vs1.TextMatrix(I, 1) & " and BOOKCODE='" & vs1.TextMatrix(I, 2) & "'"
  con.Execute "update TmpBook set head='" & vs1.TextMatrix(I, 5) & "' where OrderNo=" & vs1.TextMatrix(I, 1) & " and BCODE='" & vs1.TextMatrix(I, 2) & "'"
End If

Next


mnuMenu_ = "mnuSaleOrder"
SetButton cmdEdit_4, cmdEdit_4, cmdSave_2, cmdDelete_3


MsgBox "Data updated....", vbInformation
End Sub

Private Sub cmdSDis_Click()

    
frmSeriesWiseDis.cmdSave_2.Enabled = True
frmSeriesWiseDis.cmdDelete_3.Enabled = True
frmSeriesWiseDis.cmdAdd_1.Enabled = True
frmSeriesWiseDis.cmdEdit_4.Enabled = True

frmSeriesWiseDis.Show 1
        
End Sub

Private Sub cmdSendMail_Click()


On Error Resume Next

Dim rsf As New ADODB.Recordset
Dim rs_inv As New ADODB.Recordset

Dim k1 As Integer
k1 = 1


'-----------------------------


Set rs_inv = New ADODB.Recordset
Set rs_inv = con.Execute("exec searchList 'orderno'")

'----------------------------



vs_Mail.Clear
frmMail.Visible = True

vs_Mail.rows = 2
txt_NetAmt.text = 0

Set rsf = New ADODB.Recordset
rsf.Open "SELECT INVOICENO as OrderNo,INVOICEDATE as OrderDate,PartyName,Address1,Address2,MailSended,SUBLEDGER,Sale_sp,NETAMOUNT,REPNAME FROM ORDERA where (mailsended='no' AND Sale_sp='sale')", con
While rsf.EOF = False

DoEvents
vs_Mail.Cell(flexcpFontBold, k1, 0) = True
vs_Mail.Cell(flexcpFontSize, k1) = 11
txt_NetAmt.text = Val(txt_NetAmt.text) + rsf!netamount
DoEvents
vs_Mail.TextMatrix(k1, 0) = rsf(0)

vs_Mail.TextMatrix(k1, 1) = rsf(1)

If rsf!sale_sp = "sale" Then
   vs_Mail.TextMatrix(k1, 2) = rsf(2)
Else
   vs_Mail.TextMatrix(k1, 2) = rsf!subledger
End If

vs_Mail.TextMatrix(k1, 3) = rsf!RepName
'vs_Mail.TextMatrix(k1, 4) = rsf(2)

If LCase(rsf(5)) = "no" Then
   vs_Mail.TextMatrix(k1, 5) = 0
Else
   vs_Mail.TextMatrix(k1, 5) = 1
End If

vs_Mail.rows = vs_Mail.rows + 1

rs_inv.MoveFirst
If rs_inv.EOF = False Then

rs_inv.Find "orderno=" & rsf(0) & ""
If rs_inv.EOF = False Then

    For kk1 = 0 To 5
      vs_Mail.Cell(flexcpBackColor, k1, kk1) = vbGreen
    DoEvents
    Next
        
End If
End If




k1 = k1 + 1
rsf.MoveNext
Wend



vs_Mail.FormatString = "<OrderNo|OrderDate|PartyName|RepName||Mail"
vs_Mail.ColWidth(0) = 1100
vs_Mail.ColWidth(1) = 1300
vs_Mail.ColWidth(2) = 5000
vs_Mail.ColWidth(3) = 3500
vs_Mail.ColWidth(4) = 0
vs_Mail.ColWidth(5) = 1000


End Sub

Private Sub Command1_Click()
frmPendingClear.Visible = True

vs1.Cols = 6
vs1.ColComboList(5) = "n|y"

If rs1.State = 1 Then rs1.close
rs1.Open "select Area,OrderNo,BCode,BName,Qty,issueQty,head from TmpBook  where Login='" & UserName & "' order by Area,OrderNo,BCode", con
If rs1.EOF = False Then
vs1.rows = rs1.RecordCount + 1
End If

For I = 1 To rs1.RecordCount

vs1.TextMatrix(I, 0) = Trim(Mid(rs1!Area, 6))
vs1.TextMatrix(I, 1) = rs1!orderNo
vs1.TextMatrix(I, 2) = rs1!bcode
vs1.TextMatrix(I, 3) = rs1!BName
vs1.TextMatrix(I, 4) = (rs1!qty - rs1!issueQty)
vs1.TextMatrix(I, 5) = rs1!head & ""
'vs1.TextMatrix(I, 6) = rs1!BalanceQty & ""


rs1.MoveNext
Next

vs1.FormatString = "Party|OrderNo|BCode|BName|Pending Qty|Pending Clear"
vs1.ColWidth(0) = 4000
vs1.ColWidth(1) = 1000
vs1.ColWidth(2) = 1000
vs1.ColWidth(3) = 3000
vs1.ColWidth(4) = 2000


End Sub

Private Sub Command7_Click()
frmMail.Visible = False
End Sub

Private Sub Form_Activate()
'SetButton cmdEdit_4, cmdEdit_4, cmdSave_2, cmdDelete_3
cmdSave.Enabled = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  If Me.frmPendingClear.Visible = True Then
   Me.frmPendingClear.Visible = False
  Else
   Unload Me
  End If
End If
End Sub
Sub sale_sp()
  If Option1_sale.value = True Then
     sale_sp_ = "sale"
  Else
     sale_sp_ = "sp"
  End If
  
  
  
  setVSWidth
End Sub
Private Sub Form_Load()

Timer1.Enabled = False

cboDisType.ListIndex = 0

cbofrt.ListIndex = 0


cboOrderBy.ListIndex = 0
cboBilty.ListIndex = 1

bcode = ""
BName = ""

Me.Width = 14400
Me.Height = 10950

bb = False

Me.top = 0
Me.Left = 0
setVSWidth

If RS.State = 1 Then RS.close
'RS.Open "select * from books where " & stringyear & " order by bookcode", CCON, adOpenDynamic, adLockReadOnly, adCmdText
RS.Open "select * from books where " & stringyear & " order by bookcode", CCON, adOpenDynamic, adLockReadOnly, adCmdText
If Not RS.BOF Then
    Do While Not RS.EOF
        
       If bcode = "" Then
        bcode = RS("bookcode")
       Else
        bcode = bcode & "|" & RS("bookcode")
       End If
        If Not RS.EOF Then
            RS.MoveNext
        End If
    Loop
End If
RS.close


If RS.State = 1 Then RS.close
RS.Open "select * from books where " & stringyear & " order by bookname", CCON, adOpenDynamic, adLockReadOnly, adCmdText
If Not RS.BOF Then
    Do While Not RS.EOF
        
       If bcode = "" Then
        BName = RS("bookname")
       Else
        BName = BName & "|" & RS("bookname")
       End If
        
        If Not RS.EOF Then
            RS.MoveNext
        End If
    Loop
End If
RS.close


'vs.ColComboList(1) = bcode
vs.ColComboList(2) = BName


RS.Open "select  transportname from transportMaster order by transportname", con, adOpenDynamic, adLockReadOnly, adCmdText
cboTrans.Clear
If Not RS.EOF Then
   Do While Not RS.EOF
      If IsNull(RS(0)) = False Then
        Me.cboTrans.AddItem RS(0)
      End If
      If Not RS.EOF Then RS.MoveNext
    Loop
End If
RS.close



maxOrder

'-----------------------------------------------
'*******Agent  combo fill-----------------------
'popuplist10 "select Rep as Representative,Add1,Add2,District,[state] from SalesRepQry order by Rep", CON_blue
If RS.State = 1 Then RS.close
RS.Open "select Rep as Representative from SalesRepQry where (email is not null and len(email)>1) order by Rep", CON_blue
cmbAgentName.Clear
If Not RS.EOF Then
   Do While Not RS.EOF
      If IsNull(RS(0)) = False Then
        Me.cmbAgentName.AddItem RS(0)
      End If
      If Not RS.EOF Then RS.MoveNext
    Loop
End If


txtOrderDate.value = Format(Date, "dd/MM/yyyy")
txtOrderDate1.value = Format(Date, "dd/MM/yyyy")

Dim rs_godwn As New ADODB.Recordset

If rs_godwn.State = 1 Then rs_godwn.close
rs_godwn.Open "select * from GodownMaster where len(Godwn)<=3 and " & stringyear & " order by id", con, adOpenForwardOnly, adLockReadOnly
txtMark.Clear
cbogd1.Clear
If Not rs_godwn.EOF Then
Do While Not rs_godwn.EOF
   If IsNull(rs_godwn(0)) = False Then
     Me.txtMark.AddItem rs_godwn(0)
     Me.cbogd1.AddItem rs_godwn(0)
   End If
   If Not rs_godwn.EOF Then rs_godwn.MoveNext
 Loop
End If



If popupvalue5 <> "" Then

   If RS.State = 1 Then RS.close
   RS.Open "select Sale_sp from ordera where invoiceno='" & popupvalue5 & "'", con
   If RS.EOF = False Then
      sale_sp_ = RS(0)
   End If

   txtOrderNo = popupvalue5
   searchData
   popupvalue5 = ""
   Exit Sub
End If

sale_sp
vs.ColComboList(10) = "y|n"
Edit = False
cmdSave_2.Enabled = True
cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = False
cmd_update.Enabled = False

BackColorFrom Me
    
cbogd1.ListIndex = 0



End Sub
Sub setVSWidth()

 
If sale_sp_ = "sale" Then

    vs.FormatString = "S.N|B.CODE|BOOK NAME|QTY.|Sp.QTY.|Billing Qty|PRICE|NetAmt|Bill No|Dis.||GAmount|"
    
    vs.ColWidth(0) = 500
    vs.ColWidth(1) = 1100
    vs.ColWidth(2) = 3800
    vs.ColWidth(3) = 900
    vs.ColWidth(4) = 1000
    vs.ColWidth(5) = 1100
    vs.ColWidth(6) = 1100
    vs.ColWidth(7) = 1200
    vs.ColWidth(8) = 1000
    vs.ColWidth(9) = 700
    vs.ColWidth(10) = 0
    vs.ColWidth(11) = 1200
    vs.ColWidth(12) = 0
    vs.ColWidth(13) = 0
    
Else

    vs.FormatString = "S.N|B.CODE|BOOK NAME||Sp.QTY.|Billing Qty|PRICE|NetAmt|Bill No|Dis.||GAmount|"
    
    vs.ColWidth(0) = 500
    vs.ColWidth(1) = 1100
    vs.ColWidth(2) = 3800
    vs.ColWidth(3) = 0
    vs.ColWidth(4) = 1000
    vs.ColWidth(5) = 1100
    vs.ColWidth(6) = 1100
    vs.ColWidth(7) = 1200
    vs.ColWidth(8) = 1000
    vs.ColWidth(9) = 700
    vs.ColWidth(10) = 0
    vs.ColWidth(11) = 1200
    vs.ColWidth(12) = 0
    vs.ColWidth(13) = 0

End If
    
    
    
    
End Sub



Private Sub Label1_Click(Index As Integer)
If txtPartyName.text <> "" Then
    PopUpValue6 = txtPartyName.text
    
   If (LCase(UserName) = "admin") Then
    
        frmSeriesWiseDis.cmdSave_2.Enabled = True
        frmSeriesWiseDis.cmdDelete_3.Enabled = True
        frmSeriesWiseDis.cmdAdd_1.Enabled = True
        frmSeriesWiseDis.cmdEdit_4.Enabled = True
    Else
        frmSeriesWiseDis.cmdSave_2.Enabled = False
        frmSeriesWiseDis.cmdDelete_3.Enabled = False
        frmSeriesWiseDis.cmdAdd_1.Enabled = False
        frmSeriesWiseDis.cmdEdit_4.Enabled = False

    End If
    
    frmSeriesWiseDis.Show 1
Else
   MsgBox "Plz Search Party ..", vbInformation
End If
End Sub

Private Sub List_emptyList_DblClick()
   txtOrderNo.text = List_emptyList.text
   Check1_AddOrderNo.value = 1
   List_emptyList.Visible = False
End Sub

Private Sub Option1_sale_Click()
sale_sp
End Sub

Private Sub Option2_sp_Click()
sale_sp
End Sub

Private Sub Text1_GotFocus()

If PopUpValue1 <> "" Then
   
   Text1.text = PopUpValue1
   cmdpendingbook.Enabled = True
   PopUpValue1 = ""
   
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
   searchType = "books"
   popuplist10 "select SerName from BOOKS where " & stringyear & "  group by SerName", con
End If

End Sub

Private Sub Timer1_Timer()
Static L As Integer
If L = 0 Then
    Label1(20).ForeColor = vbYellow
    Label1(21).ForeColor = vbBlue
    L = 1
    Exit Sub
ElseIf L = 1 Then
    Label1(20).ForeColor = vbBlue
    Label1(21).ForeColor = vbYellow
    L = 0
    Exit Sub
End If

End Sub

Private Sub Timer2_Timer()

Static L As Integer

If L = 0 Then
    Label1(25).ForeColor = vbYellow
    Label1(26).ForeColor = vbBlue
    L = 1
    Exit Sub
ElseIf L = 1 Then
    Label1(25).ForeColor = vbBlue
    Label1(26).ForeColor = vbYellow
    L = 0
    Exit Sub
End If

End Sub

Private Sub txtBankAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtDist_ship.SetFocus
End Sub
Private Sub txtBankName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtBankAdd.SetFocus
End Sub

Private Sub txtbk_Change()
cmdpendingbook.Enabled = True
End Sub

Private Sub txtbk_GotFocus()
If PopUpValue1 <> "" Then
   txtbk = PopUpValue1
   cmdpendingbook.Enabled = True
   PopUpValue1 = ""
   
End If
End Sub

Private Sub txtbk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then

 If Text1.text = "" Then
 
   searchType = "books"
   popuplist10 "select BOOKCODE,BOOKNAME from BOOKS where " & stringyear & "  order by BOOKCODE", con
   
 Else
 
    searchType = "books"
    popuplist10 "select BOOKCODE,BOOKNAME from BOOKS where SerName='" & Text1.text & "'  order by BOOKCODE", con

 End If
   
End If
End Sub

Private Sub txtBookingStn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   vs.SetFocus
   vs.Col = 1
   Do While vs.Row > 1
      sendkeys "{home}"
      vs.Row = vs.Row - 1
   Loop

End If

End Sub
Private Sub txtBookingStn_LostFocus()
txtBookingStn.text = UCase(txtBookingStn)
End Sub

Private Sub txtContectNo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   cmbAgentName.SetFocus
   'txtschool.SetFocus
   
End If

End Sub

Private Sub txtDist_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   txtShip.SetFocus
End If

End Sub
Private Sub txtDist_ship_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtPIN_Ship.SetFocus
End Sub

Private Sub txtMark_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   'cmbAgentName.SetFocus
   If txtMark.text <> "" Then
      cboDisType.SetFocus
   End If
End If
End Sub

Private Sub txtNarration_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtMark.SetFocus
End Sub

Private Sub txtOrderDate_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then cboOrderBy.SetFocus
End Sub

Private Sub txtOrderDate1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   cboBilty.SetFocus
End If
End Sub

Private Sub txtOrderNo_GotFocus()
If PopUpValue1 <> "" Then


refreshFld

If search_ = "f2" Then
   txtOrderNo = PopUpValue1
Else
   txtOrderNo = PopUpValue2
End If
searchData
setVSWidth

PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""

End If
End Sub

Private Sub txtOrderNo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    search_ = "f2"
    searchType = "inv"
    'popuplist10 "select InvoiceNo as OrderNo,InvoiceDate as OrderDate,Subledger as partyName,NetAmount from OrderA where " & stringyear & "  order by InvoiceNo", con
    popuplistFast "select InvoiceNo,InvoiceDate,Subledger,NetAmount from InvoiceA where " & stringyear & "  order by InvoiceNo", con, , , "ORDER"
    
    
ElseIf KeyCode = 112 Then
    search_ = "f1"
    searchType = "party"
    popuplist_client "select Subledger  + ',' + party_dist as PartyName,InvoiceNo as OrderNo,InvoiceDate as OrderDate,NetAmount from OrderA where " & stringyear & "  order by Subledger", con
    
End If

End Sub
Private Sub txtOrderNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       refreshFld
       searchData
       setVSWidth
       txtOrderDate.SetFocus
    End If

End Sub
Sub checkOrder()
    Dim s10 As String
    Dim J As Integer
    
    J = 1
    
    s10 = ""
    
    If rs1.State = 1 Then rs1.close
    
    If Option1_sale.value = True Then
    
    rs1.Open "select top 10 INVOICENO from ordera where partyname='" & txtPartyName.text & "'" & _
    " and INVOICEDATE=convert(smalldatetime,'" + Trim(Me.txtOrderDate.value) + "',103)", con
    
    Else
    
    rs1.Open "select top 10 INVOICENO from ordera where Subledger='" & txtParty.text & "'" & _
    " and INVOICEDATE=convert(smalldatetime,'" + Trim(Me.txtOrderDate.value) + "',103)", con


    End If
    
    While rs1.EOF = False
      
      If s10 = "" Then
         s10 = "Order No :  " & rs1!invoiceNo
      Else
         s10 = s10 & " , " & rs1!invoiceNo
      End If
      
      J = J + 1
    rs1.MoveNext
    Wend
    
    
    If J > 1 Then
      lblOrder_reminder.Caption = s10 & " of this party is already made today.."
    End If
    
    
End Sub
Sub seriesWiseDiscount()
  
  Dim rs_dis As New ADODB.Recordset
  
   
  If (DateDiff("d", Now, SessionLastDate) <= 0) Then
      rs_dis.Open "select top 1 * from SeriesWiseDiscount where substring(Party,1,5)='" & Mid(txtPartyName.text, 1, 5) & "'", con
  Else
      rs_dis.Open "select top 1 * from SeriesWiseDiscount where substring(Party,1,5)='" & Mid(txtPartyName.text, 1, 5) & "'", con_LAST
  End If
  
  If rs_dis.EOF = False Then
        Label1(25).Visible = True
        Label1(26).Visible = True
        Timer2.Enabled = True
  Else
        Label1(25).Visible = False
        Label1(26).Visible = False
        Timer2.Enabled = False
  End If


End Sub
Private Sub txtParty_GotFocus()
  
On Error Resume Next
  
If PopUpValue1 <> "" Then

txtrem.text = ""


lblCAF(1).Caption = ""

If (PopUpValue2 <> "") Then
   lblCAF(1).Caption = fillDocument(PopUpValue2)
End If

If Option2_sp.value = True Then
   
   txtPartyName.text = ""
   
   
   txtParty = PopUpValue1
   txtPartyAdd1 = PopUpValue2
   txtPartyAdd2 = PopUpValue3
   txtDist = popupvalue4
   txtPartySt = popupvalue5
   
   cmdDisVisible
   
   
   
   txtShip.SetFocus
   
   
   If RS.State = 1 Then RS.close
   RS.Open "select City,District,Pin,Phone from SalesRepQry where Rep='" & PopUpValue1 & "'", CON_blue
   If RS.EOF = False Then
        
      If LCase(RS!city) = LCase(RS!District) Then
         txtDist = RS!city
      Else
         txtDist = RS!city & " (" & RS!District & ")"
      End If
      txtPin.text = RS!pin & ""
      
      txtContectNo.text = RS!phone & ""
   End If
   
   
   
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   popupvalue5 = ""

Else

  
 
     
   
  ' txtPartyAdd1 = PopUpValue2
  ' txtPartyAdd2 = PopUpValue3
  ' txtPartySt = popupvalue5
  
  ' txtShip.SetFocus
     
   If RS.State = 1 Then RS.close
   RS.Open "select Transport,DESCFORINVOICE,DISTCODE,pan,ADDRESS1,ADDRESS2,ADDRESS3 as City,states as State,Subledger,pin,Profile_,freight,PartyRemarks from SLEDGER where SUBLEDGER='" & PopUpValue3 & "'", con
   If RS.EOF = False Then
       
      'If IsNull(RS!pan) Then
        If RS!pan = "" Then
           MsgBox "PAN Not available .... ", vbCritical
        End If
      'End If
       
       txtrem.text = RS!PartyRemarks & ""
       
       cbofrt.text = RS!freight & ""
       
       txtParty = RS!DESCFORINVOICE & ""
       txtPartyAdd1 = RS!address1 & ""
       txtPartyAdd2 = RS!address2 & ""
       txtPartySt = RS!State & ""

      
      cboTrans.text = RS!transport & ""
      txtPartyName.text = RS!subledger & ""
      
      
      If LCase(RS!distcode) = LCase(RS!city) Then
         txtDist = UCase(RS!distcode)
      Else
         txtDist = UCase(RS!city & " (" & RS!distcode & ")")
      End If
      
      txtPin.text = RS!pin & ""
      
      
      If UCase(RS!profile_) = UCase("CASH PARTY") Then
        Label1(20).Visible = True
        Label1(21).Visible = True
        Timer1.Enabled = True
      Else
        Label1(20).Visible = False
        Label1(21).Visible = False
        Timer1.Enabled = False
      End If
      
   End If
   
   
   seriesWiseDiscount
   
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   popupvalue5 = ""
   

   
     
End If
End If

  
End Sub

Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    If Option2_sp.value = True Then
    
        searchType = "party"
        popuplist_client "select Rep,Add1,Add2,City,State from SalesRepQry order by rep", CON_blue
    Else
        'searchType = "party"
        'popuplist_client "select DESCFORINVOICE,ADDRESS1,ADDRESS2,ADDRESS3 as District,states as State,Subledger from Sledger where " & stringyear & " order by Subledger", con
        searchType = "party"
        value = "select distinct(Party),Code,subledger from SLEDGER where gledger='SUNDRY DEBTORS' and " & stringyear & "  order by party"
        popuplist_client value, CCON
        set_focus = True
        
    End If
End If



If KeyCode = 13 Then
    If txtParty <> "" Then
     'txtschool.SetFocus
     txtPartyAdd1.SetFocus
    End If
    
    
'   TXTBAL.Text = 0
'
'   If Option1_sale.value = True Then
'      PartyLedgerNew txtPartyName.Text
'   End If

End If

End Sub
Private Sub txtParty_LostFocus()
If InStr(txtParty, "(EM)") > 0 Then
   party_type = "EM"
Else
   party_type = "BP"
End If

If txtParty.text <> "" Then
   checkOrder
End If


If txtParty.text = "" Then
   txtPartyName.text = ""
End If




If (txtPartyName.text <> "") Then
If (noDealing(txtPartyName.text) = False) Then
     MsgBox "NO DEALING Party....", vbCritical
     txtParty.SetFocus
     Exit Sub
End If
End If




End Sub

Private Sub txtPartyAdd1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   txtPartyAdd2.SetFocus
End If
End Sub

Private Sub txtPartyAdd2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   txtDist.SetFocus
End If

End Sub

Private Sub txtPIN_Ship_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cboTrans.SetFocus
End Sub

Private Sub txtschool_GotFocus()
If RS.State = 1 Then RS.close
If PopUpValue1 <> "" Then
   
a = txtSchool.MaxLength
   
txtScId = PopUpValue1

k_ = PopUpValue2 & ", " & PopUpValue3


If Len(k_) < 60 Then
  txtSchool.text = PopUpValue2 & ", " & PopUpValue3
Else
  txtSchool.text = PopUpValue2
End If

lblgpSchool(27).Caption = popupvalue5

PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
popupvalue4 = ""
popupvalue5 = ""

End If
End Sub
Private Sub txtschool_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
   
   Screen.MousePointer = vbHourglass
   tblNo = 9
   frmSearchItem.Show
   Screen.MousePointer = vbDefault
   
End If

If KeyCode = 13 Then
txtBookingStn.SetFocus
End If

End Sub

Private Sub txtShip_GotFocus()
If Check1_school.value = 0 Then

If RS.State = 1 Then RS.close
If PopUpValue1 <> "" Then
    
    txtShip = PopUpValue2
    lblBookSId.Caption = PopUpValue1
    txtBankName.text = PopUpValue3
    txtBankAdd.text = popupvalue4
    
    If RS.State = 1 Then RS.close
    RS.Open "select BookSeler,Add1,Add2,BankName,BankAdd,District,[State],city from QryBookSeller where BookSelerID='" & PopUpValue1 & "'", CON_blue
    If RS.EOF = False Then
    
       If LCase(RS!District) = LCase(RS!city) Then
          txtDist_ship = RS!city & ""
       Else
          
          If Not IsNull(RS!District) Then
             txtDist_ship = RS!city & " (" & RS!District & ")"
          Else
            txtDist_ship = RS!city & ""
          End If
          
       End If
       
       'txtDist_ship = RS!city & ""
       txtShipState.text = RS.Fields("State").value & ""
    End If
    
    cboTrans.SetFocus
    
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    popupvalue4 = ""
    popupvalue5 = ""
    
End If


Else
    'for school--------------------
    '------------------------------
    
    If RS.State = 1 Then RS.close
    If PopUpValue1 <> "" Then
        txtShip = PopUpValue2 & "," & popupvalue4
        lblBookSId.Caption = PopUpValue1
        PopUpValue1 = ""
        PopUpValue2 = ""
        PopUpValue3 = ""
        popupvalue4 = ""
        popupvalue5 = ""
    End If

End If
End Sub

Private Sub txtShip_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    If Check1_school.value = 0 Then
      
      tblNo = 50
      frmSearchItem.Show
      
    Else
       '''For School
      Screen.MousePointer = vbHourglass
      tblNo = 9
      frmSearchItem.Show
      Screen.MousePointer = vbDefault
    End If
End If

If KeyCode = 13 Then

If txtShip <> "" Then
   txtBankName.SetFocus
Else
   cboTrans.SetFocus
End If

End If

End Sub

Private Sub txtTransAdd_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtNarration.SetFocus
End Sub
Private Sub vs_Click()
If (vs.Col = 5 Or vs.Col = 7 Or vs.Col = 8 Or vs.Col = 11) Then
   vs.Editable = flexEDNone
Else
   vs.Editable = flexEDKbdMouse
End If
End Sub
Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   
   
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      
      If vs.Row >= 1 Then
      
       If add_ = False Then
        
            If Edit = True Then
               If (vs.TextMatrix(vs.RowSel, 13) <> "") Then
                con.Execute "delete from ORDERB where (INVOICENO=" & txtOrderNo & " and BOOKCODE='" & vs.TextMatrix(vs.RowSel, 1) & "' and sno=" & vs.TextMatrix(vs.RowSel, 13) & ")"
                vs.RemoveItem vs.Row
                vs.SetFocus
               Else
                 MsgBox "Plz Select Rows Again ...", vbCritical
               End If
               
             Else
              MsgBox "Plz Press Edit Button ...", vbCritical
            End If
        
       Else
               vs.RemoveItem vs.Row
               vs.SetFocus
       
       End If
         
           
    End If
    
   End If
   
End If
End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
      
      Dim net As Double
      Dim balQty As Integer
      Dim rs_ As New ADODB.Recordset
      
      If InStr(txtParty, "(EM)") > 0 Then
         party_type = "EM"
      Else
         party_type = "BP"
      End If
      
      nt = 0
      balQty = 0
      
      If KeyCode = 13 Then
      
      
      lblQtySpBalance.Caption = ""
      
      If vs.Col = 1 Then
          If RS.State = 1 Then RS.close
          RS.Open "select Bookcode,bookname,rate,DISCOUNT,GROUPCODE from BOOKS where bookcode='" & vs.TextMatrix(vs.RowSel, 1) & "'", con
          If RS.EOF = False Then
              
              
          '' Validation
          If Option1_sale.value = True Then
          
          If InStr(txtParty, "(SUNDRY)") = 0 Then
          
              If party_type = "EM" Then
                
                If (RS!groupcode = "EM" Or RS!groupcode = "SMD" Or RS!groupcode = "GI" Or RS!groupcode = "QB" Or RS!groupcode = "SMD1" Or RS!groupcode = "SMD2" Or RS!groupcode = "QT1" Or RS!groupcode = "QT2" Or RS!groupcode = "Q-CUET") Then
                   gp_ = "EM"
                End If
                
                If (party_type <> gp_) Then
                   MsgBox "Book is not valid for this customer ...", vbCritical
                   vs.SetFocus
                   Exit Sub
                End If
                
              Else
              
              If (RS!groupcode = "EM" Or RS!groupcode = "SMD" Or RS!groupcode = "GI" Or RS!groupcode = "QB" Or RS!groupcode = "SMD1" Or RS!groupcode = "SMD2" Or RS!groupcode = "QT1" Or RS!groupcode = "QT2" Or RS!groupcode = "Q-CUET") Then
                   MsgBox "Book is not valid for this customer ...", vbCritical
                   vs.SetFocus
                   Exit Sub
                End If
              End If
              
           End If
              
           End If
               
              vs.TextMatrix(vs.RowSel, 1) = UCase(vs.TextMatrix(vs.RowSel, 1))
              vs.TextMatrix(vs.RowSel, 2) = RS!Bookname
              vs.TextMatrix(vs.RowSel, 6) = RS!rate
              vs.TextMatrix(vs.RowSel, 0) = vs.Row
              vs.TextMatrix(vs.RowSel, 9) = RS!discount
              
          
              
              sendkeys "{right}"
              sendkeys "{right}"
 
          End If
           
      ElseIf vs.Col = 2 Then
          
          If RS.State = 1 Then RS.close
          RS.Open "select Bookcode,bookname,rate,DISCOUNT from BOOKS where bookname='" & vs.TextMatrix(vs.RowSel, 2) & "'", con
          If RS.EOF = False Then
              vs.TextMatrix(vs.RowSel, 1) = UCase(RS!Bookcode)
              vs.TextMatrix(vs.RowSel, 6) = RS!rate
              vs.TextMatrix(vs.RowSel, 0) = vs.Row
              sendkeys "{right}"
          End If
          
      ElseIf vs.Col = 3 Then
           
           
           net = (Val(vs.TextMatrix(vs.RowSel, 3)) * Val(vs.TextMatrix(vs.RowSel, 6)))
           If vs.TextMatrix(vs.RowSel, 9) <> "" Then
           vs.TextMatrix(vs.RowSel, 7) = net - Format(Round(net * (vs.TextMatrix(vs.RowSel, 9) / 100), 2), "0.00")
           End If
           sendkeys "{right}"
           

           
           Total
      ElseIf vs.Col = 4 Then
      
           qty = IIf(vs.TextMatrix(vs.RowSel, 3) = "", 0, vs.TextMatrix(vs.RowSel, 3))
           
           If Option1_sale.value = False Then
              qty = Val(qty) + Val(IIf(vs.TextMatrix(vs.RowSel, 4) = "", 0, vs.TextMatrix(vs.RowSel, 4)))
           End If
           
           net = (qty * Val(vs.TextMatrix(vs.RowSel, 6)))
           If vs.TextMatrix(vs.RowSel, 9) <> "" Then
           vs.TextMatrix(vs.RowSel, 7) = net - Format(Round(net * (vs.TextMatrix(vs.RowSel, 9) / 100), 2), "0.00")
           End If
           
           
           vs.TextMatrix(vs.RowSel, 11) = (qty * IIf(vs.TextMatrix(vs.RowSel, 6) = "", 0, vs.TextMatrix(vs.RowSel, 6)))
           
           Total
           
           If rs_.State = 1 Then rs_.close
           rs_.Open "select groupcode from books where bookcode='" & vs.TextMatrix(vs.RowSel, 1) & "'", con
           If rs_.EOF = False Then
           If rs_!groupcode = "BP" Then
            
           If Left(session, 4) >= 2018 Then
                If Val(vs.TextMatrix(vs.RowSel, 4)) > 0 Then
                   balQty = ReturnBalanceSpQty("" & txtBillQty.text, cmbAgentName.text, txtOrderNo.text, txtOrderDate.value)
                   lblQtySpBalance.Caption = "Balance Sp.Allotment Qty : " & balQty
                Else
                   balQty = ReturnBalanceSpQty("" & txtBillQty.text, cmbAgentName.text, txtOrderNo.text, txtOrderDate.value)
                   lblQtySpBalance.Caption = "Balance Sp.Allotment Qty : " & balQty
                End If
            End If
            
           End If
           End If
           
           sendkeys "{down}"
           sendkeys "{home}"
     
      End If
      
      End If
      
End Sub
Sub Total()
txtTotal = 0
txtTotQty = 0
txtBillQty = 0
txtGAmt.text = 0
txtOnlineAmt.text = 0



For I = 1 To vs.rows - 1
If vs.TextMatrix(I, 6) <> "" Then

txtTotal = Val(txtTotal) + Val(vs.TextMatrix(I, 7))
txtGAmt.text = Val(txtGAmt.text) + Val(vs.TextMatrix(I, 11))

txtTotQty = Val(txtTotQty) + Val(vs.TextMatrix(I, 3))
txtBillQty = Val(txtBillQty) + Val(vs.TextMatrix(I, 4))

End If

Next


txtTotal = Round(txtTotal, 0)


End Sub

Private Sub vs_Mail_DblClick()

If vs_Mail.TextMatrix(vs_Mail.RowSel, 0) <> "" Then

    refreshFld
    txtOrderNo = vs_Mail.TextMatrix(vs_Mail.RowSel, 0)
    cmdPrint_Click

End If

End Sub

Private Sub vs_Mail_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then

If vs_Mail.TextMatrix(vs_Mail.RowSel, 0) <> "" Then
    refreshFld
    txtOrderNo = vs_Mail.TextMatrix(vs_Mail.RowSel, 0)
    cmdPrint_Click
End If

End If

End Sub

Private Sub vs_Mail_SelChange()

If vs_Mail.Col = 5 Then
   vs_Mail.Editable = flexEDKbdMouse
Else
   vs_Mail.Editable = flexEDNone
End If

End Sub

Private Sub vs_SelChange()
   If vs.Col > 4 Then
      vs.Editable = flexEDNone
   Else
      vs.Editable = flexEDKbdMouse
   End If
End Sub
Private Sub vs1_Click()
   If vs1.Col = 5 Then
      vs1.Editable = flexEDKbdMouse
   Else
      vs1.Editable = flexEDNone
   End If
End Sub

Private Sub VSFlexGrid1_Click()

End Sub
