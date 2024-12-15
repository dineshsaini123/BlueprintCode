VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSaleRet 
   ClientHeight    =   10176
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   15852
   Icon            =   "frmSaleRet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10176
   ScaleWidth      =   15852
   Begin VB.CheckBox Check1_AllSchool 
      Caption         =   "ALL School"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   552
      Left            =   13860
      TabIndex        =   43
      Top             =   3780
      Width           =   1848
   End
   Begin VB.CheckBox Check1_gen 
      Caption         =   "General Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   552
      Left            =   13860
      TabIndex        =   42
      Top             =   3132
      Width           =   1848
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   14796
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImportExcel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fetch Excel Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   768
      Left            =   4176
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4320
      Width           =   1452
   End
   Begin VB.CommandButton cmdAddFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add E&xcel File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   768
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   4320
      Width           =   1272
   End
   Begin VB.TextBox txtpath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6840
      MaxLength       =   80
      TabIndex        =   39
      Top             =   4536
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.CommandButton cmdSearchSchool 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Search School"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   744
      Left            =   13824
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   576
      Width           =   1188
   End
   Begin VB.TextBox txtBiltyCharges 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1296
      TabIndex        =   36
      Top             =   1152
      Width           =   1320
   End
   Begin Crystal.CrystalReport cr 
      Left            =   396
      Top             =   3852
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CommandPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Height          =   612
      Left            =   5724
      Picture         =   "frmSaleRet.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5436
      Width           =   1068
   End
   Begin VB.CommandButton cmdremoveSchool 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete School"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   744
      Left            =   13824
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2112
      Width           =   1188
   End
   Begin VB.TextBox txtNoGaddi 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5445
      TabIndex        =   5
      Top             =   780
      Width           =   1812
   End
   Begin VB.TextBox txtBiltyNo 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1305
      TabIndex        =   3
      Top             =   780
      Width           =   1320
   End
   Begin VB.CommandButton cmdRepQty 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export to Excel"
      Height          =   612
      Left            =   6816
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5436
      Width           =   1068
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "+ Add New Party"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   768
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4320
      Width           =   1524
   End
   Begin VB.TextBox txtNQty 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5832
      TabIndex        =   29
      Top             =   4512
      Width           =   912
   End
   Begin VB.TextBox txtId 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   13320
      TabIndex        =   28
      Top             =   630
      Width           =   375
   End
   Begin VB.CommandButton cmdSearch 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   4632
      Picture         =   "frmSaleRet.frx":0BF0
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5436
      Width           =   1068
   End
   Begin VB.TextBox txtScid 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   12555
      TabIndex        =   25
      Top             =   630
      Width           =   735
   End
   Begin VB.TextBox txtSchName 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7560
      TabIndex        =   8
      Top             =   630
      Width           =   4965
   End
   Begin VB.TextBox txtRemarks 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1305
      TabIndex        =   6
      Top             =   1512
      Width           =   5952
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12420
      TabIndex        =   19
      Top             =   4185
      Width           =   1005
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
      Height          =   612
      Left            =   2415
      Picture         =   "frmSaleRet.frx":17D4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5436
      Width           =   1068
   End
   Begin VB.CommandButton close 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   612
      Left            =   7944
      Picture         =   "frmSaleRet.frx":1BE1
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5436
      Width           =   1068
   End
   Begin VB.CommandButton Abandon 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   612
      Left            =   180
      Picture         =   "frmSaleRet.frx":27C5
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5436
      Width           =   1068
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Height          =   612
      Left            =   3525
      Picture         =   "frmSaleRet.frx":33A9
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5436
      Width           =   1068
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "S&ave"
      Height          =   612
      Left            =   1275
      Picture         =   "frmSaleRet.frx":3F8D
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5436
      Width           =   1068
   End
   Begin VB.TextBox txtEntryNo 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1305
      TabIndex        =   0
      Top             =   45
      Width           =   1320
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "+ &Add School"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   744
      Left            =   13824
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1320
      Width           =   1188
   End
   Begin VB.ComboBox cboParty 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   336
      Left            =   1305
      Locked          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Top             =   420
      Width           =   5952
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3045
      Left            =   7560
      TabIndex        =   10
      Top             =   1035
      Width           =   6135
      _cx             =   10821
      _cy             =   5371
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
      BackColorFixed  =   12640511
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   8404992
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
      Rows            =   3
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSaleRet.frx":4B71
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
   End
   Begin MSComCtl2.DTPicker dates 
      Height          =   330
      Left            =   2670
      TabIndex        =   1
      Top             =   45
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   572
      _Version        =   393216
      Format          =   508100609
      CurrentDate     =   39127
   End
   Begin VSFlex7Ctl.VSFlexGrid vs_1 
      Height          =   3720
      Left            =   48
      TabIndex        =   27
      Top             =   6360
      Width           =   15708
      _cx             =   27707
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   8404992
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
      Rows            =   3
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   400
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSaleRet.frx":4BD7
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
   End
   Begin VSFlex7Ctl.VSFlexGrid vs1 
      Height          =   2340
      Left            =   1308
      TabIndex        =   7
      Top             =   1884
      Width           =   5952
      _cx             =   10499
      _cy             =   4128
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
      BackColorFixed  =   12640511
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   8404992
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
      Rows            =   3
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSaleRet.frx":4C3D
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
   End
   Begin MSMask.MaskEdBox txtBiltyDate 
      Height          =   312
      Left            =   2652
      TabIndex        =   4
      Top             =   816
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   550
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Freight :"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   348
      Left            =   36
      TabIndex        =   37
      Top             =   1152
      Width           =   1476
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Height          =   60
      Left            =   -48
      TabIndex        =   32
      Top             =   6168
      Width           =   15816
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F4 Key To Delete Grid Row"
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
      Left            =   7560
      TabIndex        =   30
      Top             =   4185
      Width           =   2955
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "School Name :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7560
      TabIndex        =   24
      Top             =   360
      Width           =   1125
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "No of Gaddi :"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   4272
      TabIndex        =   23
      Top             =   780
      Width           =   1176
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   48
      TabIndex        =   22
      Top             =   1512
      Width           =   1212
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Bilty No && Date :"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   348
      Left            =   48
      TabIndex        =   21
      Top             =   780
      Width           =   1476
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   11790
      TabIndex        =   20
      Top             =   4185
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "EntryNo/Date :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   18
      Top             =   90
      Width           =   1200
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 for Search Party..."
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   192
      Left            =   4236
      TabIndex        =   17
      Top             =   180
      Width           =   3132
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   48
      TabIndex        =   16
      Top             =   420
      Width           =   1260
   End
End
Attribute VB_Name = "frmSaleRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit As Boolean
Dim scid, scname As String

Private Sub ABANDON_Click()
Edit = False
Check1_AllSchool.value = 0
cmdEdit_4.Enabled = True

txtTotal.text = ""

vs.Clear
vs.rows = 2
setWidth

txtBiltyCharges.text = 0
Total

vs_1.Clear

fillGrid
cmdremoveSchool.Enabled = False

cmdEdit_4.Enabled = False
cmdDel.Enabled = False
cmdSave.Enabled = True

End Sub
Sub maxId()

If rs1.State = 1 Then rs1.close
rs1.Open "select max(convert(int,EntryNo)) from schoolWiseBookReturn", con, adOpenDynamic, adLockOptimistic
If Not IsNull(rs1(0)) Then
   txtEntryNo.text = rs1(0) + 1
Else
   txtEntryNo.text = 1
End If
   
End Sub
Private Sub cboParty_GotFocus()

If PopUpValue3 <> "" Then
   cboParty.text = PopUpValue3
   
'   cboschool.Clear
'
'   If RS.State = 1 Then RS.close
'   RS.Open "SELECT ScName,ScID FROM INVOICEA where (subledger='" & cboParty & "' and (len(ScID)>0 and len(Scname)>0)) group by ScName,ScID", con
'   While RS.EOF = False
'        cboschool.AddItem RS(0) & ":" & RS(1)
'        RS.MoveNext
'   Wend
   
   fillGrid
   Total
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
End If

End Sub
Private Sub cboParty_KeyDown(KeyCode As Integer, Shift As Integer)
    
If KeyCode = 113 Then

    searchType = "party"
    lblCr = "dr"
    value = "select distinct(Party),Code,subledger from SLEDGER where gledger='SUNDRY DEBTORS' and " & stringyear & "  order by party"
    popuplist_client value, CCON
    set_focus = True

End If

If KeyCode = 13 Then
   txtBiltyNo.SetFocus
End If
    
End Sub

Private Sub Check1_AllSchool_Click()
If (Check1_AllSchool.value = 1) Then
   'txtSchName.text = ""
   'txtScId.text = ""
   Check1_gen.value = 0
Else
   'txtSchName.text = ""
   'txtScId.text = ""
   Check1_gen.value = 0

End If
End Sub

Private Sub Check1_gen_Click()
If (Check1_gen.value = 1) Then
   txtSchName.text = "GENERAL ENTRY"
   txtScId.text = "G99999"
   Check1_AllSchool.value = 0
Else
   txtSchName.text = ""
   txtScId.text = ""
   Check1_AllSchool.value = 0

End If
End Sub

Private Sub close_Click()
Unload Me
End Sub
Sub fillGrid()
'Dim rsfill As New ADODB.Recordset
'
'Set rsfill = New ADODB.Recordset
'
'If cboParty.Text <> "" Then
'    rsfill.Open "SELECT EntryNo,EntDate,PartyName,ScId,ScName,SN,BooKCode,RetQty " & _
'    " from schoolWiseBookReturn where partyName='" & cboParty & "' order by EntryNo,SN", con, adOpenDynamic, adLockOptimistic
'    Set vs1.DataSource = rsfill
'
'    vs1.FormatString = "EntryNo|EntDate|PartyName|ScId|ScName|SN|BooKCode|RetQty"
'    vs1.ColWidth(0) = 1000
'    vs1.ColWidth(1) = 1000
'    vs1.ColWidth(2) = 4200
'    vs1.ColWidth(3) = 1000
'    vs1.ColWidth(4) = 4200
'    vs1.ColWidth(5) = 800
'    vs1.ColWidth(6) = 1000
'    vs1.ColWidth(7) = 1000
'
'
'End If
End Sub

Private Sub cmdAdd_Click()
Edit = False

cmdEdit_4.Enabled = True
txtEntryNo.text = ""

txtTotal.text = ""

txtBiltyNo = ""
txtNoGaddi = ""
txtRemarks = ""
txtNQty.text = ""
cboParty.text = ""
txtBiltyCharges.text = "0"

vs.Clear
vs1.Clear

setWidth

Total
maxId


vs_1.Clear


fillGrid

cmdEdit_4.Enabled = False
cmdDel.Enabled = False
cmdSave.Enabled = True
cboParty.SetFocus

End Sub

Private Sub cmdAddFile_Click()

cd.ShowOpen
txtpath.text = cd.filename




End Sub

Private Sub cmdDel_Click()
If MsgBox("Want to Delete ? ", vbQuestion + vbYesNo) = vbYes Then
   If cboParty <> "" Then
   
   con.Execute "delete from schoolWiseBookReturn where (PartyName='" & cboParty & "' and EntryNo=" & txtEntryNo & ")"
   con.Execute "delete from SchoolWiseBookReturnDet where (EntryNo=" & txtEntryNo & ")"
   con.Execute "delete from PartyWiseBookReturnDet  where (EntryNo=" & txtEntryNo & ")"
   
   
   ABANDON_Click
   End If
End If
End Sub

Private Sub cmdEdit_4_Click()
cmdDel.Enabled = True
cmdEdit_4.Enabled = False
cmdSave.Enabled = True
Edit = True
End Sub
Private Sub cmdImportExcel_Click()

Dim sconn As String
Dim I As Integer

sFile = Me.txtpath.text
''sconn = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & sFile

sconn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & sFile & "';Extended Properties=""Excel 12.0;HDR=YES;IMEX=1;"";"




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

rs_em.Open "select groupcode,bookcode,rate from books", con, adOpenDynamic, adLockOptimistic

If RS.State = 1 Then RS.close
RS.Open "SELECT * FROM [sheet1$]", sconn
While RS.EOF = False

  


v1 = IIf(IsNull(RS(2)), 0, RS(2))

If (v1 > 0) Then
    

rs_em.MoveFirst
rs_em.Find "bookcode='" & UCase(RS(0)) & "'"
If rs_em.EOF = False Then
                 
    vs1.TextMatrix(k1, 0) = k1
    vs1.TextMatrix(k1, 1) = UCase(RS(0))
    vs1.TextMatrix(k1, 2) = UCase(rs_em(2))
    vs1.TextMatrix(k1, 3) = UCase(RS(2))
    vs1.rows = vs1.rows + 1
    k1 = k1 + 1
End If

End If

RS.MoveNext

Wend

Total



MsgBox "Data import Successfully", vbInformation


End Sub

Private Sub cmdremoveSchool_Click()
If MsgBox("Want to Delete School ? ", vbQuestion + vbYesNo) = vbYes Then
   If txtScId <> "" Then
   con.Execute "delete from SchoolWiseBookReturnDet where (EntryNo=" & txtEntryNo & " and scid='" & txtScId & "')"
   txtSchName.text = ""
   txtScId.text = ""
   txtId.text = ""
   txtTotal.text = ""
   vs.Clear
   
   searchData
   
   End If
End If
End Sub

Private Sub cmdRepQty_Click()
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim rs_1 As New ADODB.Recordset
Dim rs_2 As New ADODB.Recordset
Dim str_ As String

On Error GoTo err:





If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

xl.Visible = True
Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add

Dim c, r As Long
Dim Q1, q2, J As Double

Dim b1 As Boolean

b1 = False


c = 1
r = 1





row_ = 1
    col_ = 1
   
    xl.Columns("A:H").ColumnWidth = 12
    J = 2
    xlSheet.Cells(1, 1).value = "Bilty Return Status "
    
    For I = 0 To vs_1.rows - 1
        For J = 0 To vs_1.Cols - 1
               xlSheet.Cells(row_, col_).value = vs_1.TextMatrix(I, J)
              col_ = col_ + 1
        Next
        row_ = row_ + 1
        col_ = 1
    Next
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
err:
    
    
End Sub
Sub saveData()


If Edit = False Then

If RS.State = 1 Then RS.close
RS.Open "select * from schoolWiseBookReturn where (EntryNo='" & txtEntryNo.text & "')", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
   maxId
End If

End If



If RS.State = 1 Then RS.close
RS.Open "select * from schoolWiseBookReturn where (PartyName='" & cboParty.text & "' and EntryNo='" & txtEntryNo.text & "')", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
RS.AddNew
End If

RS!entryNo = txtEntryNo.text
RS!EntDate = Format(dates.value, "dd/MM/yyyy")
RS!partyname = Trim(cboParty.text)
RS!biltyno = Trim(txtBiltyNo.text)

If IsDate(txtBiltyDate.text) Then
RS!BILTYDATE = Format(txtBiltyDate.text, "dd/MM/yyyy")
End If

RS!noofgaddi = Trim(txtNoGaddi.text)
RS!remarks = Trim(txtRemarks.text)
If (txtBiltyCharges.text <> "") Then
    RS!BiltyCharges = txtBiltyCharges.text
Else
    RS!BiltyCharges = 0
End If

RS.update

'=============================================================
 For k1 = 1 To vs1.rows - 1
 If vs1.TextMatrix(k1, 1) <> "" Then
    If RS.State = 1 Then RS.close
    RS.Open "select * from PartyWiseBookReturnDet where (EntryNo='" & txtEntryNo.text & "' and BCode='" & vs1.TextMatrix(k1, 1) & "' and Price='" & Trim(vs1.TextMatrix(k1, 2)) & "' and SN=" & Val(vs1.TextMatrix(k1, 0)) & ")", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = True Then
       RS.AddNew
    End If
    RS!entryNo = txtEntryNo.text
    RS!sn = Trim(vs1.TextMatrix(k1, 0))
    RS!bcode = Trim(vs1.TextMatrix(k1, 1))
    RS!Price = Trim(vs1.TextMatrix(k1, 2))
    RS!qty = Trim(vs1.TextMatrix(k1, 3))
    RS.update
 End If
Next

'=============================================================

For k1 = 1 To vs.rows - 1
 If vs.TextMatrix(k1, 1) <> "" Then
    If RS.State = 1 Then RS.close
    RS.Open "select * from SchoolWiseBookReturnDet where (scid='" & txtScId.text & "' " & _
    " and EntryNo='" & txtEntryNo.text & "' and BCode='" & vs.TextMatrix(k1, 1) & "' " & _
    " and price=" & vs.TextMatrix(k1, 2) & ")", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = True Then
       RS.AddNew
    End If
    RS!entryNo = txtEntryNo.text
    RS!scname = Trim(txtSchName.text)
    RS!scid = Trim(txtScId.text)
    RS!sn = Trim(vs.TextMatrix(k1, 0))
    RS!bcode = Trim(vs.TextMatrix(k1, 1))
    RS!Price = Trim(vs.TextMatrix(k1, 2))
    RS!qty = Trim(vs.TextMatrix(k1, 3))
    
    If txtId <> "" Then
         RS!id = Trim(txtId)
    End If
    
    RS.update
    
 End If
 
Next


End Sub
Private Sub cmdSave_Click()


saveData


cmdEdit_4.Enabled = True
cmdDel.Enabled = False
cmdSave.Enabled = False
Abandon.Enabled = True

searchData

End Sub

Private Sub cmdSearch_Click()
value = "Select EntryNo,PartyName,EntDate from schoolWiseBookReturn order by convert(int,EntryNo)"
popuplistModel10 value, con
End Sub
Sub searchData()

Screen.MousePointer = vbHourglass

If RS.State = 1 Then RS.close
RS.Open "select * from schoolWiseBookReturn where EntryNo='" & txtEntryNo.text & "'", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
   txtEntryNo.text = RS!entryNo
   dates.value = RS!EntDate
   cboParty.text = RS!partyname
   txtBiltyNo.text = RS!biltyno
   If IsDate(RS!BILTYDATE) Then
   txtBiltyDate.text = RS!BILTYDATE
   End If
   
   txtNoGaddi.text = RS!noofgaddi
   txtRemarks.text = RS!remarks
   txtBiltyCharges = RS!BiltyCharges & ""
End If

'===============================================================
vs1.Clear
k2 = 1
vs1.rows = 2

If RS.State = 1 Then RS.close
RS.Open "SELECT SN,BCode,Price,Qty from PartyWiseBookReturnDet where (EntryNo='" & txtEntryNo.text & "') order by BCode,SN", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
vs1.TextMatrix(k2, 0) = RS!sn

vs1.TextMatrix(k2, 1) = RS!bcode
vs1.TextMatrix(k2, 2) = RS!Price
vs1.TextMatrix(k2, 3) = RS!qty
vs1.rows = vs1.rows + 1
k2 = k2 + 1

RS.MoveNext
Wend


Total
setWidth

fillDataSchoolWise

Screen.MousePointer = vbDefault

End Sub
Sub fillDataSchoolWise()

vs_1.Clear

k2 = 1

vs_1.Cols = 4
vs_1.rows = vs_1.rows + 2

vs_1.Clear
k2 = 1
vs_1.rows = 2

If RS.State = 1 Then RS.close
RS.Open "SELECT SN,BCode,Price,Qty from PartyWiseBookReturnDet where (EntryNo='" & txtEntryNo.text & "') order by SN", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False

vs_1.TextMatrix(k2, 0) = RS!sn
vs_1.TextMatrix(k2, 1) = RS!bcode
vs_1.TextMatrix(k2, 2) = RS!Price
vs_1.TextMatrix(k2, 3) = RS!qty

vs_1.rows = vs1.rows + 1
k2 = k2 + 1

RS.MoveNext
Wend

'===================================================
Dim rss_ As New ADODB.Recordset

Set rss_ = New ADODB.Recordset


If RS.State = 1 Then RS.close
RS.Open "select id,scid,scname from SchoolWiseBookReturnDet where entryNo=" & txtEntryNo.text & " group by id,scid,scname", con
While RS.EOF = False

   
If rs1.State = 1 Then rs1.close
rs1.Open "select scid,scname,bcode,Qty,id,Price from SchoolWiseBookReturnDet where entryNo=" & txtEntryNo.text & " and scid='" & RS!scid & "'  order by sn", con
If rs1.EOF = False Then
   vs_1.Cols = vs_1.Cols + 1
   vs_1.TextMatrix(0, vs_1.Cols - 1) = rs1!scid & ":" & rs1!scname
End If
    


For k1 = 1 To vs_1.rows - 1
    
    
     
    
'    rs1.MoveFirst
'    rs1.Find "bcode='" & vs_1.TextMatrix(k1, 1) & "'"
'    If rs1.EOF = False Then
'     ''   vs_1.TextMatrix(k1, vs_1.Cols - 1) = rs1!qty
'    End If
    
'    If (vs_1.TextMatrix(k1, 1) = "VP8") Then
'       MsgBox ("sss")
'    End If
    
    
    
    If rs1.EOF = False Then
    
    If vs_1.TextMatrix(k1, 2) <> "" Then
    If rss_.State = 1 Then rss_.close
        rss_.Open "select qty from SchoolWiseBookReturnDet where (entryNo=" & txtEntryNo.text & " and scid='" & rs1!scid & "' and bcode='" & vs_1.TextMatrix(k1, 1) & "' and Price='" & vs_1.TextMatrix(k1, 2) & "')", con
        If rss_.EOF = False Then
           vs_1.TextMatrix(k1, vs_1.Cols - 1) = rss_!qty
        End If
    End If
    
    End If

    
Next



RS.MoveNext
Wend

'===================================================
vs_1.Cols = vs_1.Cols + 1
vs_1.TextMatrix(0, vs_1.Cols - 1) = "Total"

For k1 = 3 To vs_1.Cols - 1
t = 0

For kk1 = 1 To vs_1.rows - 1
    If vs_1.TextMatrix(kk1, k1) <> "" Then
       t = t + IIf(vs_1.TextMatrix(kk1, k1) = "", 0, vs_1.TextMatrix(kk1, k1))
    End If
Next
    vs_1.TextMatrix(kk1 - 2, k1) = t
    vs_1.Cell(flexcpBackColor, kk1 - 2, k1) = &HC0E0FF
Next

col_ = k1
row_ = kk1

r = 0
For kk1 = 1 To vs_1.rows - 2
t = 0
For k1 = 4 To vs_1.Cols - 2
    If vs_1.TextMatrix(kk1, k1) <> "" Then
       t = t + IIf(vs_1.TextMatrix(kk1, k1) = "", 0, vs_1.TextMatrix(kk1, k1))
    End If
Next
    
    vs_1.TextMatrix(kk1, k1) = t
    r = r + IIf(vs_1.TextMatrix(kk1, k1) = "", 0, vs_1.TextMatrix(kk1, k1))
Next

'vs_1.TextMatrix(row_ - 1, k1) = r
 

'===================================================
vs_1.FormatString = "SN|Code|Price|Ret.Qty"


vs_1.ColWidth(0) = 600
vs_1.ColWidth(1) = 680
vs_1.ColWidth(2) = 700
vs_1.ColWidth(3) = 700


End Sub

Private Sub cmdSearch_GotFocus()
If PopUpValue1 <> "" Then
   txtEntryNo = PopUpValue1
   cboParty = PopUpValue2
   searchData
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
End If
End Sub



Private Sub cmdSearchSchool_Click()
    
    value = "SELECT ScName,ScID FROM billingSchoolQry where len(ScName)>0 and SUBLEDGER='" & cboParty & "' group by ScName,ScID"
    popuplistModel10 value, con

End Sub
Private Sub cmdSearchSchool_GotFocus()
    
    
    If PopUpValue1 <> "" Then

        txtScId = PopUpValue2
        txtSchName = PopUpValue1
        
        PopUpValue1 = ""
        PopUpValue2 = ""

        vs.WordWrap = True

    End If

    
End Sub
Sub createId()

If rs1.State = 1 Then rs1.close
rs1.Open "select max(id) from SchoolWiseBookReturnDet where entryNo=" & txtEntryNo.text & "", con, adOpenDynamic, adLockOptimistic
If Not IsNull(rs1(0)) Then
    txtId.text = rs1(0) + 1
Else
    txtId.text = 1
End If

End Sub

Private Sub Command4_Click()
   
If (txtScId.text = "") Then
    MsgBox "First Search School Name ? ", vbCritical
    Exit Sub
End If


If MsgBox("Want to Add School ? ", vbQuestion + vbYesNo) = vbYes Then
   
    
    
   Check1_AllSchool.value = 0
   Edit = True
   saveData
   searchData
   
   updateId
   
   vs.Clear
   setWidth
   
   txtSchName = ""
   txtScId.text = ""
   txtId.text = ""
   txtTotal.text = ""
   
   
   
   
End If

   
End Sub
Private Sub Command4_LostFocus()
If rs1.State = 1 Then rs1.close
rs1.Open "select max(id) from SchoolWiseBookReturnDet where entryNo=" & txtEntryNo.text & "", con, adOpenDynamic, adLockOptimistic
If Not IsNull(rs1(0)) Then
   txtId.text = rs1(0) + 1
Else
   txtId.text = 1
End If
   
End Sub
Sub updateId()

Dim id As Integer

id = 1

If rs1.State = 1 Then rs1.close
rs1.Open "select scid,SCName FROM SchoolWiseBookReturnDet where EntryNo=" & txtEntryNo.text & " group by scid,SCName", con
While rs1.EOF = False

con.Execute "update SchoolWiseBookReturnDet set id=" & id & "  where (EntryNo='" & txtEntryNo.text & "' and scid='" & rs1(0) & "')"

id = id + 1
rs1.MoveNext
Wend


End Sub

Private Sub CommandPrint_Click()

Dim bcode_ As String
Dim pname_ As String
Dim D As Double

D = 0

If rs1.State = 1 Then rs1.close
rs1.Open "SELECT a.ScId,a.BCode,substring(b.PartyName,1,5) as Code,a.EntryNo " & _
"FROM SchoolWiseBookReturnDet as a inner join SchoolWiseBookReturn as b on (a.EntryNo=b.EntryNo) where a.EntryNo=" & txtEntryNo.text
While rs1.EOF = False

D = ReturnDiscountNew(rs1(1), rs1(2), rs1(0))
con.Execute "update SchoolWiseBookReturnDet set billingdis=" & D & "  where (EntryNo='" & txtEntryNo.text & "' and bcode='" & rs1(1) & "' and scid='" & rs1(0) & "')"

rs1.MoveNext
Wend


con.Execute "update a set a.AgentName=b.AgentName FROM SchoolWiseBookReturnDet " & _
"as a inner join INVOICEA as b on (a.ScId =b.ScID) Where a.EntryNo = " & txtEntryNo.text


Dim CON_next As New ADODB.Connection


If rs1.State = 1 Then rs1.close
rs1.Open "select fromDate,toDate,fyear,Current_Next,DataBase,fromDateSRet,toDateSRet from turnOverDis where NotCreated='y' and Current_Next='next' order by fyear", CCON
If rs1.EOF = False Then
'------Fatch Data From Next Session--------------------------------------'
    db_ = rs1!Database
    Set CON_next = New ADODB.Connection
    If LCase(server_) = "server" Then
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & db_ & "; UID=" & sql_user & "; PWD=" & sql_pass
       CON_next.Open
    End If

    If RS.State = 1 Then RS.close
    RS.Open "select ScId  from SchoolWiseBookReturnDet Where EntryNo = " & txtEntryNo.text & " group by ScId", con
    
    While RS.EOF = False
    
        If rs1.State = 1 Then rs1.close
        rs1.Open "select AgentName from INVOICEA where ScId='" & RS!scid & "' group by AgentName", CON_next
        
        If rs1.EOF = False Then
        
            con.Execute "update SchoolWiseBookReturnDet set AgentName='" & rs1!agentname & "'  Where ScId='" & RS!scid & "' And EntryNo = " & txtEntryNo.text
        
        End If
    
    RS.MoveNext
    Wend

End If




updateId



DSNNew
  
cr.Reset
cr.ReportFileName = rptPath & "/SchoolWiseSaleRet.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
cr.ReplaceSelectionFormula "{schoolWiseBookReturn.entryno}='" & txtEntryNo.text & "'"
''cr.Formulas(0) = "unit_='" & "(A division of Chitra Prakashan (I) Pvt.Ltd.)" & "'"
''cr.Formulas(1) = "address='" & Text1.Text & "'"
cr.WindowShowPrintSetupBtn = True
cr.WindowState = crptMaximized
cr.Action = 1



End Sub

Private Sub dates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cboParty.SetFocus
End Sub

Private Sub Form_Load()

Me.top = 50
Me.Left = 50
Me.Width = 15900
Me.Height = 10800
BackColorFrom Me

setWidth
Edit = False

Screen.MousePointer = vbHourglass
con.Execute "exec tmpdataForSale "
Screen.MousePointer = vbDefault

dates.value = Format(Date, "dd/MM/yyyy")
maxId

End Sub
Sub setWidth()


vs.Cols = 4

vs.FormatString = "SN.|BCode|Price|Qty"

vs.ColWidth(0) = 1200
vs.ColWidth(1) = 1500
vs.ColWidth(2) = 1500
vs.ColWidth(3) = 1400

vs1.Cols = 4

vs1.FormatString = "SN.|BCode|Price|Qty"

vs1.ColWidth(0) = 1100
vs1.ColWidth(1) = 1400
vs1.ColWidth(2) = 1400
vs1.ColWidth(3) = 1400
   
End Sub
Sub Total()
txtTotal = 0
txtNQty = 0

For k1 = 1 To vs.rows - 1
  If vs.TextMatrix(k1, 1) <> "" Then
     txtTotal.text = Val(txtTotal.text) + IIf(vs.TextMatrix(k1, 3) = "", 0, vs.TextMatrix(k1, 3))
  End If
Next

For k1 = 1 To vs1.rows - 1
  If vs1.TextMatrix(k1, 1) <> "" Then
     txtNQty.text = Val(txtNQty.text) + IIf(vs1.TextMatrix(k1, 3) = "", 0, vs1.TextMatrix(k1, 3))
  End If
Next


End Sub
Private Sub save_Click()

End Sub

Private Sub txtBiltyDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub


Private Sub txtBiltyCharges_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   txtRemarks.SetFocus
End If
End Sub

Private Sub txtBiltyDate_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   txtNoGaddi.SetFocus
End If
End Sub

Private Sub txtBiltyNo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   txtBiltyDate.SetFocus
End If
End Sub

Private Sub txtEntryNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   dates.SetFocus
   searchData
End If
End Sub

Private Sub txtNoGaddi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   txtBiltyCharges.SetFocus
End If
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   vs1.Col = 1
   vs1.SetFocus
End If
End Sub

Private Sub vs_1_DblClick()

If (vs_1.ColSel > 2) Then

Dim scid_ As String
scid_ = ""
k1 = 1

If InStr(vs_1.TextMatrix(0, vs_1.ColSel), ":") = 0 Then Exit Sub

scid_ = Mid(vs_1.TextMatrix(0, vs_1.ColSel), 1, InStr(vs_1.TextMatrix(0, vs_1.ColSel), ":") - 1)

vs.rows = 2
If RS.State = 1 Then RS.close
RS.Open "select SN,bcode,price,Qty,scid,scname,id  from SchoolWiseBookReturnDet where (entryNo='" & txtEntryNo & "' and scid='" & scid_ & "') order by sn", con

If RS.EOF = False Then
cmdremoveSchool.Enabled = True
End If

While RS.EOF = False



txtScId = RS!scid
txtSchName = RS!scname
txtId = RS!id & ""

vs.TextMatrix(k1, 0) = RS!sn
vs.TextMatrix(k1, 1) = RS(1)
vs.TextMatrix(k1, 2) = RS(2)
vs.TextMatrix(k1, 3) = RS(3)
k1 = k1 + 1
vs.rows = vs.rows + 1
RS.MoveNext
Wend
Total

End If

End Sub
Private Sub vs_1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   txtRemarks.SetFocus
End If
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 115 Then

    If vs1.TextMatrix(vs.RowSel, 1) <> "" Then
       If MsgBox("Want to Delete ?", vbQuestion + vbYesNo) = vbYes Then
          con.Execute "delete from SchoolWiseBookReturnDet where (entryno=" & txtEntryNo & " and bcode='" & vs.TextMatrix(vs.RowSel, 1) & "' and ScId='" & txtScId.text & "')"
          vs.RemoveItem (vs.RowSel)
          Total
       End If
    End If

End If
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then

If (vs.Col = 1 Or vs.Col = 2) Then
    
    


Dim str_date As String
str_date = "(invoicedate>=convert(smalldatetime,'" & financialyear_Fdate & "',103) and invoicedate<=convert(smalldatetime,'" & financialyear_Tdate & "',103))"
  
  
   If (Check1_gen.value = 0 And Check1_AllSchool.value = 0) Then
   
        If RS.State = 1 Then RS.close
        RS.Open "select bookcode from BookAndSchoolWiseSaleRet where (bookcode='" & vs.TextMatrix(vs.RowSel, 1) & "' and scid='" & txtScId & "' and " & str_date & ")", con
        If RS.EOF = True Then
           MsgBox "Sale is not Exist in this School in this Date Rage : " & financialyear_Fdate & " to " & financialyear_Tdate, vbCritical
           vs.SetFocus
           Exit Sub
        End If
   
   End If
   
   
   
   


   If RS.State = 1 Then RS.close
   If (vs.TextMatrix(vs.RowSel, 2) = "") Then
       RS.Open "select bcode,price  from PartyWiseBookReturnDet where (EntryNo=" & txtEntryNo & " and bcode='" & vs.TextMatrix(vs.RowSel, 1) & "')", con
   Else
       RS.Open "select bcode,price  from PartyWiseBookReturnDet where (EntryNo=" & txtEntryNo & " and bcode='" & vs.TextMatrix(vs.RowSel, 1) & "' and price=" & vs.TextMatrix(vs.RowSel, 2) & ")", con
   End If
   
   If RS.EOF = False Then
      vs.TextMatrix(vs.RowSel, 0) = vs.Row
      vs.TextMatrix(vs.RowSel, 1) = RS(0)
      vs.TextMatrix(vs.RowSel, 2) = RS(1)
      sendkeys "{right}"
      sendkeys "{right}"
   End If
ElseIf vs.Col = 3 Then

    If RS.State = 1 Then RS.close
    RS.Open "select convert(int,qty) from PartyWiseBookReturnDet where (EntryNo=" & txtEntryNo & " and bcode='" & vs.TextMatrix(vs.RowSel, 1) & "')", con
    If Not IsNull(RS(0)) Then
       If Val(IIf(vs.TextMatrix(vs.RowSel, 3) = "", 0, vs.TextMatrix(vs.RowSel, 3))) > RS(0) Then
          MsgBox "Quantity must be less from the given Quantity....", vbCritical
          vs.SetFocus
          Exit Sub
       End If
    End If
    


    sendkeys "{home}"
    sendkeys "{down}"
    Total
    vs.rows = vs.rows + 1
End If

End If

End Sub
Private Sub vs1_Click()
 
'k1 = 1
'
'vs.Rows = 2
'vs.Clear
'
'If RS.State = 1 Then RS.close
'RS.Open "select EntryNo from schoolWiseBookReturn where (EntryNo='" & vs1.TextMatrix(vs1.RowSel, 0) & "' and PartyName='" & cboParty.Text & "')", con
'If RS.EOF = False Then
'    cmdDel.Enabled = False
'    cmdEdit_4.Enabled = True
'    cmdSave.Enabled = False
'    txtEntryNo = RS!EntryNo
'End If
'While RS.EOF = False
'
'   vs.TextMatrix(k1, 0) = RS!sn
'   vs.TextMatrix(k1, 1) = RS!Bookcode
'
'   If rs1.State = 1 Then rs1.close
'   rs1.Open "select bookname from books where bookcode='" & vs.TextMatrix(k1, 1) & "'", con
'   If rs1.EOF = False Then
'   vs.TextMatrix(k1, 2) = rs1(0)
'   End If
'   vs.TextMatrix(k1, 3) = RS!RetQty
'   k1 = k1 + 1
'   vs.Rows = vs.Rows + 1
'   RS.MoveNext
'Wend
'
''cboschool.Text = vs1.TextMatrix(vs1.RowSel, 4) & ":" & vs1.TextMatrix(vs1.RowSel, 3)
'
'setWidth
'total

End Sub
Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 115 Then

    If vs1.TextMatrix(vs.RowSel, 1) <> "" Then
       If MsgBox("Want to Delete ?", vbQuestion + vbYesNo) = vbYes Then
          con.Execute "delete from PartyWiseBookReturnDet where entryno=" & txtEntryNo & " and bcode='" & vs1.TextMatrix(vs1.RowSel, 1) & "'"
          vs1.RemoveItem (vs1.RowSel)
          Total
       End If
    End If

End If

End Sub

Private Sub vs1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

If vs1.Col = 1 Then
   If RS.State = 1 Then RS.close
   RS.Open "select bookname,bookcode,rate  from books where bookcode='" & vs1.TextMatrix(vs1.RowSel, 1) & "'", con
   If RS.EOF = False Then
      vs1.TextMatrix(vs1.RowSel, 0) = vs1.Row
      vs1.TextMatrix(vs1.RowSel, 1) = RS(1)
      vs1.TextMatrix(vs1.RowSel, 2) = RS(2)
      sendkeys "{right}"
      sendkeys "{right}"
   End If
ElseIf vs1.Col = 2 Then
   sendkeys "{right}"
ElseIf vs1.Col = 3 Then
    sendkeys "{home}"
    sendkeys "{down}"
    Total
    vs1.rows = vs1.rows + 1
End If

End If
End Sub
