VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesAdjustment 
   ClientHeight    =   10008
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   14460
   Icon            =   "frmSalesAdjustment.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10008
   ScaleWidth      =   14460
   Begin VB.CheckBox Check1_partyTerms 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Party Terms Not Required"
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
      Left            =   5148
      TabIndex        =   57
      Top             =   8424
      Width           =   2592
   End
   Begin VB.TextBox txtManually 
      Height          =   348
      Left            =   3852
      MaxLength       =   250
      TabIndex        =   56
      Top             =   7632
      Width           =   3804
   End
   Begin VB.CheckBox Check1_manullay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Manually Change Adj."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   792
      TabIndex        =   55
      Top             =   7632
      Width           =   3060
   End
   Begin VB.CommandButton cmdView1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View-2"
      Height          =   255
      Left            =   8595
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   1170
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.CommandButton cmdVew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "V&iew-1"
      Height          =   255
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   1170
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CheckBox Check1_manual 
      BackColor       =   &H8000000E&
      Caption         =   "Adj.No Enter Manually"
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
      Left            =   1215
      TabIndex        =   51
      Top             =   405
      Width           =   2040
   End
   Begin VB.CheckBox Check2_AppDet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Approval Details"
      Height          =   390
      Left            =   6255
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   855
      Width           =   1695
   End
   Begin VB.CommandButton cmdEditDet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Adjustment Edit Details"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3552
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   8172
      Width           =   1512
   End
   Begin VB.CommandButton cmdRepQty 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Export to Excel"
      Height          =   555
      Left            =   7245
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   8988
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CheckBox Check1_PendingCr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Pending Bill In (Credit Note Item)"
      ForeColor       =   &H000000FF&
      Height          =   465
      Left            =   780
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8172
      Width           =   2592
   End
   Begin VSFlex7Ctl.VSFlexGrid vs_Cr 
      Height          =   2175
      Left            =   8010
      TabIndex        =   44
      Top             =   7695
      Visible         =   0   'False
      Width           =   6855
      _cx             =   12091
      _cy             =   3836
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
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
      AutoResize      =   0   'False
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
   Begin VB.CheckBox Check1_sc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search School Wise"
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
      Left            =   5160
      TabIndex        =   43
      Top             =   8148
      Width           =   2556
   End
   Begin VB.ListBox cboschool 
      Appearance      =   0  'Flat
      Height          =   1752
      Left            =   7800
      Style           =   1  'Checkbox
      TabIndex        =   42
      Top             =   8010
      Visible         =   0   'False
      Width           =   4995
   End
   Begin VB.CheckBox Check1_donation 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Include Donation Bill"
      Height          =   300
      Left            =   3315
      TabIndex        =   40
      Top             =   60
      Width           =   1785
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "View All"
      Height          =   345
      Left            =   11400
      TabIndex        =   39
      Top             =   840
      Width           =   690
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   5085
      TabIndex        =   38
      Top             =   45
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Set &Adj.(%)"
      Height          =   375
      Left            =   13470
      TabIndex        =   37
      Top             =   795
      Width           =   960
   End
   Begin VB.TextBox txtAdPer 
      Height          =   315
      Left            =   12855
      TabIndex        =   35
      Top             =   840
      Width           =   555
   End
   Begin VB.ComboBox cboSer 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmSalesAdjustment.frx":000C
      Left            =   11925
      List            =   "frmSalesAdjustment.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   180
      Width           =   2325
   End
   Begin VB.TextBox txtDiffNet 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   13140
      TabIndex        =   30
      Top             =   6840
      Width           =   915
   End
   Begin VB.TextBox txtAdjAmt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   11520
      TabIndex        =   29
      Top             =   6900
      Width           =   915
   End
   Begin VB.TextBox txtscid 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5445
      TabIndex        =   17
      Top             =   885
      Width           =   735
   End
   Begin VB.ComboBox cmbAgentName 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmSalesAdjustment.frx":0010
      Left            =   9135
      List            =   "frmSalesAdjustment.frx":0012
      TabIndex        =   16
      Top             =   840
      Width           =   2265
   End
   Begin VB.ComboBox cboSession 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmSalesAdjustment.frx":0014
      Left            =   945
      List            =   "frmSalesAdjustment.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   885
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox txtSchoolName 
      Height          =   315
      Left            =   945
      TabIndex        =   14
      Top             =   885
      Width           =   4485
   End
   Begin VB.TextBox txtGTotal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8220
      TabIndex        =   13
      Top             =   6900
      Width           =   915
   End
   Begin VB.TextBox txtNetTotal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9900
      TabIndex        =   12
      Top             =   6900
      Width           =   915
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   165
      ScaleHeight     =   876
      ScaleWidth      =   6996
      TabIndex        =   3
      Top             =   8832
      Width           =   6990
      Begin VB.CommandButton save 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   720
         Left            =   945
         Picture         =   "frmSalesAdjustment.frx":0018
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   45
         Width           =   990
      End
      Begin VB.CommandButton Del 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   720
         Left            =   2925
         Picture         =   "frmSalesAdjustment.frx":0BFC
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   60
         Width           =   1005
      End
      Begin VB.CommandButton Abandon 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   720
         Left            =   0
         Picture         =   "frmSalesAdjustment.frx":17E0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   60
         Width           =   945
      End
      Begin VB.CommandButton REPORTCD 
         Caption         =   "&Print"
         Height          =   585
         Left            =   4860
         TabIndex        =   8
         Top             =   1020
         Width           =   1245
      End
      Begin VB.CommandButton close 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   720
         Left            =   6030
         Picture         =   "frmSalesAdjustment.frx":23C4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
         Width           =   930
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
         Height          =   720
         Left            =   1935
         Picture         =   "frmSalesAdjustment.frx":2FA8
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   945
      End
      Begin VB.CommandButton cmdPrint_7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   720
         Left            =   3975
         Picture         =   "frmSalesAdjustment.frx":33B5
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1005
      End
      Begin VB.CommandButton cmdPrint1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print Option"
         Height          =   720
         Left            =   4995
         Picture         =   "frmSalesAdjustment.frx":3F99
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   1005
      End
   End
   Begin VB.TextBox txtRemarks 
      Height          =   315
      Left            =   780
      MaxLength       =   250
      TabIndex        =   2
      Top             =   7275
      Width           =   13305
   End
   Begin VB.TextBox txtSponsorshipNo 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   45
      Width           =   735
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6360
      TabIndex        =   0
      Top             =   6900
      Width           =   975
   End
   Begin Crystal.CrystalReport cr 
      Left            =   12840
      Top             =   9180
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   5325
      Left            =   60
      TabIndex        =   18
      Top             =   1515
      Width           =   14385
      _cx             =   25374
      _cy             =   9393
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
      Cols            =   19
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSalesAdjustment.frx":4B7D
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
      Begin VSFlex7Ctl.VSFlexGrid VS_AppDet 
         Height          =   1830
         Left            =   0
         TabIndex        =   49
         Top             =   3645
         Visible         =   0   'False
         Width           =   13305
         _cx             =   23469
         _cy             =   3228
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
         FormatString    =   $"frmSalesAdjustment.frx":4C7D
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
   End
   Begin MSComCtl2.DTPicker txtDates 
      Height          =   330
      Left            =   1980
      TabIndex        =   19
      Top             =   45
      Width           =   1305
      _ExtentX        =   2307
      _ExtentY        =   572
      _Version        =   393216
      Format          =   149618689
      CurrentDate     =   39795
   End
   Begin VB.Label lblrow 
      BackStyle       =   0  'Transparent
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
      Left            =   90
      TabIndex        =   54
      Top             =   6885
      Width           =   2955
   End
   Begin VB.Label lblPRemarks 
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
      Height          =   675
      Left            =   5265
      TabIndex        =   50
      Top             =   45
      Width           =   5775
   End
   Begin VB.Label lblsc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "School Name :"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7800
      TabIndex        =   41
      Top             =   7800
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Adjust(%) :"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   12135
      TabIndex        =   36
      Top             =   885
      Width           =   735
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Series:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11370
      TabIndex        =   34
      Top             =   195
      Width           =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Diff. Net :"
      Height          =   285
      Left            =   12480
      TabIndex        =   32
      Top             =   6900
      Width           =   795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Adj. Amt :"
      Height          =   285
      Left            =   10920
      TabIndex        =   31
      Top             =   6900
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name :"
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
      Left            =   60
      TabIndex        =   28
      Top             =   885
      Width           =   1515
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 For Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   1200
      TabIndex        =   27
      Top             =   660
      Width           =   2715
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Representative :"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7980
      TabIndex        =   26
      Top             =   885
      Width           =   1170
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Total :"
      Height          =   285
      Left            =   7380
      TabIndex        =   25
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Total :"
      Height          =   285
      Left            =   9180
      TabIndex        =   24
      Top             =   6900
      Width           =   915
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   120
      Top             =   8784
      Width           =   7068
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks :"
      Height          =   285
      Left            =   60
      TabIndex        =   23
      Top             =   7335
      Width           =   1515
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment No :"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   22
      Top             =   60
      Width           =   1125
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F4 Key To Delete A Grid Item"
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
      Left            =   2580
      TabIndex        =   21
      Top             =   6915
      Width           =   2955
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty:"
      Height          =   285
      Left            =   5640
      TabIndex        =   20
      Top             =   6900
      Width           =   915
   End
End
Attribute VB_Name = "frmSalesAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dt_from As Date
Dim dt_to As Date
Dim dt_str As String
Dim bb_2 As Boolean
Dim bb1 As Boolean
Dim Edit As Boolean
Dim Add As Boolean
Dim str_ As String
Dim kk1 As Integer
Dim dt_strSaleNext As String
Dim dt_strSaleRNext As String
Dim dt_strR As String
Dim Sp_dt_from As Date
Dim Sp_dt_to As Date
Dim scid_st As String

Dim fromDt_sale As String
Dim toDt_sale As String

Dim fromDt_saleret As String
Dim toDt_saleret As String

   

Dim CON_next As ADODB.Connection
Function fatchDate(fyear_ As String, type_ As String, inv As Long, rows_) As String

   '''================================================================
   '''Checking
   '''================================================================
    
   Select Case fyear_
       
   Case session
    If (type_ = "I" Or type_ = "C/M") Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select top 1 INVOICEDATE from PartyWiseItemWiseQtySales where INVOICENO='" & inv & "'", con
        If rs1.EOF = False Then
            fatchDate = "PartyWiseItemWiseQtySales"
            current_next = "current"
            
        End If
    ElseIf type_ = "C" Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select top 1 INVOICEDATE from PartyWiseItemWiseQtySales_return where INVOICENO='" & inv & "'", con
        If rs1.EOF = False Then
           current_next = "current"
           fatchDate = "PartyWiseItemWiseQtySales_return"
        End If
    End If
    
    
    Case session_next
    
      If (type_ = "I" Or type_ = "C/M") Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select top 1 INVOICEDATE from PartyWiseItemWiseQtySales where INVOICENO='" & inv & "'", CON_next
        If rs1.EOF = False Then
           'fatchDate = rs1!INVOICEDATE
           current_next = "next"
           fatchDate = "PartyWiseItemWiseQtySales"
        End If
    ElseIf type_ = "C" Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select top 1 INVOICEDATE from PartyWiseItemWiseQtySales_return where INVOICENO='" & inv & "'", CON_next
        If rs1.EOF = False Then
           'fatchDate = rs1!INVOICEDATE
           fatchDate = "PartyWiseItemWiseQtySales_return"
           current_next = "next"
        End If
    End If
    
    
    End Select


End Function
Sub searchData()

  Dim rs_ As New ADODB.Recordset
  bb1 = False
  vs.Clear
  lblPRemarks.Caption = ""

  If RS.State = 1 Then RS.close
  RS.Open "select * from SalesAdjustment where DNo=" & txtSponsorshipNo & "", con, adOpenDynamic, adLockOptimistic
  If RS.EOF = False Then
     Check1_manual.value = 0
     
     bb1 = True
     vs.Enabled = False
     txtDates.value = RS!DDate
     txtScId.text = RS!scid
     txtSchoolName.text = RS!scname
     
     If rs_.State = 1 Then rs_.close
     rs_.Open "select partyremarks from SLEDGER where SUBLEDGER='" & RS!scid & "'", con
     If rs_.EOF = False Then
        lblPRemarks.Caption = rs_!PartyRemarks & ""
     End If
     
     cmbAgentName.text = RS!RepName & ""
     
     txtPrincipal = RS!Principal & ""
     txtMob = RS!mobile & ""
     txtRemarks = RS!remarks & ""
     txtGTotal = RS!GrossAmt
     txtNetTotal = RS!net
     If Len(RS!SponsorshipOn) > 0 Then
     cboSponse.text = RS!SponsorshipOn
     End If
     
     txtReturnAdj = RS!ReturnAdj
     txtAmtAfterAdj = RS!AmtAfter_ReturnAdj
     'txtPercentSp.Text = RS!Sponsorship_per
     txtFAmt = RS!finalAmt
     
     'txtAdvAmt.Text = RS!AdvAmt & ""
     txtRoundOf = RS!RoundOfAAmt & ""
     txtNetBal = RS!NetBalance & ""
     
     txtWhomToBeGivenMob = RS!MobileWhomtoGiven & ""
     
     txtAdjAmt = RS!Net_Adj & ""
     txtDiffNet = RS!Net_Diff & ""

     txtManually.text = RS!MannuallyRem & ""
     
     If (txtManually.text <> "") Then
        Check1_manullay.value = 1
     Else
        Check1_manullay.value = 0
     End If
     
     
     save.Enabled = False
     Del.Enabled = False
     cmdEdit_4.Enabled = True
     
   End If
   

txtGTotal = 0
txtNetTotal = 0
txtQty = 0

vs.Cols = 19

If RS.State = 1 Then RS.close
RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QTY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,ScId,Net_Adj,Net_Diff,DISCOUNT_Adj,id,scode,repname from SalesAdjustmentDet where  dno='" & txtSponsorshipNo & "' order by id", con
If RS.EOF = False Then
vs.rows = RS.RecordCount + 50
End If

For I = 1 To RS.RecordCount
    vs.TextMatrix(I, 0) = RS!fyear
    vs.TextMatrix(I, 1) = RS!Godown
    vs.TextMatrix(I, 2) = RS!invoiceNo
    vs.TextMatrix(I, 3) = RS!invoiceDate
    vs.TextMatrix(I, 4) = RS!Bookcode
    vs.TextMatrix(I, 5) = RS!Bookname
    vs.TextMatrix(I, 6) = RS!qty
    vs.TextMatrix(I, 7) = RS!rate
    vs.TextMatrix(I, 8) = RS!GrossAmt
    vs.TextMatrix(I, 9) = RS!discount
    vs.TextMatrix(I, 10) = Round(RS!net, 0)
    If (RS!Godown = "I" Or RS!Godown = "C/M") Then
        txtGTotal = Val(txtGTotal) + RS!GrossAmt
        txtNetTotal = Val(txtNetTotal) + Round(RS!net, 0)
    Else
        txtGTotal = Val(txtGTotal) - RS!GrossAmt
        txtNetTotal = Val(txtNetTotal) - Round(RS!net, 0)
    End If
    
    vs.TextMatrix(I, 11) = RS!DISCOUNT_Adj & ""
    vs.TextMatrix(I, 18) = RS!DISCOUNT_Adj & ""
    
    vs.TextMatrix(I, 12) = RS!Net_Adj & ""
    vs.TextMatrix(I, 13) = RS!Net_Diff & ""
    vs.TextMatrix(I, 14) = RS!id & ""
    
    vs.TextMatrix(I, 15) = RS!scode & ""
    
    vs.TextMatrix(I, 16) = RS!scname & ""
    vs.TextMatrix(I, 17) = RS!RepName & ""
    
    
    txtQty = Val(txtQty) + RS!qty
    

    RS.MoveNext
Next

txtGTotal = Round(txtGTotal, 0)

vs.Cols = 19

vs.FormatString = "Session|Inv.Type|Inv. No.|Inv. Date|B.Code|B.Name|>Quantity|>Rate|>GrossAmt|>Dis.%|>NetAmt|>Adj.Dis(%)|>Adj.Net|>Diff.NetAmt||Scode"

vs.ColWidth(1) = 700
vs.ColWidth(2) = 700
vs.ColWidth(3) = 900
vs.ColWidth(4) = 700
vs.ColWidth(5) = 2500
vs.ColWidth(6) = 800
vs.ColWidth(7) = 700
vs.ColWidth(8) = 950

vs.ColWidth(9) = 600
vs.ColWidth(10) = 1000

vs.ColWidth(11) = 1000
vs.ColWidth(12) = 1000
vs.ColWidth(13) = 1000
vs.ColWidth(14) = 0
vs.ColWidth(15) = 700
vs.ColWidth(16) = 0
vs.ColWidth(17) = 0



     
End Sub
Sub SearchDataNew(inv As String)

Dim conadj As ADODB.Connection
Dim rs_adj As ADODB.Recordset

Set conadj = New ADODB.Connection
Set rs_adj = New ADODB.Recordset
rs_adj.Open "select LastDatabase from data", CCON
If rs_adj.EOF = False Then
    Set con_don = New ADODB.Connection
    If LCase(server_) = "server" Then
       conadj.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & rs_adj!LastDatabase & "; UID=" & sql_user & "; PWD=" & sql_pass
       conadj.Open
    Else
       conadj.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=HP\SQL2008; DATABASE=" & rs_adj!LastDatabase & "; UID=; PWD=;"
       conadj.Open
    End If
End If


'=================================================================================
Dim rs_1 As New ADODB.Recordset
Dim lastData_ As Boolean
Dim billDate_ As Date

lastData_ = False

str11 = " SELECT distinct SalesAdjustment.DNo,SalesAdjustment.DDate,SalesAdjustmentDet.Godown FROM SalesAdjustment INNER JOIN" & _
      " SalesAdjustmentDet ON SalesAdjustment.DNo = SalesAdjustmentDet.DNo where SalesAdjustmentDet.fyear='" & session & "' and SalesAdjustmentDet.INVOICENO=" & inv_ledger & ""
ss1_ = ""


str11 = " SELECT  ENTRYNO AS DNo,DATES AS DDate,YRS FROM tmpSaladjust where  substring(PARTY,1,5)='" & Mid(pname_, 1, 5) & "' and BILLNO=" & inv_ledger & " AND YRS IS NOT NULL"



If rs_1.State = 1 Then rs_1.close
rs_1.Open str11, con
If rs_1.EOF = False Then

   n1 = Right(session, 2)
   N2 = Right(rs_1!YRS, 2)
   If N2 < n1 Then
      lastData_ = True
   Else
      lastData_ = False
   End If
   popupvalue5 = ""
End If

billDate_ = PopUpValue6

''If (billDate_ >= financialyear_Fdate And billDate_ <= financialyear_Tdate) Then
''   lastData_ = False
''   PopUpValue6 = ""
''Else
''   lastData_ = True
''End If




If lastData_ = True Then
   Picture3.Enabled = False
   Me.Caption = rs_adj!LastDatabase
   cmdEdit_4.Enabled = False
   Abandon.Enabled = False
   cmdPrint_7.Enabled = False
   cmdPrint1.Enabled = False
Else
   Picture3.Enabled = True
   Me.Caption = session
   cmdEdit_4.Enabled = True
   Abandon.Enabled = True
   cmdPrint_7.Enabled = True
   cmdPrint1.Enabled = True

End If


'==================================================================================
  bb1 = False
  vs.Clear

  If RS.State = 1 Then RS.close
  If lastData_ = True Then
    RS.Open "select * from SalesAdjustment where DNo=" & txtSponsorshipNo & "", conadj, adOpenDynamic, adLockOptimistic
  Else
    RS.Open "select * from SalesAdjustment where DNo=" & txtSponsorshipNo & "", con, adOpenDynamic, adLockOptimistic
  End If
  
  If RS.EOF = False Then
     bb1 = True
     vs.Enabled = False
     txtDates.value = RS!DDate
     txtScId.text = RS!scid
     txtSchoolName.text = RS!scname
     cmbAgentName.text = RS!RepName & ""
     
     txtPrincipal = RS!Principal & ""
     txtMob = RS!mobile & ""
     txtRemarks = RS!remarks & ""
     txtGTotal = RS!GrossAmt
     txtNetTotal = RS!net
     If Len(RS!SponsorshipOn) > 0 Then
     cboSponse.text = RS!SponsorshipOn
     End If
     
     txtReturnAdj = RS!ReturnAdj
     txtAmtAfterAdj = RS!AmtAfter_ReturnAdj
     'txtPercentSp.Text = RS!Sponsorship_per
     txtFAmt = RS!finalAmt
     
     'txtAdvAmt.Text = RS!AdvAmt & ""
     txtRoundOf = RS!RoundOfAAmt & ""
     txtNetBal = RS!NetBalance & ""
     
     txtWhomToBeGivenMob = RS!MobileWhomtoGiven & ""
     
     txtAdjAmt = RS!Net_Adj & ""
     txtDiffNet = RS!Net_Diff & ""

     
     
     save.Enabled = False
     Del.Enabled = False
     cmdEdit_4.Enabled = True
     
   End If
   

txtGTotal = 0
txtNetTotal = 0
txtQty = 0

''vs.Cols = 14
vs.Cols = 19

If RS.State = 1 Then RS.close
If lastData_ = True Then
   RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QTY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,ScId,Net_Adj,Net_Diff,DISCOUNT_Adj,scode,scode,repname,id from SalesAdjustmentDet where  dno='" & txtSponsorshipNo & "' order by id", conadj
Else
   RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QTY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,ScId,Net_Adj,Net_Diff,DISCOUNT_Adj,scode,scode,repname,id from SalesAdjustmentDet where  dno='" & txtSponsorshipNo & "' order by id", con
End If

If RS.EOF = False Then
vs.rows = RS.RecordCount + 50
End If

For I = 1 To RS.RecordCount
    vs.TextMatrix(I, 0) = RS!fyear
    vs.TextMatrix(I, 1) = RS!Godown
    vs.TextMatrix(I, 2) = RS!invoiceNo
    vs.TextMatrix(I, 3) = RS!invoiceDate
    vs.TextMatrix(I, 4) = RS!Bookcode
    vs.TextMatrix(I, 5) = RS!Bookname
    vs.TextMatrix(I, 6) = RS!qty
    vs.TextMatrix(I, 7) = RS!rate
    vs.TextMatrix(I, 8) = RS!GrossAmt
    vs.TextMatrix(I, 9) = RS!discount
    vs.TextMatrix(I, 10) = Round(RS!net, 0)
    If (RS!Godown = "I" Or RS!Godown = "C/M") Then
        txtGTotal = Val(txtGTotal) + RS!GrossAmt
        txtNetTotal = Val(txtNetTotal) + Round(RS!net, 0)
    Else
        txtGTotal = Val(txtGTotal) - RS!GrossAmt
        txtNetTotal = Val(txtNetTotal) - Round(RS!net, 0)
    End If
    
    vs.TextMatrix(I, 11) = RS!DISCOUNT_Adj & ""
    vs.TextMatrix(I, 12) = RS!Net_Adj & ""
    vs.TextMatrix(I, 13) = RS!Net_Diff & ""
    
    vs.TextMatrix(I, 14) = RS!id & ""
    
    vs.TextMatrix(I, 15) = RS!scode & ""
    
    vs.TextMatrix(I, 16) = RS!scname & ""
    vs.TextMatrix(I, 17) = RS!RepName & ""
    
    txtQty = Val(txtQty) + RS!qty
    

    RS.MoveNext
Next

txtGTotal = Round(txtGTotal, 0)

'vs.Cols = 13

vs.FormatString = "Session|Inv.Type|Inv. No.|Inv. Date|B.Code|B.Name|>Quantity|>Rate|>GrossAmt|>Dis.%|>NetAmt|>Adj.Dis(%)|>Adj.Net|>Diff.NetAmt|SCode"

vs.ColWidth(1) = 700
vs.ColWidth(2) = 700
vs.ColWidth(3) = 900
vs.ColWidth(4) = 700
vs.ColWidth(5) = 2600
vs.ColWidth(6) = 800
vs.ColWidth(7) = 700
vs.ColWidth(8) = 950

vs.ColWidth(9) = 600
vs.ColWidth(10) = 1000

vs.ColWidth(11) = 1000
vs.ColWidth(12) = 1000
vs.ColWidth(13) = 1000
vs.ColWidth(14) = 0
vs.ColWidth(15) = 700
vs.ColWidth(16) = 0
vs.ColWidth(17) = 0







     
End Sub

Private Sub ABANDON_Click()

refresh_
max_sp
Edit = False
con.Execute "delete from tmpSalesAdj"
cmdEdit_4.Enabled = False
Del.Enabled = False

'List1_sc.Clear
'List1_Sp.Visible = False


End Sub

Private Sub cboPayment_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtWhomTobeGiven.SetFocus
End Sub
Private Sub cboschool_Click()
   
   str_ = ""
   For I = 0 To cboschool.ListCount - 1
   If cboschool.Selected(I) = True Then
      If str_ = "" Then
         str_ = "scname='" & cboschool.List(I) & "'"
      Else
         str_ = str_ & " or scname='" & cboschool.List(I) & "'"
      End If
   End If
   Next

   fillGrid_
   
End Sub
Private Sub cboschool_LostFocus()
   
   ss_ = ""
   For I = 0 To cboschool.ListCount - 1
   If cboschool.Selected(I) = True Then
      If ss_ = "" Then
         ss_ = "" & cboschool.List(I) & ""
      Else
         ss_ = ss_ & "," & cboschool.List(I) & ""
      End If
   End If
   Next
   
   txtRemarks.text = ss_

End Sub

Private Sub cboSer_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   txtAdPer.SetFocus
End If

End Sub

Private Sub cboSession_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cboPayment.SetFocus
End Sub


Private Sub cboSponse_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtReturnAdj.SetFocus
End Sub
Private Sub Check1_manullay_Click()


''If (Check1_manullay.value = 1) Then
''    For I = 1 To vs.rows - 1
''
''        If (vs.TextMatrix(I, 11) <> "") Then
''            vs.TextMatrix(I, 18) = vs.TextMatrix(I, 11)
''        End If
''
''    Next
''Else
''
''    For I = 1 To vs.rows - 1
''
''        If (vs.TextMatrix(I, 11) <> "") Then
''            vs.TextMatrix(I, 18) = 0
''        End If
''
''    Next
''End If



End Sub

Private Sub Check1_PendingCr_Click()
    
    If Check1_PendingCr.value = 0 Then
        vs_Cr.Visible = False
        cmdRepQty.Visible = False
        Exit Sub
     Else
        cmdRepQty.Visible = True
    End If
    
    Dim rsf As New ADODB.Recordset
    
    
    vs_Cr.Visible = True
    vs_Cr.Clear
    
    Screen.MousePointer = vbHourglass
    con.Execute "exec tmpdata_saleadj"
    con.Execute "update credita set TXT1=''"
        
    ' con.Execute "update credita set TXT1='y'  where INVOICENO in " & _
    '"(select distinct INVOICENO from SalesAdjustmentDet where godown='C' and fyear='" & session & "')"
     con.Execute "update credita set TXT1='y'  where INVOICENO in " & _
    "(select distinct INVOICENO from tmpSAdjDet where godown='C' and fyear='" & session & "')"
    
      
    
    
    DoEvents
    DoEvents
    DoEvents
    
      
    Dim K As Integer
    Dim adj As String
    K = 1
    
    vs_Cr.Cols = 4
    vs_Cr.rows = 2
    
     
     If RS.State = 1 Then RS.close
     'RS.Open "SELECT distinct SUBLEDGER FROM SalesAdjustmentDet where godown='C'", con
     RS.Open "SELECT distinct party FROM tmpSAdjDet where godown='C' and fyear='" & session & "'", con
     While RS.EOF = False
        
        adj = ""
       If rs1.State = 1 Then rs1.close
       rs1.Open "SELECT distinct dno FROM tmpSAdjDet where godown='C' and  party='" & RS(0) & "' and fyear='" & session & "'", con
       While rs1.EOF = False
           If adj = "" Then
              If Val(rs1!dno) > 0 Then
                 adj = rs1!dno
              End If
           Else
              If Val(rs1!dno) > 0 Then
                  adj = adj & "," & rs1!dno
              End If
           End If
       rs1.MoveNext
       Wend
        
        
        
        If rsf.State = 1 Then rsf.close
        rsf.Open "select INVOICENO,InvoiceDate,SUBLEDGER,TXT1 from CREDITA where  SUBLEDGER='" & RS(0) & "' order by INVOICENO", con, adOpenDynamic, adLockReadOnly
       While rsf.EOF = False
             
             If rsf!txt1 = "" Then
                  
                  '=====================================
                  
                  vs_Cr.TextMatrix(K, 0) = rsf!invoiceNo
                  vs_Cr.TextMatrix(K, 1) = rsf!invoiceDate
                  vs_Cr.TextMatrix(K, 2) = rsf!subledger
                  vs_Cr.TextMatrix(K, 3) = adj
                  DoEvents
                  DoEvents
    
                  vs_Cr.rows = vs_Cr.rows + 1
                  K = K + 1
                 adj = rsf!txt1
             End If
           rsf.MoveNext
       Wend
       
     
     RS.MoveNext
    Wend
    
    vs_Cr.FormatString = "Inv.No|InvoiceDate|Party Name...|Party Adj.No"
    vs_Cr.ColWidth(0) = 700
    vs_Cr.ColWidth(1) = 1000
    vs_Cr.ColWidth(2) = 3200
    vs_Cr.ColWidth(3) = 1600
    
        
    Screen.MousePointer = vbDefault
    
    
End Sub

Private Sub Check1_sc_Click()
If Check1_sc.value = 1 Then
 lblsc.Visible = True
 cboschool.Visible = True
 
 
Else
 lblsc.Visible = False
 cboschool.Visible = False

End If



End Sub

Private Sub Check2_AppDet_Click()

Dim rs_f As ADODB.Recordset
Set rs_f = New ADODB.Recordset

Screen.MousePointer = vbHourglass

rs_f.Open "SELECT appNo,scid,scname,Promo,Adj,Discount,AppPer,(Promo+Adj+Discount+AppPer) as TDis,Net_Gross,SerName " & _
"from PartyRemarksQryNew where substring(SUBLEDGER,1,5)='" & Mid(txtScId, 1, 5) & "'  order by appNo,scid", con

Set VS_AppDet.DataSource = rs_f

VS_AppDet.FormatString = "AppNo|ScId|ScName|Promo|Adj|Discount|AppPer|TDis|Net_Gross|SerName"

VS_AppDet.ColWidth(0) = 800
VS_AppDet.ColWidth(1) = 900
VS_AppDet.ColWidth(2) = 4000
VS_AppDet.ColWidth(3) = 800
VS_AppDet.ColWidth(4) = 800
VS_AppDet.ColWidth(5) = 800
VS_AppDet.ColWidth(6) = 800
VS_AppDet.ColWidth(7) = 900
VS_AppDet.ColWidth(8) = 1000
VS_AppDet.ColWidth(9) = 1000

VS_AppDet.Editable = flexEDNone


Screen.MousePointer = vbDefault
'======================================

If Check2_AppDet.value = 1 Then
   VS_AppDet.Visible = True
Else
   VS_AppDet.Visible = False
End If



End Sub

Private Sub close_Click()
Unload Me
End Sub

Private Sub cmbAgentName_Click()
fillGrid_
End Sub

Private Sub cmbAgentName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    'cboPayment.SetFocus
    cboser.SetFocus
End If
End Sub
Sub max_sp()
If RS.State = 1 Then RS.close
RS.Open "select max(DNo) from SalesAdjustment", con
If Not IsNull(RS(0)) Then
   txtSponsorshipNo = RS(0) + 1
Else
   txtSponsorshipNo = 1
End If

End Sub

Private Sub cmdAll_Click()
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

End Sub

Private Sub cmdEdit_4_Click()
con.Execute "delete from tmpSalesAdj where Sno=" & txtSponsorshipNo & ""
con.Execute "insert into tmpSalesAdj(INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,sNo,repname,SCode) select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,DNo,repname,SCode from SalesAdjustmentDet where dno=" & txtSponsorshipNo & ""
Del.Enabled = True
save.Enabled = True
cmdEdit_4.Enabled = False
Edit = True
vs.Enabled = True
End Sub

Private Sub cmdEditDet_Click()
 
Dim s_ As String
 
s_ = ""
 
If txtSponsorshipNo.text <> "" Then
 If rs1.State = 1 Then rs1.close
 rs1.Open "SELECT UserName,No,vtype,desc_,dates,id FROM logtbl where vtype='adjustment' and no=" & txtSponsorshipNo.text & " order by id", con
 While rs1.EOF = False
    
   If s_ = "" Then
      s_ = "User Name : " & rs1!UserName & "; Edit Details & Diff.Net : " & rs1!desc_ & "; Date : " & rs1!dates & " " & vbCrLf
   Else
      s_ = s_ & "User Name : " & rs1!UserName & "; Edit Details & Diff.Net : " & rs1!desc_ & "; Date : " & rs1!dates & " " & vbCrLf
   End If
  
   rs1.MoveNext
 Wend
 
If s_ <> "" Then
   MsgBox "" & s_, vbInformation, "Adjustment Edit Details"
End If
 
End If

End Sub

Private Sub cmdOk_Click()

Screen.MousePointer = vbHourglass

    
    txtAdjAmt = 0
    txtDiffNet = 0
    
    
    
  'If cboser <> "" Then
    
    For I = 1 To vs.rows - 1
    If vs.TextMatrix(I, 0) <> "" Then
    If RS.State = 1 Then RS.close
    RS.Open "select bookcode from books where sername='" & cboser & "' order by bookcode", con
    While RS.EOF = False
    If vs.TextMatrix(I, 4) = RS(0) Then
       
       vs.TextMatrix(I, 11) = Val(txtAdPer)
    
       If (Val(vs.TextMatrix(I, 18)) > 0) Then
            If (Val(vs.TextMatrix(I, 11)) > Val(vs.TextMatrix(I, 18))) Then
               vs.TextMatrix(I, 11) = vs.TextMatrix(I, 18)
               MsgBox "You can set mannually Less Disscount Only...", vbCritical
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
        End If
    
    
        
        vs.TextMatrix(I, 12) = Round(Val(vs.TextMatrix(I, 8)) - (Val(vs.TextMatrix(I, 8)) * Val(txtAdPer) / 100), 0)
        vs.TextMatrix(I, 13) = (Val(vs.TextMatrix(I, 10)) - Val(vs.TextMatrix(I, 12)))
        
    End If
    RS.MoveNext
    Wend
    End If
    Next


    
    
    
    
    txtAdjAmt = 0
    txtDiffNet = 0
    
    For k1 = 1 To vs.rows - 1
    If vs.TextMatrix(k1, 0) <> "" Then
    If (vs.TextMatrix(k1, 1) = "I" Or vs.TextMatrix(k1, 1) = "C/M") Then
       txtAdjAmt = Val(txtAdjAmt) + Val(vs.TextMatrix(k1, 12))
       txtDiffNet = Val(txtDiffNet) + Val(vs.TextMatrix(k1, 13))
    Else
       txtAdjAmt = Val(txtAdjAmt) - Val(vs.TextMatrix(k1, 12))
       txtDiffNet = Val(txtDiffNet) - Val(vs.TextMatrix(k1, 13))
    End If
    End If
    Next








Screen.MousePointer = vbDefault
           
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


If (txtScId.text <> "") Then
   If scid_st <> "" Then
      str1 = "SELECT SUBLEDGER,PartyRemarks,appNo,Promo,Net_Gross,adj,discount,appPer,tod,cd,remarks FROM PartyRemarksQrynew where " & scid_st & " group by SUBLEDGER,PartyRemarks,appNo,Promo,Net_Gross,adj,discount,appPer,tod,cd,remarks"
   Else
      str1 = "SELECT SUBLEDGER,PartyRemarks,appNo,Promo,Net_Gross,adj,discount,appPer,tod,cd,remarks FROM PartyRemarksQrynew where SUBLEDGER='" & txtScId.text & "' group by SUBLEDGER,PartyRemarks,appNo,Promo,Net_Gross,adj,discount,appPer,tod,cd,remarks"
   End If
Else
   If txtPartyName.text <> "" Then
      con.Execute "insert into AppPrintTmp1(Party,Remarks) select subledger,partyremarks from sledger where subledger='" & txtPartyName.text & "'"
   End If
   Exit Sub
End If


ss_ = ""



    If f.State = 1 Then f.close
    '''If (DateValue(txtDates.value) >= DateValue(financialyear_Fdate) And DateValue(txtDates.value) <= DateValue(financialyear_Tdate)) Then
       f.Open str1, con
    '''Else
    '''f.Open str1, con_LAST
    '''End If



If f.RecordCount = 0 Then
   If txtSchoolName.text <> "" Then
      con.Execute "insert into AppPrintTmp1(Party,Remarks) select subledger,partyremarks from sledger where code='" & txtScId.text & "'"
   End If
End If


While f.EOF = False
      
       ss_ = ""
       remarks1 = ""
       If rs1.State = 1 Then rs1.close
       If (DateValue(txtDates.value) >= DateValue(financialyear_Fdate) And DateValue(txtDates.value) <= DateValue(financialyear_Tdate)) Then
          rs1.Open "select sername,Promo,PartyRemarks,Adj,discount,appPer,remarks from PartyRemarksQryNew where (appNo=" & f(2) & " and Promo=" & f!Promo & " and adj=" & f.Fields("adj").value & " and discount=" & f.Fields("discount").value & " and appPer=" & f.Fields("appPer").value & ") group by sername,Promo,PartyRemarks,Adj,discount,appPer,remarks", con
       Else
         rs1.Open "select sername,Promo,PartyRemarks,Adj,discount,appPer,remarks from PartyRemarksQryNew where (appNo=" & f(2) & " and Promo=" & f!Promo & " and adj=" & f.Fields("adj").value & " and discount=" & f.Fields("discount").value & " and appPer=" & f.Fields("appPer").value & ") group by sername,Promo,PartyRemarks,Adj,discount,appPer,remarks", con_LAST
       End If
       
       If rs1.EOF = False Then
          'stt1 = RemoveEnterChar(rs1!PartyRemarks)
          
          remarks1 = Trim(rs1!PartyRemarks) & ""
          
       End If
       
       While rs1.EOF = False
          If ss_ = "" Then
           ss_ = Trim(rs1(0))
          Else
           ss_ = ss_ & "," & Trim(rs1(0))
          End If
          rs1.MoveNext
       Wend
     
      If (f(2) > 0) Then
         tdis = (f.Fields("appPer").value + f.Fields("adj").value + f.Fields("discount").value + f.Fields("Promo").value)
         con.Execute "insert into AppPrintTmp1(party,Remarks,appno,PromPer,adjper,sername,remarks1,gross_,upto_5,TDis,tod,cd,appremarks) values('" & f(0) & "','" & remarks1 & "','" & f(2) & "','" & f(3) & "','" & f.Fields("adj").value & "','" & ss_ & "','" & f.Fields("discount").value & "','" & f(4) & "','" & f.Fields("appper").value & "','" & tdis & "','" & f.Fields("tod").value & "','" & f.Fields("cd").value & "','" & f.Fields("remarks") & "')"
      End If
     
     
     
     ss_ = ""

       
    f.MoveNext
Wend


'If RS.State = 1 Then RS.close
'RS.Open "SELECT SUBLEDGER,PartyRemarks,appNo FROM invoiceaQry where (scid='" & txtScId.Text & "' and SUBLEDGER='" & txtPartyName.Text & "') group by SUBLEDGER,PartyRemarks,appNo", con, adOpenDynamic, adLockOptimistic
'While RS.EOF = False
'   If rs1.State = 1 Then rs1.close
'   rs1.Open "SELECT party,Remarks,appNo FROM AppPrintTmp1 where party='" & RS!SUBLEDGER & "'", con, adOpenDynamic, adLockOptimistic
'   If rs1.RecordCount > 0 Then
'      con.Execute "delete from AppPrintTmp1 where party='" & rs1!party & "' and (AppNo ='' or AppNo =0) and (remarks='NA' or remarks='')"
'   End If
'   RS.MoveNext
'Wend




'''Dim f As New ADODB.Recordset
'''Set f = New ADODB.Recordset
'''st10 = ""
'''
'''
'''con.Execute "delete from AppPrintTmp"
'''If (txtScId.Text <> "" And txtSchoolName.Text <> "") Then
'''   str1 = "SELECT SUBLEDGER,PartyRemarks FROM invoiceaQry where SUBSTRING(SUBLEDGER,1,5)='" & Trim(Mid(txtScId.Text, 1, 5)) & "' group by SUBLEDGER,PartyRemarks"
'''End If
'''
''''-----------------------------------------
'''
'''If f.State = 1 Then f.close
'''f.Open str1, con
'''While f.EOF = False
'''    If f(1) <> "NA" Then
'''    If Len(f(1)) > 0 Then
'''       con.Execute "insert into AppPrintTmp(party,Remarks) values('" & f(0) & "','" & f(1) & "')"
'''    End If
'''    End If
'''    f.MoveNext
'''Wend
'''
'''
'''
'''
'''Dim ss1_ As String
'''
'''ss1_ = ""
'''
'''If scid_st <> "" Then
'''   str1 = "SELECT APPNO FROM invoiceaQry where (SUBSTRING(SUBLEDGER,1,5)='" & Trim(Mid(txtScId.Text, 1, 5)) & "' and len(APPNO)>0) and " & scid_st & " group by APPNO"
'''Else
'''   Exit Sub
'''End If
'''
'''If f.State = 1 Then f.close
'''f.Open str1, con
'''While f.EOF = False
'''
'''ss_ = ""
'''If rs1.State = 1 Then rs1.close
'''rs1.Open "select  ADJ from AppForm where appNo=" & f(0) & " group by ADJ", con
'''While rs1.EOF = False
'''   If ss_ = "" Then
'''      ss_ = rs1(0)
'''   Else
'''      ss_ = ss_ & ";" & rs1(0)
'''   End If
'''   rs1.MoveNext
'''Wend
'''
'''If ss1_ = "" Then
'''  ss1_ = " AppNo : " & f(0) & " - " & ss_
'''Else
'''  ss1_ = ss1_ & " ; AppNo : " & f(0) & " - " & ss_
'''End If
'''
'''f.MoveNext
'''Wend
'''
'''If ss1_ <> "" Then
'''con.Execute "update AppPrintTmp set PromPer='" & ss1_ & "'"
'''End If







End Sub

Private Sub cmdPrint_7_Click()

Screen.MousePointer = vbHourglass

If Check1_partyTerms.value = 1 Then
   con.Execute "update SalesAdjustment set partyTerms='n' where dno=" & txtSponsorshipNo.text & ""
Else
   con.Execute "update SalesAdjustment set partyTerms='y' where dno=" & txtSponsorshipNo.text & ""
End If



DSNNew

'''-----------------------------------------------------------
scid_st = ""




For I = 1 To vs.rows - 1


If vs.TextMatrix(I, 1) = "C" Then
   If session = vs.TextMatrix(I, 0) Then
      con.Execute "update a set a.t2=b.AppNO  FROM CREDITA as a inner join AppForm as b on (a.SUBLEDGER = b.code + ' ' + b.PName and a.ScID = b.id) Where a.invoiceNo = " & vs.TextMatrix(I, 2) & ""
   Else
   
   If Not (CON_next Is Nothing) Then
   
      If RS.State = 1 Then RS.close
      RS.Open "select subledger,scid from CREDITA where invoiceno=" & vs.TextMatrix(I, 2) & "", CON_next
      If RS.EOF = False Then
         
         If rs1.State = 1 Then rs1.close
         rs1.Open "select code,PName,id,appno from AppForm where (id='" & RS!scid & "')", con
         If rs1.EOF = False Then
            If scid_st = "" Then
               scid_st = "scid='" & rs1!id & "'"
            Else
               scid_st = scid_st & " or scid='" & rs1!id & "'"
            End If
         End If
      End If
      
    End If
      
      
   End If
End If
Next



scid_st = ""
con.Execute "delete from temp"

For I = 1 To vs.rows - 1
    
    If (vs.TextMatrix(I, 15) <> "") Then
    If scid_st = "" Then
            scid_st = "scid='" & vs.TextMatrix(I, 15) & "'"
        Else
            scid_st = scid_st & " or scid='" & vs.TextMatrix(I, 15) & "'"
         End If
         
    If (Len(vs.TextMatrix(I, 15)) > 2) Then
    con.Execute "insert into temp(text) values('" & vs.TextMatrix(I, 15) & "')"
    End If
         
    End If

Next



If Len(scid_st) > 4 Then
   scid_st = "(" & scid_st & ")"
End If

''
'''-----------------------------------------------------------


con.Execute "delete from tmpPartyRemarksQryNew"
con.Execute "insert into tmpPartyRemarksQryNew(PartyRemarks,SerName,appNo,discount,AppPer,Adj,Promo,Net_Gross,cd,tod) " & _
"select distinct PartyRemarks,SerName,appNo,discount,AppPer,Adj,Promo,Net_Gross,cd,tod from PartyRemarksQryNew  where SUBLEDGER ='" & txtScId.text & "' and scid in(SELECT [text] FROM temp group by text)"

'''-----------------------------------------------------------


''printPRemarks


If MsgBox("Want to Print ?", vbQuestion + vbYesNo) = vbYes Then


    CR.Reset
    CR.ReportFileName = rptPath & "/Adjustment.rpt"
    CR.ReplaceSelectionFormula "{SalesAdjustment.dno}=" & txtSponsorshipNo & ""
    CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    
    CR.Formulas(0) = "netamt=" & Val(txtNetTotal) & ""
    CR.Formulas(1) = "grossamt=" & Val(txtGTotal) & ""
    CR.Formulas(2) = "adj_net=" & Val(txtAdjAmt) & ""
    CR.Formulas(3) = "adj_diff=" & Val(txtDiffNet) & ""
    CR.Formulas(4) = "fyear='" & session & "'"
    

    CR.WindowShowPrintSetupBtn = True
    CR.WindowShowPrintBtn = True
    
    CR.WindowShowExportBtn = True
    CR.WindowState = crptMaximized
    CR.Action = 1

End If

Screen.MousePointer = vbDefault

End Sub

Private Sub cmdPrint1_Click()
frmDonnationPrint.Show
End Sub





Private Sub Combo1_Change()

End Sub

Private Sub cmdRepQty_Click()

Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim I As Integer
Dim xl As Excel.Application
Dim str_ As String




If xl Is Nothing Then
   Set xl = CreateObject("Excel.Application")
End If

Set xlBook = xl.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add
xl.Visible = True


Dim c, r As Long
Dim row_ As Integer



 col_ = 1
 row_ = 1

 xl.Columns("A:H").ColumnWidth = 12
 J = 2
 xlSheet.Cells(1, 1).value = " "
 
 For I = 0 To vs_Cr.rows - 1
     For J = 0 To vs_Cr.Cols - 1
         If (col_ = 2 Or col_3) Then
            xlSheet.Cells(row_, col_).value = Format(vs_Cr.TextMatrix(I, J), "dd/MM/yyyy")
         Else
            xlSheet.Cells(row_, col_).value = vs_Cr.TextMatrix(I, J)
         End If
         col_ = col_ + 1
     Next
     row_ = row_ + 1
     col_ = 1
 Next
    
 

End Sub

Private Sub cmdVew_Click()

Dim str_10 As String
Dim dateRet As String

Dim dateSale As String

''txtscid.Text = "M2012 MAHESH TRADERS, BHOPAL"

dateSale = "(INVOICEDATE >= convert(smalldatetime,'" & fromDt_sale & "',103) and INVOICEDATE <= convert(smalldatetime,'" & toDt_sale & "',103))"

dateRet = "(INVOICEDATE >= convert(smalldatetime,'" & fromDt_saleret & "',103) and INVOICEDATE <= convert(smalldatetime,'" & toDt_saleret & "',103))"



vs.Clear
vs.FormatString = "Session|In.Type|Inv. No.|Inv. Date|B.Code|B.Name|>Quantity|>Rate|>GrossAmt|>Dis.%|>NetAmt|>Adj.Dis(%)|>Adj.Net|>Diff.NetAmt||Scode"


vs.ColWidth(1) = 700
vs.ColWidth(2) = 600
vs.ColWidth(3) = 900
vs.ColWidth(4) = 750
vs.ColWidth(5) = 2450
vs.ColWidth(6) = 800
vs.ColWidth(7) = 700
vs.ColWidth(8) = 950

vs.ColWidth(9) = 600
vs.ColWidth(10) = 1000

vs.ColWidth(11) = 1000
vs.ColWidth(12) = 1000
vs.ColWidth(13) = 1000
vs.ColWidth(14) = 0
vs.ColWidth(15) = 700
vs.ColWidth(16) = 0
vs.ColWidth(17) = 0

vs.rows = 20
Dim k1 As Integer
Dim Qty_, dqty As Integer



k1 = 1
Qty_ = 0
dqty = 0

str10 = ""
 For I = 0 To cboschool.ListCount - 1
 If cboschool.Selected(I) = True Then
    If str10 = "" Then
       str10 = "scname='" & IIf(cboschool.List(I) = "N", "", cboschool.List(I)) & "'"
    Else
       str10 = str10 & " or scname='" & IIf(cboschool.List(I) = "N", Null, cboschool.List(I)) & "'"
    End If
 End If
 Next

If (cmbAgentName.text <> "") Then
   If (str10 = "") Then
       str10 = "agentname='" & cmbAgentName.text & "'"
   Else
       str10 = str10 & " and agentname='" & cmbAgentName.text & "'"
   End If
End If


'=============================================================
'=============================================================

If (str10 <> "") Then
    str_10 = "SELECT fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,RATE,GrossAmt,DISCOUNT,Net,sum(Donqty) as Donqty,sum(Adjqty) as Adjqty,sum(Saleqty) as Saleqty ,sum(SaleRQty) as SaleRQty,subledger,Scid,scname,agentname as repname" & _
    " FROM adjustmentqry where Godown='I' and (" & str10 & ") and subledger='" & txtScId.text & "' and " & dateSale & " group by fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,RATE,GrossAmt,DISCOUNT,Net,subledger,Scid,scName,agentname"
Else
    str_10 = "SELECT fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,RATE,GrossAmt,DISCOUNT,Net,sum(Donqty) as Donqty,sum(Adjqty) as Adjqty,sum(Saleqty) as Saleqty ,sum(SaleRQty) as SaleRQty,subledger,Scid,scname,agentname as repname" & _
    " FROM adjustmentqry where Godown='I' and subledger='" & txtScId.text & "' and " & dateSale & " group by fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,RATE,GrossAmt,DISCOUNT,Net,subledger,Scid,scname,agentname"

End If

Set rs1 = New ADODB.Recordset

If RS.State = 1 Then RS.close
RS.Open str_10, con
While RS.EOF = False

Qty_ = 0
dqty = 0

If rs1.State = 1 Then rs1.close
rs1.Open "SELECT  sum(Donqty) as Donqty,sum(Adjqty) as Adjqty,sum(Saleqty) as Saleqty ,sum(SaleRQty) as SaleRQty" & _
" FROM adjustmentqry where (rate=" & RS!rate & " and bookcode='" & RS!Bookcode & "' and Godown='I' and subledger='" & txtScId.text & "' and INVOICENO=" & RS!invoiceNo & " and fyear='" & RS!fyear & "')", con

If Check1_donation.value = 0 Then
   
   If Not IsNull(rs1!Donqty) Then
   dqty = rs1!Donqty
   End If
   
   If Not IsNull(rs1!Adjqty) Then
      dqty = dqty + rs1!Adjqty
   End If
   
   If (RS!Godown = "I" Or RS!Godown = "C/M") Then
      If Not IsNull(rs1!saleQty) Then
         Qty_ = rs1!saleQty - dqty
      End If
   End If
   
Else
   
   If Not IsNull(rs1!Adjqty) Then
      dqty = dqty + rs1!Adjqty
   End If
   
   If (RS!Godown = "I" Or RS!Godown = "C/M") Then
      If Not IsNull(rs1!saleQty) Then
         Qty_ = rs1!saleQty - dqty
      End If
   End If
   
End If


If (Qty_ > 0) Then

    vs.TextMatrix(k1, 0) = RS!fyear
    vs.TextMatrix(k1, 1) = RS!Godown
    vs.TextMatrix(k1, 2) = RS!invoiceNo
    vs.TextMatrix(k1, 3) = RS!invoiceDate
    vs.TextMatrix(k1, 4) = RS!Bookcode
    vs.TextMatrix(k1, 5) = RS!Bookname
    vs.TextMatrix(k1, 6) = Qty_
    vs.TextMatrix(k1, 7) = RS!rate
    vs.TextMatrix(k1, 8) = RS!GrossAmt
    vs.TextMatrix(k1, 9) = RS!discount
    vs.TextMatrix(k1, 10) = RS!net
    If txtAdPer.text <> "" Then
       vs.TextMatrix(k1, 11) = Val(txtAdPer.text)
    Else
       vs.TextMatrix(k1, 11) = 0
    End If
    
    vs.TextMatrix(k1, 12) = Round(Val(vs.TextMatrix(k1, 8)) - (Val(vs.TextMatrix(k1, 8)) * Val(vs.TextMatrix(k1, 11)) / 100), 0)
    vs.TextMatrix(k1, 13) = (Val(vs.TextMatrix(k1, 10)) - Val(vs.TextMatrix(k1, 12)))
    
    If (RS!scid = "") Then
       vs.TextMatrix(k1, 15) = "N"
       vs.TextMatrix(k1, 16) = "N"
    Else
       vs.TextMatrix(k1, 15) = RS!scid
       vs.TextMatrix(k1, 16) = RS!scname
    End If
    
    vs.TextMatrix(k1, 17) = RS!RepName
    
        
    vs.rows = vs.rows + 1
    k1 = k1 + 1

End If


RS.MoveNext

Wend

'=============================================================
'' Sale End
'=============================================================
'=============================================================

If (str10 <> "") Then
    str_10 = "SELECT fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,RATE,GrossAmt,DISCOUNT,Net,sum(Donqty) as Donqty,sum(Adjqty) as Adjqty,sum(Saleqty) as Saleqty ,sum(SaleRQty) as SaleRQty,subledger,Scid,scname,agentname as repname" & _
    " FROM adjustmentqry where Godown='C' and (" & str10 & ") and subledger='" & txtScId.text & "' and " & dateRet & " group by fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,RATE,GrossAmt,DISCOUNT,Net,subledger,Scid,scName,agentname"
Else
    str_10 = "SELECT fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,RATE,GrossAmt,DISCOUNT,Net,sum(Donqty) as Donqty,sum(Adjqty) as Adjqty,sum(Saleqty) as Saleqty ,sum(SaleRQty) as SaleRQty,subledger,Scid,scname,agentname as repname" & _
    " FROM adjustmentqry where Godown='C' and subledger='" & txtScId.text & "' and " & dateRet & " group by fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,RATE,GrossAmt,DISCOUNT,Net,subledger,Scid,scname,agentname"

End If


If RS.State = 1 Then RS.close
RS.Open str_10, con
While RS.EOF = False

Qty_ = 0
dqty = 0

If rs1.State = 1 Then rs1.close
rs1.Open "SELECT  sum(Donqty) as Donqty,sum(Adjqty) as Adjqty,sum(Saleqty) as Saleqty ,sum(SaleRQty) as SaleRQty" & _
" FROM adjustmentqry where (rate=" & RS!rate & " and bookcode='" & RS!Bookcode & "' and Godown='C' and subledger='" & txtScId.text & "' and INVOICENO=" & RS!invoiceNo & " and fyear='" & RS!fyear & "')", con

If Check1_donation.value = 0 Then
   
   If Not IsNull(rs1!Donqty) Then
      dqty = rs1!Donqty
   End If
   
   If Not IsNull(rs1!Adjqty) Then
      dqty = dqty + rs1!Adjqty
   End If
   
   If (RS!Godown = "C") Then
      If Not IsNull(rs1!SaleRQty) Then
          Qty_ = rs1!SaleRQty - dqty
      End If
   End If
   
Else
   
   If Not IsNull(rs1!Adjqty) Then
      dqty = dqty + rs1!Adjqty
   End If
   
   If (RS!Godown = "C") Then
      If Not IsNull(rs1!SaleRQty) Then
          Qty_ = rs1!SaleRQty - dqty
      End If
   End If
   
End If


If (Qty_ > 0) Then



    vs.TextMatrix(k1, 0) = RS!fyear
    vs.TextMatrix(k1, 1) = RS!Godown
    vs.TextMatrix(k1, 2) = RS!invoiceNo
    vs.TextMatrix(k1, 3) = RS!invoiceDate
    vs.TextMatrix(k1, 4) = RS!Bookcode
    vs.TextMatrix(k1, 5) = RS!Bookname
    vs.TextMatrix(k1, 6) = Qty_
    vs.TextMatrix(k1, 7) = RS!rate
    vs.TextMatrix(k1, 8) = RS!GrossAmt
    vs.TextMatrix(k1, 9) = RS!discount
    vs.TextMatrix(k1, 10) = RS!net
    If txtAdPer.text <> "" Then
       vs.TextMatrix(k1, 11) = Val(txtAdPer.text)
    Else
       vs.TextMatrix(k1, 11) = 0
    End If
    
    vs.TextMatrix(k1, 12) = Round(Val(vs.TextMatrix(k1, 8)) - (Val(vs.TextMatrix(k1, 8)) * Val(vs.TextMatrix(k1, 11)) / 100), 0)
    vs.TextMatrix(k1, 13) = (Val(vs.TextMatrix(k1, 10)) - Val(vs.TextMatrix(k1, 12)))
    
    If (RS!scid = "") Then
       vs.TextMatrix(k1, 15) = "N"
       vs.TextMatrix(k1, 16) = "N"
    Else
       vs.TextMatrix(k1, 15) = RS!scid
       vs.TextMatrix(k1, 16) = RS!scname
    End If
    
    vs.TextMatrix(k1, 17) = RS!RepName
    
        
    vs.rows = vs.rows + 1
    k1 = k1 + 1


End If


RS.MoveNext

Wend



'=============================================================



Dim qty, GTotal, net, adamt, diffnet
Dim r As Integer

qty = 0
GTotal = 0
net = 0
adamt = 0
diffnet = 0

For I = 1 To vs.rows - 1
   If (vs.TextMatrix(I, 6) <> "") Then
   
   If (vs.TextMatrix(I, 1) = "I" Or vs.TextMatrix(I, 1) = "C/M") Then
       qty = qty + Val(IIf(vs.TextMatrix(I, 6) = "", 0, vs.TextMatrix(I, 6)))
       GTotal = GTotal + Val(IIf(vs.TextMatrix(I, 8) = "", 0, vs.TextMatrix(I, 8)))
       net = net + Val(IIf(vs.TextMatrix(I, 10) = "", 0, vs.TextMatrix(I, 10)))
       adamt = adamt + Val(IIf(vs.TextMatrix(I, 12) = "", 0, vs.TextMatrix(I, 12)))
       diffnet = diffnet + Val(IIf(vs.TextMatrix(I, 13) = "", 0, vs.TextMatrix(I, 13)))
       r = r + 1
   Else
       r = r + 1
       
       qty = qty - Val(IIf(vs.TextMatrix(I, 6) = "", 0, vs.TextMatrix(I, 6)))
       GTotal = GTotal - Val(IIf(vs.TextMatrix(I, 8) = "", 0, vs.TextMatrix(I, 8)))
       net = net - Val(IIf(vs.TextMatrix(I, 10) = "", 0, vs.TextMatrix(I, 10)))
       adamt = adamt - Val(IIf(vs.TextMatrix(I, 12) = "", 0, vs.TextMatrix(I, 12)))
       diffnet = diffnet - Val(IIf(vs.TextMatrix(I, 13) = "", 0, vs.TextMatrix(I, 13)))
   
   End If
       
   End If
Next

txtQty.text = qty
txtGTotal.text = Round(GTotal, 0)
txtNetTotal.text = Round(net, 0)
txtAdjAmt.text = Round(adamt, 0)
txtDiffNet.text = Round(diffnet, 0)

lblrow.Caption = "Total Rows : " & r


MsgBox "Data View ...", vbInformation

End Sub
Sub fillschool()


Dim str_10 As String


Dim k1 As Integer
Dim Qty_, dqty As Integer

cboschool.Clear

con.Execute "delete from tmpTable2"

k1 = 1
Qty_ = 0
dqty = 0

str10 = ""

' For I = 0 To cboschool.ListCount - 1
' If cboschool.Selected(I) = True Then
'    If str10 = "" Then
'       str10 = "scname='" & IIf(cboschool.List(I) = "N", "", cboschool.List(I)) & "'"
'    Else
'       str10 = str10 & " or scname='" & IIf(cboschool.List(I) = "N", "", cboschool.List(I)) & "'"
'    End If
' End If
' Next

'If (cmbAgentName.Text <> "") Then
'   If (str10 = "") Then
'       str10 = "agentname='" & cmbAgentName.Text & "'"
'   Else
'       str10 = str10 & " and agentname='" & cmbAgentName.Text & "'"
'   End If
'End If


'=============================================================
'=============================================================

If (str10 <> "") Then
    str_10 = "SELECT fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,agentname,RATE,GrossAmt,DISCOUNT,Net,sum(Donqty) as Donqty,sum(Adjqty) as Adjqty,sum(Saleqty) as Saleqty ,sum(SaleRQty) as SaleRQty,subledger,Scid,scname" & _
    " FROM adjustmentqry where Godown='I' and " & str10 & " and subledger='" & txtScId.text & "' and " & dt_str & " group by fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,agentname,RATE,GrossAmt,DISCOUNT,Net,subledger,Scid,scName order by Godown desc"
Else
    str_10 = "SELECT fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,agentname,RATE,GrossAmt,DISCOUNT,Net,sum(Donqty) as Donqty,sum(Adjqty) as Adjqty,sum(Saleqty) as Saleqty ,sum(SaleRQty) as SaleRQty,subledger,Scid,scname" & _
    " FROM adjustmentqry where Godown='I' and subledger='" & txtScId.text & "' and " & dt_str & " group by fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,agentname,RATE,GrossAmt,DISCOUNT,Net,subledger,Scid,scname order by Godown desc"

End If


If RS.State = 1 Then RS.close
RS.Open str_10, con
While RS.EOF = False

Qty_ = 0
dqty = 0

If Check1_donation.value = 0 Then
   dqty = RS!Donqty
   dqty = dqty + RS!Adjqty
   If (RS!Godown = "I" Or RS!Godown = "C/M") Then
      Qty_ = RS!saleQty - dqty
   Else
      Qty_ = RS!SaleRQty - dqty
   End If
   
Else
   
   dqty = dqty + RS!Adjqty
   If (RS!Godown = "I" Or RS!Godown = "C/M") Then
      Qty_ = RS!saleQty - dqty
   Else
      Qty_ = RS!SaleRQty - dqty
   End If

End If


st_ = ""

If (RS!scid = "") Then
   st_ = "N"
Else
   st_ = RS!scname
End If



If (Qty_ > 0) Then
    con.Execute "insert into tmpTable2(scname,scid,repname) values('" & st_ & "','" & RS!agentname & "','" & UId & "')"
End If


RS.MoveNext

Wend

'=============================================================
'' Sale End
'=============================================================
'=============================================================

If (str10 <> "") Then
    str_10 = "SELECT fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,agentname,RATE,GrossAmt,DISCOUNT,Net,sum(Donqty) as Donqty,sum(Adjqty) as Adjqty,sum(Saleqty) as Saleqty ,sum(SaleRQty) as SaleRQty,subledger,Scid,scname" & _
    " FROM adjustmentqry where Godown='C' and " & str10 & " and subledger='" & txtScId.text & "' and " & dt_strR & " group by fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,agentname,RATE,GrossAmt,DISCOUNT,Net,subledger,Scid,scName order by Godown desc"
Else
    str_10 = "SELECT fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,agentname,RATE,GrossAmt,DISCOUNT,Net,sum(Donqty) as Donqty,sum(Adjqty) as Adjqty,sum(Saleqty) as Saleqty ,sum(SaleRQty) as SaleRQty,subledger,Scid,scname" & _
    " FROM adjustmentqry where Godown='C' and subledger='" & txtScId.text & "' and " & dt_strR & " group by fyear,Godown,INVOICENO,INVOICEDATE,BOOKCODE,bookname,agentname,RATE,GrossAmt,DISCOUNT,Net,subledger,Scid,scname order by Godown desc"

End If


If RS.State = 1 Then RS.close
RS.Open str_10, con
While RS.EOF = False

Qty_ = 0
dqty = 0

If Check1_donation.value = 0 Then
   dqty = RS!Donqty
   dqty = dqty + RS!Adjqty
   If (RS!Godown = "I" Or RS!Godown = "C/M") Then
      Qty_ = RS!saleQty - dqty
   Else
      Qty_ = RS!SaleRQty - dqty
   End If
   
Else
   
   dqty = dqty + RS!Adjqty
   If (RS!Godown = "I" Or RS!Godown = "C/M") Then
      Qty_ = RS!saleQty - dqty
   Else
      Qty_ = RS!SaleRQty - dqty
   End If

End If

If (Qty_ > 0) Then

st_ = ""

If (RS!scid = "") Then
   st_ = "N"
Else
   st_ = RS!scname
End If



If (Qty_ > 0) Then
    con.Execute "insert into tmpTable2(scname,scid,repname) values('" & st_ & "','" & RS!agentname & "','" & UId & "')"
End If


End If


RS.MoveNext

Wend

''----------------------------------
cboschool.Clear
cmbAgentName.Clear

If RS.State = 1 Then RS.close
RS.Open "select scname from tmpTable2 group by scname", con
While RS.EOF = False

cboschool.AddItem RS(0)

RS.MoveNext
Wend

If RS.State = 1 Then RS.close
RS.Open "select scid from tmpTable2 group by scid", con
While RS.EOF = False

cmbAgentName.AddItem RS(0)

RS.MoveNext
Wend


End Sub
Private Sub cmdView1_Click()

Dim scname As String
Dim scode As String
scode = ""
scname = ""

On Error GoTo err1


Screen.MousePointer = vbHourglass
HIT




Dim gps As String
Dim Qty_ As Long
bb_2 = False
gps = "n"
'=====================================================================



If rs1.State = 1 Then rs1.close
rs1.Open "select top 1 * from tmpSalesAdj where Scid='" & txtScId.text & "'", con, adOpenDynamic, adLockOptimistic
If rs1.EOF = False Then
    PopUpValue1 = ""
    PopUpValue2 = ""
    bb_2 = True
    GoTo dinesh:
End If


'===============New Coding Sale============================================
If RS.State = 1 Then RS.close

Set RS = con.Execute("exec SearchAdj_Qry '1','" & txtScId & "', '" & Sp_dt_from & "', '" & Sp_dt_to & "'")

While RS.EOF = False

If (RS!scname = "" Or Len(RS!scname) = 0) Then
scname = "N"
scode = "N"
Else
scname = RS!scname
scode = RS!scid
End If

Qty_ = 0

If RS!invoiceNo = 6887 Then
'MsgBox "a"
End If


If Check1_donation.value = 0 Then
    If rs1.State = 1 Then rs1.close
    Set rs1 = con.Execute("exec tmpDDetQry_new '" & RS!invoiceNo & "', '" & RS!Bookcode & "', '" & RS!fyear & "','" & UId & "','" & RS!rate & "'")
     If rs1.EOF = False Then
       Qty_ = RS!qty - rs1(0)
    Else
       Qty_ = RS!qty
    End If
Else
    Qty_ = RS!qty
End If


If rs1.State = 1 Then rs1.close
rs1.Open "select Qty FROM tmpSAdjDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='" & RS!vtype & "' and rate= '" & RS!rate & "' and fyear='" & RS!fyear & "'", con
If rs1.EOF = False Then
   Qty_ = Qty_ - rs1(0)
End If
If Qty_ > 0 Then
   con.Execute "insert into tmpSalesAdj(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname,scode) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & Qty_ & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & scname & "','" & RS!net & "','" & RS!discount & "','" & RS!vtype & "','" & RS!fyear & "','" & UserName & "','" & txtScId & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "','" & scode & "')"
End If
 RS.MoveNext
Wend
Qty_ = 0


'===============New Coding Sale Ret============================================
If RS.State = 1 Then RS.close

 RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,sum(QUANTITY) as QUANTITY," & _
"DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,NETAMOUNT," & _
"DISCOUNT,Godown,fyear,agentname,scid from PartyWiseItemWiseQtySales_Return " & _
"where subledger='" & txtScId & "' and " & dt_strR & "" & _
" group by INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,DESCFORINVOICE," & _
"ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName," & _
"NETAMOUNT,DISCOUNT,Godown,fyear,agentname,scid  order by INVOICENO,BOOKCODE", con


While RS.EOF = False

If (RS!scname = "" Or Len(RS!scname) = 0) Then
scname = "N"
scode = "N"
Else
scname = RS!scname
scode = RS!scid
End If


'Qty_ = RS!Quantity

Qty_ = 0

If Check1_donation.value = 0 Then
    If rs1.State = 1 Then rs1.close
    Set rs1 = con.Execute("exec tmpDDetQryRet_New '" & RS!invoiceNo & "', '" & RS!Bookcode & "', '" & RS!fyear & "', '" & RS!rate & "'")
    If rs1.EOF = False Then
       Qty_ = RS!QUANTITY - rs1(0)
    Else
       Qty_ = RS!QUANTITY
    End If
Else
    Qty_ = RS!QUANTITY
End If



If rs1.State = 1 Then rs1.close
rs1.Open "select sum(Qty) FROM tmpSAdjDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='C' and fyear='" & RS!fyear & "' AND RATE='" & RS!rate & "'", con
If Not IsNull(rs1(0)) Then
   Qty_ = Qty_ - rs1(0)
End If

If Qty_ > 0 Then
    con.Execute "insert into tmpSalesAdj(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname,scode) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & RS!QUANTITY & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & scname & "','" & RS!netamount & "','" & RS!discount & "','C','" & RS!fyear & "','" & UserName & "','" & txtScId & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "','" & scode & "')"
End If
RS.MoveNext
Wend


'==============================================================================
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'==============================================================================



If rs1.State = 1 Then rs1.close
rs1.Open "select fromDate,toDate,fyear,Current_Next,DataBase,fromDateSRet,toDateSRet from turnOverDis where NotCreated='y' and Current_Next='next' order by fyear", CCON
If rs1.EOF = False Then
'------Fatch Data From Next Session--------------------------------------'
    db_ = rs1!Database
    Set CON_next = New ADODB.Connection
    If LCase(server_) = "server" Then
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & db_ & "; UID=" & sql_user & "; PWD=" & sql_pass
       CON_next.Open
    Else
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=HP\SQL2008; DATABASE=" & db_ & "; UID=; PWD=;"
       CON_next.Open
    End If

    Qty_ = 0
    
    
    If RS.State = 1 Then RS.close
    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,agentname,scid,vtype from PartyWiseItemWiseQtySales where Subledger='" & txtScId & "' and " & dt_strSaleNext & " order by INVOICENO", CON_next
    
    While RS.EOF = False
     
    If (RS!scname = "" Or Len(RS!scname) = 0) Then
    scname = "N"
    scode = "N"
    Else
    scname = RS!scname
    scode = RS!scid
    End If
     
     
     '''''check donation
    If Check1_donation.value = 0 Then
        If rs1.State = 1 Then rs1.close
        Set rs1 = con.Execute("exec tmpDDetQry_new '" & RS!invoiceNo & "', '" & RS!Bookcode & "', '" & RS!fyear & "','" & UId & "','" & RS!rate & "'")
        
        If rs1.EOF = False Then
            Qty_ = RS!qty - rs1(0)
        Else
            Qty_ = RS!qty
        End If
     Else
            Qty_ = RS!qty
    End If
    
          
       
    If rs1.State = 1 Then rs1.close
    rs1.Open "select Qty FROM tmpSAdjDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='" & RS!vtype & "' and fyear='" & RS!fyear & "'", con
    If rs1.EOF = False Then
       Qty_ = Qty_ - rs1(0)
    End If
     
    If Qty_ > 0 Then
      con.Execute "insert into tmpSalesAdj(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname,scode) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & Qty_ & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & scname & "','" & RS!net & "','" & RS!discount & "','" & RS!vtype & "','" & RS!fyear & "','" & UserName & "','" & txtScId & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "','" & scode & "')"
    End If
    RS.MoveNext
    Wend
    
    
    'SaleReturn
    
    If RS.State = 1 Then RS.close
    
    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,sum(QUANTITY) as QUANTITY," & _
    "DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,NETAMOUNT," & _
    "DISCOUNT,Godown,fyear,agentname,scid from PartyWiseItemWiseQtySales_Return " & _
    "where subledger='" & txtScId & "' and " & dt_strSaleRNext & "" & _
    " group by INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,DESCFORINVOICE," & _
    "ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName," & _
    "NETAMOUNT,DISCOUNT,Godown,fyear,agentname,scid  order by INVOICENO,BOOKCODE", CON_next
    While RS.EOF = False
    
    
    If (RS!scname = "" Or Len(RS!scname) = 0) Then
    scname = "N"
    scode = "N"
    Else
    scname = RS!scname
    scode = RS!scid
    End If
    
       
    Qty_ = 0
    
     
    If Check1_donation.value = 0 Then
        If rs1.State = 1 Then rs1.close
        Set rs1 = con.Execute("exec tmpDDetQryRet_New '" & RS!invoiceNo & "', '" & RS!Bookcode & "', '" & RS!fyear & "', '" & RS!rate & "'")
        
        If rs1.EOF = False Then
           Qty_ = RS!QUANTITY - rs1(0)
        Else
           Qty_ = RS!QUANTITY
        End If
        
    Else
       Qty_ = RS!QUANTITY
    End If
   
       
   
       
    If rs1.State = 1 Then rs1.close
    rs1.Open "select sum(Qty) FROM tmpSAdjDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='C' and fyear='" & RS!fyear & "' AND RATE='" & RS!rate & "'", con
    If Not IsNull(rs1(0)) Then
       Qty_ = Qty_ - rs1(0)
       rs1.MoveNext
    End If
    
       
    If Qty_ > 0 Then
       con.Execute "insert into tmpSalesAdj(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname,scode) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & Qty_ & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & scname & "','" & RS!netamount & "','" & RS!discount & "','C','" & RS!fyear & "','" & UserName & "','" & txtScId & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "','" & scode & "')"
    End If
    RS.MoveNext
    Wend


'------end code------------------------------------------------------------------------------------'

'------end code------------------------------------------------------------------------------------'
End If
'------end code------------------------------------------------------------------------------------'
'==================================================================================================


din_1 = 0
cmbAgentName.Clear
cboschool.Clear

str1 = "SELECT repname from tmpSalesAdj where ScId='" & txtScId & "' group by repname"
If RS.State = 1 Then RS.close
RS.Open str1, con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
  cmbAgentName.AddItem RS(0)
  RS.MoveNext
  din_1 = 1
Wend

str1 = "SELECT scname,scode from tmpSalesAdj where ScId='" & txtScId & "' group by scname,scode"
If RS.State = 1 Then RS.close
RS.Open str1, con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
  cboschool.AddItem RS!scname
  RS.MoveNext
Wend






If din_1 = 1 Then
   cmbAgentName.ListIndex = 0
End If




fillGrid_
PopUpValue1 = ""
PopUpValue2 = ""



'End If

dinesh:

Screen.MousePointer = vbDefault


Exit Sub
err1:
Screen.MousePointer = vbDefault
MsgBox "" & err.DESCRIPTION


End Sub

Private Sub Command1_Click()
fillGrid_
End Sub

Private Sub Del_Click()
If MsgBox("want to delete ?", vbQuestion + vbYesNo) = vbYes Then
   
   createLog UserName, txtSponsorshipNo, "adjustment", " Delete : " & txtDiffNet, Date
   
   con.Execute "delete from SalesAdjustment where dno=" & txtSponsorshipNo & ""
   con.Execute "delete from SalesAdjustmentDet where dno=" & txtSponsorshipNo & ""
   con.Execute "delete from tmpSalesAdj where Sno=" & txtSponsorshipNo & ""
   
   con.Execute "delete FROM tmpSaladjust where EntryNo =" & txtSponsorshipNo & " "
   
   con.Execute "exec tmpdata " & UId & ""
   con.Execute "exec tmpdata_saleadj"
   
   
   If rs1.State = 1 Then rs1.close
   rs1.Open "select fromDate,toDate,fyear,Current_Next,DataBase,fromDateSRet,toDateSRet from turnOverDis where NotCreated='y' and Current_Next='next' order by fyear", CCON
   If rs1.EOF = False Then
   '------Fatch Data From Next Session--------------------------------------'
    db_ = rs1!Database
    Set CON_next = New ADODB.Connection
    If LCase(server_) = "server" Then
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & db_ & "; UID=" & sql_user & "; PWD=" & sql_pass
       CON_next.Open
    Else
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=HP\SQL2008; DATABASE=" & db_ & "; UID=; PWD=;"
       CON_next.Open
    End If

    
    CON_next.Execute "delete from tmpSalesAdj where (Sno=" & txtSponsorshipNo & " and subledger='" & txtScId.text & "')"
    CON_next.Execute "delete FROM tmpSaladjust where (EntryNo =" & txtSponsorshipNo & " and party='" & txtScId.text & "')"
    CON_next.Execute "exec tmpdata " & UId & ""
    CON_next.Execute "exec tmpdata_saleadj"
    
    
    End If
    
   
   
   
    refresh_
   
   
End If
End Sub
Sub refresh_()


lblsc.Visible = False
cboschool.Visible = False
lblPRemarks.Caption = ""
Check1_manullay.value = 0

vs.Enabled = True
Add = True
Edit = False
vs.Clear

txtManually.text = ""

Check1_manual.value = 0
txtNetBal = ""
txtWhomToBeGivenMob = ""
txtSponsorshipNo.text = ""
txtDates.value = Format(Date, "dd/MM/yyyy")
txtScId.text = ""
txtSchoolName.text = ""
cmbAgentName.ListIndex = -1
'cmbAgentName1.ListIndex = -1
'cboPayment.ListIndex = -1
'txtWhomTobeGiven.Text = ""
txtPrincipal = ""
txtMob = ""
txtRemarks = ""
txtGTotal = 0
txtNetTotal = 0
'cboSponse.ListIndex = -1
txtReturnAdj = 0
txtAmtAfterAdj = 0
'txtPercentSp.Text = 0
txtFAmt = 0

txtRoundOf = 0
txtAdvAmt = 0

txtAdjAmt = 0
txtDiffNet = 0

txtSponsorshipNo.SetFocus
save.Enabled = True
End Sub

Private Sub Form_Activate()
If RS.State = 1 Then RS.close
RS.Open "select sername from books group by sername", con
cboser.Clear

If Not RS.EOF Then
   Do While Not RS.EOF
      If IsNull(RS(0)) = False Then
        Me.cboser.AddItem RS(0)
      End If
      If Not RS.EOF Then RS.MoveNext
    Loop
End If
End Sub
Private Sub Form_Load()

On Error GoTo ee

Screen.MousePointer = vbHourglass

'If UCase(UserName) = "NADEEM" Then
'   cmdVew.Visible = True
'End If

kk1 = 1



If RS.State = 1 Then RS.close
RS.Open "select * from financialyear where fyear='" & session & "'", CCON
If RS.EOF = False Then
   dt_from = RS!fromdate
   dt_to = RS!todate
End If


fdate = ""
If RS.State = 1 Then RS.close
RS.Open "select fromDate,toDate from turnOverDis where fyear='" & session & "'", CCON
If RS.EOF = False Then
   'dt_str = "(INVOICEDATE >= convert(smalldatetime,'" & RS!FromDate & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!toDate & "',103))"
   fdate = RS!fromdate
End If


'1DEC-2018

'''''''If RS.State = 1 Then RS.close
''''''''RS.Open "select fromDate,toDate,NotCreated from turnOverDis where fyear='" & session_next & "'", CCON
'''''''RS.Open "select fromDate,toDate from SaleAdj_donnationDateRange", CCON
'''''''If RS.EOF = False Then
'''''''   dt_str = "(INVOICEDATE >= convert(smalldatetime,'" & RS!FromDate & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!toDate & "',103))"
'''''''   Sp_dt_from = RS!FromDate
'''''''   Sp_dt_to = RS!toDate
'''''''End If

'========================================================================

If RS.State = 1 Then RS.close
RS.Open "select fromDate,toDate,fromDateSRet,toDateSRet from turnOverDis where (fyear='" & session & "' and Current_Next='current')", CCON
If RS.EOF = False Then
   dt_str = "(INVOICEDATE >= convert(smalldatetime,'" & RS!fromdate & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!todate & "',103))"
   dt_strR = "(INVOICEDATE >= convert(smalldatetime,'" & RS!fromDateSRet & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!toDateSRet & "',103))"
   fdate = RS!fromdate
   
   fromDt_sale = RS!fromdate
   fromDt_saleret = RS!fromDateSRet
   
   Sp_dt_from = RS!fromdate
   Sp_dt_to = RS!todate
   
   toDt_sale = RS!todate
   toDt_saleret = RS!toDateSRet
   
End If


If RS.State = 1 Then RS.close
RS.Open "select fromDateSRet,toDateSRet,NotCreated,fromDate,toDate from turnOverDis where (Current_Next='next')", CCON
If RS.EOF = False Then
  If RS!NotCreated = "y" Then
   dt_strSaleNext = "(INVOICEDATE >= convert(smalldatetime,'" & RS!fromdate & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!todate & "',103))"
   dt_strSaleRNext = "(INVOICEDATE >= convert(smalldatetime,'" & RS!fromDateSRet & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!toDateSRet & "',103))"
   
   toDt_sale = RS!todate
   toDt_saleret = RS!toDateSRet
   
   
   
  End If
End If

'========================================================================

If Len(inviceNo) > 0 Then
   txtSponsorshipNo = inviceNo
   SearchDataNew inviceNo
   PopUpValue1 = ""
   PopUpValue2 = ""
   inviceNo = ""
   
   
    Screen.MousePointer = vbDefault
    '------------------
    Me.top = 0
    Me.Left = 0
    Me.Width = 14800
    Me.Height = 10435
    BackColorFrom Me

   
   Exit Sub
End If

'========================================================================


con.Execute "delete from tempLedger_net"
con.Execute "insert into tempLedger_net(Billtype,des) SELECT  ScID,ScName FROM INVOICEA where (len(ScID)>0 and len(ScName)>0) group by ScID,ScName"

con.Execute "delete from tmpSalesAdj where username='" & UserName & "'"


txtDates.value = Format(Date, "dd/MM/yyyy")
max_sp

'------------------
a1 = Left(session, 4) + 1
a2 = Right(session, 2) + 1
vs.ColComboList(0) = session & "|" & a1 & "-" & a2

db_ = Mid(a1, 3) & a2


If RS.State = 1 Then RS.close
RS.Open "select fyear,DataBase,Current_Next,NotCreated from turnOverDis where NotCreated='y' and current_next = 'next' order by fyear", CCON
If RS.EOF = False Then

    Set CON_next = New ADODB.Connection
    If LCase(server_) = "server" Then
       
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=chitraData_" & db_ & "; UID=" & sql_user & "; PWD=" & sql_pass
       CON_next.Open
       PopUpValue6 = ""
    Else
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=HP\SQL2008; DATABASE=chitraData_" & db_ & "; UID=; PWD=;"
       CON_next.Open
    End If



    If rs1.State = 1 Then rs1.close
    rs1.Open "select distinct Billtype,des from tempLedger_net", con, adOpenDynamic, adLockOptimistic
    If RS.State = 1 Then RS.close
    RS.Open "SELECT ScName,ScID FROM INVOICEA where len(ScName)>0 group by ScName,ScID", CON_next
    While RS.EOF = False
   
        rs1.MoveFirst
        rs1.Find "Billtype='" & RS!scid & "'"
        If rs1.EOF = True Then
            con.Execute "insert into tempLedger_net(des,Billtype) values('" & RS(0) & "','" & RS(1) & "')"
        End If
        RS.MoveNext
    Wend



End If




If RS.State = 1 Then RS.close
RS.Open "select sername from books group by sername", con
cboser.Clear

If Not RS.EOF Then
   Do While Not RS.EOF
      If IsNull(RS(0)) = False Then
        Me.cboser.AddItem RS(0)
      End If
      If Not RS.EOF Then RS.MoveNext
    Loop
End If







Screen.MousePointer = vbDefault
'------------------
Me.top = 0
Me.Left = 0
Me.Width = 14800
Me.Height = 10435
BackColorFrom Me


Exit Sub
ee:

Screen.MousePointer = vbDefault
'------------------
Me.top = 0
Me.Left = 0
Me.Width = 14400
Me.Height = 10435
BackColorFrom Me



End Sub

Private Sub List1_sc_DblClick()
txtSponsorshipNo = Trim(Mid(List1_sc.text, 1, InStr(List1_sc.text, "=") - 1))
searchData
End Sub

Private Sub save_Click()


If (Check1_manullay.value = 1) Then
   If (txtManually.text = "") Then
      
      MsgBox "Enter Remarks...", vbInformation
      txtManually.SetFocus
      Exit Sub
      
   End If
End If


If Edit = True Then
   ModifyData
Else
   saveData
End If


End Sub
Sub saveData()

On Error GoTo aa10

Dim rs_ As New ADODB.Recordset
Dim address1 As String
Dim address2 As String
Dim DESCFORINVOICE As String
Dim states As String
Dim distcode As String
Dim mobile As String

address1 = ""
address2 = ""
DESCFORINVOICE = ""
states = ""
distcode = ""
mobile = ""

rs_.Open "select ADDRESS1,ADDRESS2,DESCFORINVOICE,states,distcode,mobile from sledger where subledger='" & txtScId.text & "'", con
If rs_.EOF = False Then

address1 = rs_!address1 & ""
address2 = rs_!address2 & ""
DESCFORINVOICE = rs_!DESCFORINVOICE
states = rs_!states
distcode = rs_!distcode & ""
mobile = rs_!mobile & ""

End If

If txtScId = "" Then
   MsgBox "Select School Name ...", vbCritical
   txtSchoolName.SetFocus
   Exit Sub
End If




createLog UserName, txtSponsorshipNo, "adjustment", " Save/Edit : " & txtDiffNet, Date

If Edit = True Then
    con.Execute "delete from SalesAdjustmentDet where (DNo=" & txtSponsorshipNo & ")"
    con.Execute "delete from SalesAdjustment where (DNo=" & txtSponsorshipNo & ")"
    Edit = False
Else
    If Check1_manual.value = 0 Then
       max_sp
    End If
End If

Set RS = New ADODB.Recordset
RS.Open "select * from SalesAdjustment where DNo=" & txtSponsorshipNo & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
 RS.AddNew
 For k1 = 1 To vs.rows - 1
   If vs.TextMatrix(k1, 0) <> "" Then
      'con.Execute "insert into SalesAdjustmentDet(Dno,INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,id,repname,scode) select '" & txtSponsorshipNo & "' ,INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,id,repname,scode from tmpSalesAdj where (fyear='" & vs.TextMatrix(k1, 0) & "' and godown='" & vs.TextMatrix(k1, 1) & "' and invoiceno='" & vs.TextMatrix(k1, 2) & "' and bookcode='" & vs.TextMatrix(k1, 4) & "' and UserName='" & UserName & "' and REPNAME='" & cmbAgentName.Text & "' and rate='" & vs.TextMatrix(k1, 7) & "')"
      
      
      con.Execute "insert into SalesAdjustmentDet(Dno,INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME," & _
      "states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,RATE," & _
      "ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,id,repname,scode) " & _
      "values('" & txtSponsorshipNo & "' ,'" & vs.TextMatrix(k1, 2) & "','" & txtScId.text & "'," & _
      "'" & vs.TextMatrix(k1, 4) & "','" & vs.TextMatrix(k1, 5) & "','" & states & "','" & vs.TextMatrix(k1, 6) & "','" & DESCFORINVOICE & "','" & address1 & "'," & _
      "'" & address2 & "','" & distcode & "','" & mobile & "','" & Format(vs.TextMatrix(k1, 3), "MM/dd/yyyy") & "','" & vs.TextMatrix(k1, 8) & "','" & vs.TextMatrix(k1, 7) & "'," & _
      "'" & vs.TextMatrix(k1, 16) & "','" & vs.TextMatrix(k1, 10) & "','" & vs.TextMatrix(k1, 9) & "','" & vs.TextMatrix(k1, 1) & "','" & vs.TextMatrix(k1, 0) & "','" & UserName & "','" & txtScId.text & "'," & k1 & ",'" & vs.TextMatrix(k1, 17) & "','" & vs.TextMatrix(k1, 15) & "')"
   
   End If
  Next
End If

For k1 = 1 To vs.rows - 1
   If vs.TextMatrix(k1, 0) <> "" Then
   If vs.TextMatrix(k1, 11) <> "" Then
      con.Execute "update SalesAdjustmentDet set DISCOUNT_Adj=" & Val(vs.TextMatrix(k1, 11)) & ",Net_Adj=" & vs.TextMatrix(k1, 12) & ",Net_Diff=" & vs.TextMatrix(k1, 13) & " where (INVOICENO=" & vs.TextMatrix(k1, 2) & " and BOOKCODE='" & vs.TextMatrix(k1, 4) & "' and DNo=" & txtSponsorshipNo & " and rate='" & vs.TextMatrix(k1, 7) & "')"
   End If
   End If
Next




RS!dno = txtSponsorshipNo.text
RS!DDate = txtDates.value
RS!scid = txtScId.text
RS!scname = txtSchoolName.text
RS!RepName = cmbAgentName.text
RS!Principal = Trim(txtPrincipal)
RS!mobile = Trim(txtMob)
RS!remarks = Trim(txtRemarks)
RS!GrossAmt = Val(txtGTotal)
RS!net = Val(txtNetTotal)
RS!ReturnAdj = Val(txtReturnAdj)
RS!AmtAfter_ReturnAdj = Val(txtAmtAfterAdj)
RS!finalAmt = Val(txtFAmt)
RS!RoundOfAAmt = Val(txtRoundOf)
RS!NetBalance = Val(txtNetBal)
RS!MobileWhomtoGiven = Trim(txtWhomToBeGivenMob)
RS!Net_Adj = Val(txtAdjAmt)
RS!Net_Diff = Val(txtDiffNet)
RS!createdby = UserName
RS!MannuallyRem = txtManually.text

RS.update

save.Enabled = False


cmdEdit_4.Enabled = True
Add = False
Edit = False

MsgBox "Data Saved...", vbInformation

con.Execute "exec tmpdata " & UId & ""
con.Execute "exec tmpdata_saleadj"



Exit Sub
aa10:


MsgBox "" & err.DESCRIPTION
con.Execute "delete from SalesAdjustmentDet where (DNo=" & txtSponsorshipNo & ")"
con.Execute "delete from SalesAdjustment where (DNo=" & txtSponsorshipNo & ")"


End Sub
Sub ModifyData()

On Error GoTo save_



Dim rs_ As New ADODB.Recordset
Dim address1 As String
Dim address2 As String
Dim DESCFORINVOICE As String
Dim states As String
Dim distcode As String
Dim mobile As String

address1 = ""
address2 = ""
DESCFORINVOICE = ""
states = ""
distcode = ""
mobile = ""

rs_.Open "select ADDRESS1,ADDRESS2,DESCFORINVOICE,states,distcode,mobile from sledger where subledger='" & txtScId.text & "'", con
If rs_.EOF = False Then

address1 = rs_!address1 & ""
address2 = rs_!address2 & ""
DESCFORINVOICE = rs_!DESCFORINVOICE & ""
states = rs_!states
distcode = rs_!distcode
mobile = rs_!mobile & ""

End If



If txtScId = "" Then
   MsgBox "Select School Name ...", vbCritical
   txtSchoolName.SetFocus
   Exit Sub
End If


createLog UserName, txtSponsorshipNo, "adjustment", " Save/Edit : " & txtDiffNet, Date

If Edit = True Then
    
    Set RS = New ADODB.Recordset
    RS.Open "select * from SalesAdjustment where DNo=" & txtSponsorshipNo & "", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = False Then
       RS!dno = txtSponsorshipNo.text
       RS!DDate = txtDates.value
       RS!scid = txtScId.text
       RS!scname = txtSchoolName.text
       RS!RepName = cmbAgentName.text
       RS!Principal = Trim(txtPrincipal)
       RS!mobile = Trim(txtMob)
       RS!remarks = Trim(txtRemarks)
       RS!GrossAmt = Val(txtGTotal)
       RS!net = Val(txtNetTotal)
       RS!ReturnAdj = Val(txtReturnAdj)
       RS!AmtAfter_ReturnAdj = Val(txtAmtAfterAdj)
       RS!finalAmt = Val(txtFAmt)
       RS!RoundOfAAmt = Val(txtRoundOf)
       RS!NetBalance = Val(txtNetBal)
       RS!MobileWhomtoGiven = Trim(txtWhomToBeGivenMob)
       RS!Net_Adj = Val(txtAdjAmt)
       RS!Net_Diff = Val(txtDiffNet)
       RS!MannuallyRem = txtManually.text
       RS.update
    End If
    
    For k1 = 1 To vs.rows - 1
       If vs.TextMatrix(k1, 0) <> "" Then
       If vs.TextMatrix(k1, 11) <> "" Then
          
          Set rs1 = New ADODB.Recordset
          rs1.Open "select top 1 BOOKNAME from SalesAdjustmentDet where (INVOICENO=" & vs.TextMatrix(k1, 2) & " and BOOKCODE='" & vs.TextMatrix(k1, 4) & "' and DNo=" & txtSponsorshipNo & " and rate='" & vs.TextMatrix(k1, 7) & "')", con
          If rs1.EOF = False Then
            con.Execute "update SalesAdjustmentDet set Fyear='" & vs.TextMatrix(k1, 0) & "'" & _
            ",Godown='" & vs.TextMatrix(k1, 1) & "',INVOICENO=" & vs.TextMatrix(k1, 2) & "" & _
            ",invoiceDate='" & Format(vs.TextMatrix(k1, 3), "MM/dd/yyyy") & "'" & _
            ",BOOKCODE='" & vs.TextMatrix(k1, 4) & "'" & _
            ",BOOKNAME='" & vs.TextMatrix(k1, 5) & "'" & _
            ",qty=" & vs.TextMatrix(k1, 6) & "" & _
            ",RATE=" & vs.TextMatrix(k1, 7) & "" & _
            ",GrossAmt=" & vs.TextMatrix(k1, 8) & "" & _
            ",DISCOUNT=" & vs.TextMatrix(k1, 9) & "" & _
            ",Net=" & vs.TextMatrix(k1, 10) & "" & _
            ",DISCOUNT_Adj=" & vs.TextMatrix(k1, 11) & "" & _
            ",Net_Adj=" & vs.TextMatrix(k1, 12) & "" & _
            ",Net_Diff=" & vs.TextMatrix(k1, 13) & " " & _
            " where (INVOICENO='" & vs.TextMatrix(k1, 2) & "' and BOOKCODE='" & vs.TextMatrix(k1, 4) & "' and DNo=" & txtSponsorshipNo & " and id=" & vs.TextMatrix(k1, 14) & ")"
          Else
          
            'con.Execute "insert into SalesAdjustmentDet(Dno,INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,id,repname,DISCOUNT_Adj,Net_Adj,Net_Diff) select '" & txtSponsorshipNo & "' ,INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,id,repname,'" & vs.TextMatrix(k1, 11) & "','" & vs.TextMatrix(k1, 12) & "','" & vs.TextMatrix(k1, 13) & "' from tmpSalesAdj where (fyear='" & vs.TextMatrix(k1, 0) & "' and godown='" & vs.TextMatrix(k1, 1) & "' and invoiceno='" & vs.TextMatrix(k1, 2) & "' and bookcode='" & vs.TextMatrix(k1, 4) & "' and UserName='" & UserName & "' and REPNAME='" & cmbAgentName.Text & "')"
            
            con.Execute "insert into SalesAdjustmentDet(Dno,INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME," & _
            "states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,RATE," & _
            "ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,id,repname,scode) " & _
            "values('" & txtSponsorshipNo & "' ,'" & vs.TextMatrix(k1, 2) & "','" & txtScId.text & "'," & _
            "'" & vs.TextMatrix(k1, 4) & "','" & vs.TextMatrix(k1, 5) & "','" & states & "','" & vs.TextMatrix(k1, 6) & "','" & DESCFORINVOICE & "','" & address1 & "'," & _
            "'" & address2 & "','" & distcode & "','" & mobile & "','" & Format(vs.TextMatrix(k1, 3), "MM/dd/yyyy") & "','" & vs.TextMatrix(k1, 8) & "','" & vs.TextMatrix(k1, 7) & "','" & vs.TextMatrix(k1, 16) & "'" & _
            ",'" & vs.TextMatrix(k1, 10) & "','" & vs.TextMatrix(k1, 9) & "','" & vs.TextMatrix(k1, 1) & "','" & vs.TextMatrix(k1, 0) & "','" & UserName & "','" & txtScId.text & "'," & k1 & ",'" & vs.TextMatrix(k1, 17) & "','" & vs.TextMatrix(k1, 15) & "')"

            
          
          End If
          
       End If
       End If
    Next

    
    MsgBox "Data modify...", vbInformation
Else
    
    max_sp
    Set RS = New ADODB.Recordset
    RS.Open "select * from SalesAdjustment where DNo=" & txtSponsorshipNo & "", con, adOpenDynamic, adLockOptimistic
    If RS.EOF = True Then
       RS.AddNew
       For k1 = 1 To vs.rows - 1
       If vs.TextMatrix(k1, 0) <> "" Then
          con.Execute "insert into SalesAdjustmentDet(Dno,INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,id,repname) select '" & txtSponsorshipNo & "' ,INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,id,repname from tmpSalesAdj where (fyear='" & vs.TextMatrix(k1, 0) & "' and godown='" & vs.TextMatrix(k1, 1) & "' and invoiceno='" & vs.TextMatrix(k1, 2) & "' and bookcode='" & vs.TextMatrix(k1, 4) & "' and UserName='" & UserName & "' and REPNAME='" & cmbAgentName.text & "')"
       End If
       Next
    End If

    For k1 = 1 To vs.rows - 1
       If vs.TextMatrix(k1, 0) <> "" Then
       If vs.TextMatrix(k1, 11) <> "" Then
          con.Execute "update SalesAdjustmentDet set DISCOUNT_Adj=" & Val(vs.TextMatrix(k1, 11)) & ",Net_Adj=" & vs.TextMatrix(k1, 12) & ",Net_Diff=" & vs.TextMatrix(k1, 13) & " where (INVOICENO=" & vs.TextMatrix(k1, 2) & " and BOOKCODE='" & vs.TextMatrix(k1, 4) & "' and DNo=" & txtSponsorshipNo & ")"
       End If
       End If
    Next

    RS!dno = txtSponsorshipNo.text
    RS!DDate = txtDates.value
    RS!scid = txtScId.text
    RS!scname = txtSchoolName.text
    RS!RepName = cmbAgentName.text
    RS!Principal = Trim(txtPrincipal)
    RS!mobile = Trim(txtMob)
    RS!remarks = Trim(txtRemarks)
    RS!GrossAmt = Val(txtGTotal)
    RS!net = Val(txtNetTotal)
    RS!ReturnAdj = Val(txtReturnAdj)
    RS!AmtAfter_ReturnAdj = Val(txtAmtAfterAdj)
    RS!finalAmt = Val(txtFAmt)
    RS!RoundOfAAmt = Val(txtRoundOf)
    RS!NetBalance = Val(txtNetBal)
    RS!MobileWhomtoGiven = Trim(txtWhomToBeGivenMob)
    RS!Net_Adj = Val(txtAdjAmt)
    RS!Net_Diff = Val(txtDiffNet)
    RS.update
    
    MsgBox "Data Saved...", vbInformation
End If

 

save.Enabled = False
cmdEdit_4.Enabled = True
Add = False
Edit = False



con.Execute "exec tmpdata " & UId & ""
con.Execute "exec tmpdata_saleadj"



Exit Sub
save_:
MsgBox "" & err.DESCRIPTION

End Sub

Private Sub txtAdvAmt_Change()
calAmt
End Sub

Private Sub txtAdPer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   cmdOK.SetFocus
End If
End Sub

Private Sub txtDates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
txtSchoolName.SetFocus
End If
End Sub

Private Sub txtFAmt_Change()
txtRoundOf = Val(txtFAmt) - Val(txtNetBal)
End Sub

Private Sub txtMob_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtRemarks.SetFocus
End Sub

Private Sub txtNetBal_Change()
 txtRoundOf = Val(txtFAmt) - Val(txtNetBal)
End Sub

Private Sub txtPercentSp_Change()
calAmt
End Sub

Private Sub txtPercentSp_GotFocus()
HIT
End Sub

Private Sub txtPrincipal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtMob.SetFocus

End Sub

Private Sub txtPrincipal_LostFocus()
txtPrincipal = UCase(txtPrincipal)
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then cboSponse.SetFocus
End Sub

Private Sub txtRemarks_LostFocus()
txtRemarks = UCase(txtRemarks)
End Sub
Private Sub txtReturnAdj_Change()
calAmt
End Sub
Sub calAmt()
Dim sum1 As Double
Dim AmtretAdj As Double
Dim AmtSp As Double

On Error Resume Next

sum1 = 0
AmtretAdj = 0
AmtSp = 0

If cboSponse.text = "Gross" Then
   sum1 = txtGTotal
Else
   sum1 = txtNetTotal
End If

AmtretAdj = Round((sum1 * Val(txtReturnAdj) / 100), 0)

txtAmtAfterAdj.text = sum1 - AmtretAdj

txtFAmt = Round((Val(txtAmtAfterAdj.text) * Val(txtPercentSp) / 100), 0)

End Sub

Private Sub txtReturnAdj_GotFocus()
HIT
End Sub

Private Sub txtReturnAdj_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtPercentSp.SetFocus
End Sub

Private Sub txtRoundOf_Change()
'calAmt
End Sub
Private Sub txtSchoolName_GotFocus()

Dim scname As String
Dim scode As String
scode = ""
scname = ""
''On Error GoTo err1


'Screen.MousePointer = vbHourglass
HIT

If PopUpValue1 <> "" Then
'vs.Clear

txtSchoolName.text = PopUpValue1
If RS.State = 1 Then RS.close
RS.Open "select SUBLEDGER,partyremarks from SLEDGER where code='" & PopUpValue3 & "'", con
If RS.EOF = False Then
   txtScId = RS(0)
   lblPRemarks.Caption = RS!PartyRemarks & ""
End If
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""


cmdView1_Click

End If


End Sub
Sub fillGrid_()


On Error GoTo err1

vs.Clear

'Dim str_ As String


txtGTotal = 0
txtNetTotal = 0
txtQty = 0

txtAdjAmt = 0
txtDiffNet = 0


con.Execute "UPDATE a SET a.DISCOUNT_Adj = b.adj  FROM tmpSalesAdj AS a " & _
  "INNER JOIN BookWisePartyWiseAppAdj_Qry AS b ON (a.BOOKCODE = b.BOOKCODE and a.SUBLEDGER = b.SUBLEDGER)"
 



If RS.State = 1 Then RS.close
If txtSponsorshipNo = "" Then
   Exit Sub
End If

If Check1_sc.value = 0 Then

    If cmbAgentName.text = "" Then
     If txtScId = "" Then
        Exit Sub
     End If
      RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QTY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,scid,DISCOUNT_Adj,scode from tmpSalesAdj where userNAME='" & UserName & "' and sno=" & txtSponsorshipNo & " and ScId='" & txtScId & "' order by id", con
    Else
       RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QTY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,scid,DISCOUNT_Adj,scode from tmpSalesAdj where repname='" & cmbAgentName & "' and userNAME='" & UserName & "' and sno=" & txtSponsorshipNo & " order by id", con
    End If
Else

If cmbAgentName.text = "" Then
 If txtScId = "" Then
    Exit Sub
 End If
   If str_ <> "" Then
      RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QTY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,scid,DISCOUNT_Adj,scode from tmpSalesAdj where userNAME='" & UserName & "' and sno=" & txtSponsorshipNo & " and ScId='" & txtScId & "' and " & str_ & " order by id", con
   Else
      RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QTY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,scid,DISCOUNT_Adj,scode from tmpSalesAdj where userNAME='" & UserName & "' and sno=" & txtSponsorshipNo & " and ScId='" & txtScId & "'  order by id", con
   End If
Else
   If str_ <> "" Then
     RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QTY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,scid,DISCOUNT_Adj,scode from tmpSalesAdj where repname='" & cmbAgentName & "' and userNAME='" & UserName & "' and sno=" & txtSponsorshipNo & " and " & str_ & " order by id", con
   Else
     RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QTY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,scid,DISCOUNT_Adj,scode from tmpSalesAdj where userNAME='" & UserName & "' and sno=" & txtSponsorshipNo & "  order by id", con
   End If
End If



End If


If RS.EOF = False Then
vs.rows = RS.RecordCount + 50
End If
For I = 1 To RS.RecordCount

vs.TextMatrix(I, 0) = RS!fyear
vs.TextMatrix(I, 1) = RS!Godown
vs.TextMatrix(I, 2) = RS!invoiceNo
vs.TextMatrix(I, 3) = RS!invoiceDate
vs.TextMatrix(I, 4) = RS!Bookcode
vs.TextMatrix(I, 5) = RS!Bookname
vs.TextMatrix(I, 6) = RS!qty
vs.TextMatrix(I, 7) = RS!rate
vs.TextMatrix(I, 8) = RS!GrossAmt
vs.TextMatrix(I, 9) = RS!discount
vs.TextMatrix(I, 10) = Round(RS!net, 0)
vs.TextMatrix(I, 11) = IIf(IsNull(RS!DISCOUNT_Adj), 0, Round(RS!DISCOUNT_Adj, 0))
vs.TextMatrix(I, 15) = RS!scode & ""

calcAdjNet (I)

'========================================================
If rs1.State = 1 Then rs1.close
code1 = Trim(Mid(txtScId.text, 6))
Set rs1 = con.Execute("exec SearchAppDataNew '" & RS!scname & "','" & code1 & "','" & RS!Bookcode & "'")
If rs1.EOF = False Then
       
    a11 = (rs1!appper + rs1!discount + rs1!adj)
    vs.TextMatrix(I, 11) = Val(a11)
    vs.TextMatrix(I, 18) = Val(a11)
    
    vs.TextMatrix(I, 12) = Round(Val(vs.TextMatrix(I, 8)) - (Val(vs.TextMatrix(I, 8)) * Val(a11) / 100), 0)
    vs.TextMatrix(I, 13) = (Val(vs.TextMatrix(I, 10)) - Val(vs.TextMatrix(I, 12)))
Else
    vs.TextMatrix(I, 18) = 0

End If


'========================================================

If (RS!Godown = "I" Or RS!Godown = "C/M") Then
    txtGTotal = Val(txtGTotal) + RS!GrossAmt
    txtNetTotal = Val(txtNetTotal) + Round(RS!net, 0)
    txtAdjAmt = Val(txtAdjAmt) + Val(vs.TextMatrix(I, 12))
    txtDiffNet = Val(txtDiffNet) + Val(vs.TextMatrix(I, 13))
Else
    txtGTotal = Val(txtGTotal) - RS!GrossAmt
    txtNetTotal = Val(txtNetTotal) - Round(RS!net, 0)
    txtAdjAmt = Val(txtAdjAmt) - Val(vs.TextMatrix(I, 12))
    txtDiffNet = Val(txtDiffNet) - Val(vs.TextMatrix(I, 13))
    
End If



txtQty = Val(txtQty) + RS!qty

RS.MoveNext

Next

txtGTotal = Round(txtGTotal, 0)


vs.FormatString = "Session|Inv.Type|Inv. No.|Inv. Date|B.Code|B.Name|>Quantity|>Rate|>GrossAmt|>Dis.%|>NetAmt|>Adj.Dis(%)|>Adj.Net|>Diff.NetAmt||SCode"

vs.ColWidth(1) = 700
vs.ColWidth(2) = 700
vs.ColWidth(3) = 900
vs.ColWidth(4) = 700
vs.ColWidth(5) = 2600
vs.ColWidth(6) = 800
vs.ColWidth(7) = 700
vs.ColWidth(8) = 950

vs.ColWidth(9) = 600
vs.ColWidth(10) = 1000

vs.ColWidth(11) = 1000
vs.ColWidth(12) = 900
vs.ColWidth(13) = 1000
vs.ColWidth(14) = 0
vs.ColWidth(15) = 600
vs.ColWidth(16) = 0
vs.ColWidth(17) = 0

vs.ColWidth(18) = 0



Total

Exit Sub

err1:
MsgBox "" & err.DESCRIPTION


End Sub
Private Sub txtSchoolName_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
        
   Screen.MousePointer = vbHourglass

   searchType = "party"
   value = "SELECT  DESCFORINVOICE as PartyName,address3 as City,Code FROM SLEDGER where len(DESCFORINVOICE)>0 order by DESCFORINVOICE"
   popuplist_client value, con
   set_focus = True
   
   Screen.MousePointer = vbDefault

End If


If KeyCode = 13 Then
   'If MsgBox("Want to View ?", vbInformation + vbYesNo) = vbYes Then
      cmdView1_Click
   'End If
End If


End Sub
Private Sub txtSchoolName_LostFocus()

'''--------
''cboschool.Clear
''str_ = "SELECT scname FROM adjustmentqry where subledger='" & txtscid.Text & "' and " & dt_str & " group by scname order by scname"
''If RS.State = 1 Then RS.close
''RS.Open str_, con
''While RS.EOF = False
''  If (RS!scname = "") Then
''     cboschool.AddItem "N"
''  Else
''     cboschool.AddItem RS!scname
''  End If
''
''  RS.MoveNext
''Wend
''
''cmbAgentName.Clear
''str_ = "SELECT agentname FROM adjustmentqry where agentname is not null and subledger='" & txtscid.Text & "' and " & dt_str & " group by agentname order by agentname"
''If RS.State = 1 Then RS.close
''RS.Open str_, con
''While RS.EOF = False
''  cmbAgentName.AddItem RS!agentname
''  RS.MoveNext
''Wend

'---------
'If txtscid.Text <> "" Then
'   fillschool
'End If


End Sub

Private Sub txtSponsorshipNo_GotFocus()
  If PopUpValue1 <> "" Then
     txtSponsorshipNo = PopUpValue1
     
     searchData
     PopUpValue1 = ""
     PopUpValue2 = ""
  End If
End Sub
Private Sub txtSponsorshipNo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    searchType = "inv"
    popuplist10 "select DNo,DDate,ScName,RoundOfAAmt from SalesAdjustment order by DNo", con
End If

If KeyCode = 13 Then
     'If MsgBox("Want to Edid... ", vbQuestion + vbYesNo) = vbYes Then
     '   CON.Execute "insert into tmpSalesAdj(INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,sNo) select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,DNo from DonnationMainDet where dno=" & txtSponsorshipNo & ""
     'End If

   searchData
   txtDates.SetFocus
End If
End Sub

Private Sub txtWhomTobeGiven_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtWhomToBeGivenMob.SetFocus
End Sub

Private Sub txtWhomTobeGiven_LostFocus()
txtWhomTobeGiven = UCase(txtWhomTobeGiven)
End Sub
Sub Total()

On Error Resume Next


txtGTotal = 0
txtNetTotal = 0
txtQty = 0
txtAdjAmt = 0
txtDiffNet = 0
lblrow.Caption = ""

Dim r As Integer

r = 0

For I = 1 To vs.rows - 1
If vs.TextMatrix(I, 1) <> "" Then
If (vs.TextMatrix(I, 1) = "I" Or vs.TextMatrix(I, 1) = "C/M") Then
    txtGTotal = Val(txtGTotal) + vs.TextMatrix(I, 8)
    txtNetTotal = Val(txtNetTotal) + vs.TextMatrix(I, 10)
    txtQty = Val(txtQty) + Val(vs.TextMatrix(I, 6))
    txtAdjAmt = Val(txtAdjAmt) + Val(vs.TextMatrix(I, 12))
    txtDiffNet = Val(txtDiffNet) + Val(vs.TextMatrix(I, 13))
    r = r + 1
Else
    txtGTotal = Val(txtGTotal) - vs.TextMatrix(I, 8)
    txtNetTotal = Val(txtNetTotal) - vs.TextMatrix(I, 10)
    txtQty = Val(txtQty) - Val(vs.TextMatrix(I, 6))
    txtAdjAmt = Val(txtAdjAmt) - Val(vs.TextMatrix(I, 12))
    txtDiffNet = Val(txtDiffNet) - Val(vs.TextMatrix(I, 13))
    r = r + 1
    
End If

End If
Next


lblrow.Caption = "Total Rows : " & r

txtNetTotal.text = Round(txtNetTotal.text, 0)
txtGTotal = Round(txtGTotal, 0)
txtDiffNet = Round(txtDiffNet, 0)

End Sub

Private Sub txtWhomToBeGivenMob_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtPrincipal.SetFocus
End Sub
Private Sub txtSponsorshipNo_LostFocus()

If txtSponsorshipNo.text <> "" Then
searchData
End If

End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error Resume Next

If KeyCode = 115 Then
   con.Execute "delete from tmpSalesAdj where BOOKCODE='" & vs.TextMatrix(vs.RowSel, 4) & "' and INVOICENO=" & vs.TextMatrix(vs.RowSel, 2) & ""
   con.Execute "delete from SalesAdjustmentDet where DNo=" & txtSponsorshipNo & " and BOOKCODE='" & vs.TextMatrix(vs.RowSel, 4) & "' and INVOICENO=" & vs.TextMatrix(vs.RowSel, 2) & ""
   
   vs.RemoveItem vs.RowSel
   Total
   calAmt
End If

End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
  
On Error GoTo aa10
  
  
  Dim dis_ As Double
  Dim tblName As String
  
  
  If KeyCode = 13 Then
     
     If (vs.Col = 0 Or vs.Col = 1 Or vs.Col = 2 Or vs.Col = 3) Then
        
           
       If (Len(vs.TextMatrix(vs.RowSel, 0)) > 0 And Len(vs.TextMatrix(vs.RowSel, 1)) > 0 And Len(vs.TextMatrix(vs.RowSel, 2)) > 0) Then
         tblName = fatchDate(vs.TextMatrix(vs.RowSel, 0), vs.TextMatrix(vs.RowSel, 1), vs.TextMatrix(vs.RowSel, 2), 0)
         
         If MsgBox("Want to Add Bill", vbQuestion + vbYesNo) = vbYes Then
            
          If (vs.TextMatrix(vs.RowSel, 1) = "I" Or vs.TextMatrix(I, 1) = "C/M") Then
            
                If current_next = "current" Then
                    con.Execute "insert into tmpSalesAdj(INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,fyear,UserName,ScId,SNO,repname,scode) select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,'I','" & session & "','" & UserName & "',scid,'" & txtSponsorshipNo & "',agentname,scode from PartyWiseItemWiseQtySales where INVOICENO=" & vs.TextMatrix(vs.RowSel, 2) & ""
                Else
                    If RS.State = 1 Then RS.close
                    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,scid,agentname,scode from PartyWiseItemWiseQtySales where INVOICENO='" & vs.TextMatrix(vs.RowSel, 2) & "' and " & dt_strSaleNext & "", CON_next
                    While RS.EOF = False
                      con.Execute "insert into tmpSalesAdj(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname,scode) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & RS!qty & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!net & "','" & RS!discount & "','I','" & RS!fyear & "','" & UserName & "','" & RS!scid & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "','" & RS!scode & "')"
                      RS.MoveNext
                    Wend

                End If
            
                fillGrid_
           Else
            
                If vs.TextMatrix(vs.RowSel, 0) = session Then
                    'con.Execute "insert into tmpSalesAdj select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QUANTITY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,NETAMOUNT,DISCOUNT,'C','" & session & "','" & UserName & "',scid,'','" & gps & "','" & txtSponsorshipNo & "' from PartyWiseItemWiseQtySales_Return where INVOICENO=" & vs.TextMatrix(vs.RowSel, 2) & ""
                    con.Execute "insert into tmpSalesAdj(INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,fyear,UserName,ScId,SNO,repname,scode) select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Quantity,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,netamount,DISCOUNT,'C','" & session & "','" & UserName & "',scid,'" & txtSponsorshipNo & "',agentname,scid from PartyWiseItemWiseQtySales_Return where INVOICENO=" & vs.TextMatrix(vs.RowSel, 2) & ""
                 Else
                    If RS.State = 1 Then RS.close
                    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QUANTITY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,NETAMOUNT,DISCOUNT,Godown,fyear,scid,agentname from PartyWiseItemWiseQtySales_Return where INVOICENO='" & vs.TextMatrix(vs.RowSel, 2) & "' and " & dt_strSaleRNext & "", CON_next
                    While RS.EOF = False
                     con.Execute "insert into tmpSalesAdj(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname,scode) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & RS!QUANTITY & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!netamount & "','" & RS!discount & "','C','" & RS!fyear & "','" & UserName & "','" & RS!scid & "','addnew','" & txtSponsorshipNo & "','" & RS!agentname & "','" & RS!scid & "')"
                     RS.MoveNext
                    Wend

                 End If
                 fillGrid_
           End If

         End If
     
         
       End If
       
       sendkeys "{right}"
        
     ElseIf vs.Col = 4 Then
     

   '''================================================================
   '''Checking
   '''================================================================
    
   Select Case vs.TextMatrix(vs.RowSel, 0)
       
   Case session
    If (vs.TextMatrix(vs.RowSel, 1) = "I" Or vs.TextMatrix(vs.RowSel, 1) = "C/M") Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select top 1 INVOICENO from PartyWiseItemWiseQtySales where INVOICENO='" & vs.TextMatrix(vs.RowSel, 2) & "' and bookcode='" & vs.TextMatrix(vs.RowSel, 4) & "'", con
        If rs1.EOF = True Then
           MsgBox "This Book is not exist in this invoice...", vbCritical
           vs.SetFocus
           Exit Sub
        End If
    ElseIf vs.TextMatrix(vs.RowSel, 1) = "C" Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select * from PartyWiseItemWiseQtySales_return where INVOICENO='" & vs.TextMatrix(vs.RowSel, 2) & "' and bookcode='" & vs.TextMatrix(vs.RowSel, 4) & "'", con
        If rs1.EOF = True Then
           MsgBox "This Book is not exist in this invoice...", vbCritical
           vs.SetFocus
           Exit Sub
        End If
    End If
    
    
    Case session_next
    
      If (vs.TextMatrix(vs.RowSel, 1) = "I" Or vs.TextMatrix(vs.RowSel, 1) = "C/M") Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select top 1 INVOICENO from PartyWiseItemWiseQtySales where INVOICENO='" & vs.TextMatrix(vs.RowSel, 2) & "' and bookcode='" & vs.TextMatrix(vs.RowSel, 4) & "'", CON_next
        If rs1.EOF = True Then
           MsgBox "This Book is not exist in this invoice...", vbCritical
           vs.SetFocus
           Exit Sub
        End If
    ElseIf vs.TextMatrix(vs.RowSel, 1) = "C" Then
        If rs1.State = 1 Then rs1.close
        rs1.Open "select top 1 INVOICENO from PartyWiseItemWiseQtySales_return where INVOICENO='" & vs.TextMatrix(vs.RowSel, 2) & "' and bookcode='" & vs.TextMatrix(vs.RowSel, 4) & "'", CON_next
        If rs1.EOF = True Then
           MsgBox "This Book is not exist in this invoice...", vbCritical
           vs.SetFocus
           Exit Sub
        End If
    End If
    
    
    End Select
    
   '================================================================
   '================================================================
     
        If RS.State = 1 Then RS.close
        RS.Open "select RATE,DISCOUNT,Bookname from BOOKS where BOOKCODE='" & vs.TextMatrix(vs.RowSel, 4) & "'", con
        If RS.EOF = False Then
           vs.TextMatrix(vs.RowSel, 4) = UCase(vs.TextMatrix(vs.RowSel, 4))
           vs.TextMatrix(vs.RowSel, 5) = RS!Bookname
           vs.TextMatrix(vs.RowSel, 7) = RS!rate
           vs.TextMatrix(vs.RowSel, 9) = RS!discount
           sendkeys "{right}"
           sendkeys "{right}"
        End If
      ElseIf (vs.Col = 6 Or vs.Col = 7 Or vs.Col = 8) Then
          If (Val(vs.TextMatrix(vs.RowSel, 6)) > 0 And Val(txtSponsorshipNo) > 0 And vs.TextMatrix(vs.RowSel, 2) <> "" And vs.TextMatrix(vs.RowSel, 4) <> "") Then
           con.Execute "update tmpSalesAdj set Qty=" & Val(vs.TextMatrix(vs.RowSel, 6)) & " where SNO=" & txtSponsorshipNo & " and invoiceno=" & Val(vs.TextMatrix(vs.RowSel, 2)) & " and BOOKCODE='" & vs.TextMatrix(vs.RowSel, 4) & "'"
          End If
           
           vs.TextMatrix(vs.RowSel, 8) = Round((Val(vs.TextMatrix(vs.RowSel, 6)) * vs.TextMatrix(vs.RowSel, 7)), 0)
           dis_ = Round((vs.TextMatrix(vs.RowSel, 8) * vs.TextMatrix(vs.RowSel, 9) / 100), 0)
           vs.TextMatrix(vs.RowSel, 10) = Round(Val(vs.TextMatrix(vs.RowSel, 8)) - dis_, 0)
           sendkeys "{right}"
           '''Total
      ElseIf (vs.Col = 9) Then
      
           If MsgBox("Want to Add ?", vbQuestion + vbYesNo) = vbYes Then
            vs.TextMatrix(vs.RowSel, 11) = "addnew"
            vs.SetFocus
            vs.TextMatrix(vs.RowSel + 1, 0) = vs.TextMatrix(vs.RowSel, 0)
            
            sendkeys "{home}"
            sendkeys "{down}"
            '''Total
           End If
      
      ElseIf (vs.Col = 11) Then
      
      
        If (Val(vs.TextMatrix(vs.RowSel, 18)) > 0) Then
            If (Val(vs.TextMatrix(vs.RowSel, 11)) > Val(vs.TextMatrix(vs.RowSel, 18))) Then
               vs.TextMatrix(vs.RowSel, 11) = vs.TextMatrix(vs.RowSel, 18)
               MsgBox "You can set mannually Less Disscount Only...", vbCritical
               Exit Sub
            End If
        End If
          
          
          
      
           sendkeys "{down}"
           'ss = (Val(vs.TextMatrix(vs.RowSel, 8)) * Val(vs.TextMatrix(vs.RowSel, 11)) / 100)
           vs.TextMatrix(vs.RowSel, 12) = Round(Val(vs.TextMatrix(vs.RowSel, 8)) - (Val(vs.TextMatrix(vs.RowSel, 8)) * Val(vs.TextMatrix(vs.RowSel, 11)) / 100), 0)
           vs.TextMatrix(vs.RowSel, 13) = (Val(vs.TextMatrix(vs.RowSel, 10)) - Val(vs.TextMatrix(vs.RowSel, 12)))
           
           txtAdjAmt = 0
           txtDiffNet = 0
           
           For k1 = 1 To vs.rows - 1
           If vs.TextMatrix(k1, 0) <> "" Then
           If (vs.TextMatrix(k1, 1) = "I" Or vs.TextMatrix(k1, 1) = "C/M") Then
              txtAdjAmt = Val(txtAdjAmt) + Val(vs.TextMatrix(k1, 12))
              txtDiffNet = Val(txtDiffNet) + Val(vs.TextMatrix(k1, 13))
           Else
              txtAdjAmt = Val(txtAdjAmt) - Val(vs.TextMatrix(k1, 12))
              txtDiffNet = Val(txtDiffNet) - Val(vs.TextMatrix(k1, 13))
           End If
           End If
           Next
           
          
     End If
     
     Total
     
  End If
  
  
Exit Sub
aa10:
MsgBox "" & err.DESCRIPTION
  
End Sub
Sub calcAdjNet(k1 As Integer)

vs.TextMatrix(k1, 12) = Round(Val(vs.TextMatrix(k1, 8)) - (Val(vs.TextMatrix(k1, 8)) * Val(vs.TextMatrix(k1, 11)) / 100), 0)
vs.TextMatrix(k1, 13) = (Val(vs.TextMatrix(k1, 10)) - Val(vs.TextMatrix(k1, 12)))


'If vs.TextMatrix(k1, 0) <> "" Then
'If vs.TextMatrix(k1, 1) = "I" Then
'   txtAdjAmt = Val(txtAdjAmt) + Val(vs.TextMatrix(k1, 12))
'   txtDiffNet = Val(txtDiffNet) + Val(vs.TextMatrix(k1, 13))
'Else
'   txtAdjAmt = Val(txtAdjAmt) - Val(vs.TextMatrix(k1, 12))
'   txtDiffNet = Val(txtDiffNet) - Val(vs.TextMatrix(k1, 13))
'End If
'End If

End Sub
Private Sub vs_SelChange()
 
  If (Check1_manullay.value = 0) Then
     vs.Editable = flexEDNone
  Else
     
  
     vs.Editable = flexEDKbdMouse
     If (Val(vs.TextMatrix(vs.RowSel, 18)) > 0) Then
     
        If (Val(vs.TextMatrix(vs.RowSel, 11)) > Val(vs.TextMatrix(vs.RowSel, 18))) Then
           vs.TextMatrix(vs.RowSel, 11) = vs.TextMatrix(vs.RowSel, 18)
           MsgBox "You can mannually Less Disscount...", vbCritical
           Exit Sub
        End If
     
     End If
     
     
  End If
  
  
End Sub
