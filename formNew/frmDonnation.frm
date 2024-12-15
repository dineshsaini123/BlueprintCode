VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDonnation 
   Caption         =   "Extra Discount Calculator"
   ClientHeight    =   10140
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   14736
   Icon            =   "frmDonnation.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10140
   ScaleWidth      =   14736
   Begin VB.TextBox cboSponse1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   10620
      MaxLength       =   4
      TabIndex        =   87
      Top             =   6876
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.CheckBox Check1_gpSchool 
      BackColor       =   &H8000000E&
      Caption         =   "Select Group School"
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
      Left            =   7884
      TabIndex        =   86
      Top             =   324
      Width           =   2328
   End
   Begin VB.CommandButton cmdfatchExDis 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fetch Extra Dis. Ser. Wise (%)"
      Height          =   336
      Left            =   7848
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   8496
      Width           =   2856
   End
   Begin VB.CheckBox Check1_manullay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Manually Change Ext. Dis. && Gross/Net :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7308
      TabIndex        =   84
      Top             =   8856
      Width           =   4392
   End
   Begin VB.TextBox txtManually 
      Height          =   312
      Left            =   11700
      MaxLength       =   250
      TabIndex        =   83
      Top             =   8856
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.TextBox txtPercentSp1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   13212
      MaxLength       =   4
      TabIndex        =   82
      Top             =   8172
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.TextBox txtwave 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9540
      MaxLength       =   7
      TabIndex        =   17
      Top             =   8136
      Width           =   1020
   End
   Begin VB.CheckBox Check1_EditApproval 
      BackColor       =   &H8000000E&
      Caption         =   "Edit Approval"
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
      Height          =   390
      Left            =   13050
      TabIndex        =   78
      Top             =   1665
      Width           =   1470
   End
   Begin VB.CheckBox Check1_Addapp 
      BackColor       =   &H8000000E&
      Caption         =   "By Approval"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   11655
      TabIndex        =   72
      Top             =   1665
      Width           =   1425
   End
   Begin VB.CommandButton cmdPartyRem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View For Party Remarks"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7848
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   972
      Width           =   1785
   End
   Begin VB.ComboBox cboAppSer 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmDonnation.frx":000C
      Left            =   14130
      List            =   "frmDonnation.frx":000E
      TabIndex        =   69
      Top             =   1035
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CheckBox Check1_incAdj 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Include Adjustment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5670
      TabIndex        =   68
      Top             =   45
      Value           =   1  'Checked
      Width           =   2100
   End
   Begin VB.CheckBox Check1_Reupdate 
      BackColor       =   &H8000000E&
      Caption         =   "ReUpdate This Sp. No"
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
      Left            =   11655
      TabIndex        =   67
      Top             =   1080
      Width           =   2430
   End
   Begin VB.CheckBox Check1_untracableSc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Untraceable School"
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
      Left            =   12036
      TabIndex        =   66
      Top             =   30
      Width           =   2364
   End
   Begin VB.CheckBox Check1_manual 
      BackColor       =   &H8000000E&
      Caption         =   "Extra Dis No Enter Manually"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3900
      TabIndex        =   65
      Top             =   0
      Width           =   1680
   End
   Begin VB.CommandButton Command1 
      Height          =   135
      Left            =   90
      TabIndex        =   64
      Top             =   8505
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ComboBox cmbAgentName1 
      Appearance      =   0  'Flat
      Height          =   288
      ItemData        =   "frmDonnation.frx":0010
      Left            =   12108
      List            =   "frmDonnation.frx":0012
      TabIndex        =   10
      Top             =   690
      Width           =   2340
   End
   Begin VB.TextBox txtWhomToBeGivenMob 
      Height          =   315
      Left            =   7860
      MaxLength       =   25
      TabIndex        =   5
      Top             =   1368
      Width           =   3420
   End
   Begin VB.ListBox List1_Sp 
      Appearance      =   0  'Flat
      Height          =   792
      Left            =   7296
      TabIndex        =   60
      Top             =   9228
      Visible         =   0   'False
      Width           =   7368
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FFFFFF&
      Caption         =   "List Of School (Extra Dis Is not created But Extra Dis Amt. Already Given)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   8475
      Width           =   2730
   End
   Begin VB.CheckBox Check1_schoolAll 
      BackColor       =   &H8000000E&
      Caption         =   "Select All School"
      Height          =   255
      Left            =   10230
      TabIndex        =   57
      Top             =   30
      Width           =   1635
   End
   Begin VB.ListBox List1_sc 
      Appearance      =   0  'Flat
      Height          =   984
      Left            =   165
      TabIndex        =   56
      Top             =   9015
      Width           =   7092
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6345
      TabIndex        =   54
      Top             =   6600
      Width           =   1035
   End
   Begin VB.TextBox txtRoundOf 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   12195
      MaxLength       =   10
      TabIndex        =   18
      Top             =   7860
      Width           =   990
   End
   Begin VB.TextBox txtNetBal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   12195
      MaxLength       =   10
      TabIndex        =   52
      Top             =   8172
      Width           =   990
   End
   Begin VB.CheckBox Check1_gp 
      BackColor       =   &H8000000E&
      Caption         =   "Group of School"
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
      Left            =   7890
      TabIndex        =   51
      Top             =   30
      Width           =   2325
   End
   Begin VB.TextBox txtAdvAmt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9540
      MaxLength       =   7
      TabIndex        =   16
      Top             =   7812
      Width           =   1020
   End
   Begin Crystal.CrystalReport cr 
      Left            =   14040
      Top             =   6975
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtSponsorshipNo 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   0
      Width           =   780
   End
   Begin VB.TextBox txtRemarks 
      Height          =   315
      Left            =   135
      MaxLength       =   100
      TabIndex        =   12
      Top             =   7065
      Width           =   7200
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   240
      ScaleHeight     =   876
      ScaleWidth      =   7356
      TabIndex        =   38
      Top             =   7500
      Width           =   7350
      Begin VB.CommandButton cmdPrint1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print Option"
         Height          =   720
         Left            =   5256
         Picture         =   "frmDonnation.frx":0014
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   60
         Width           =   960
      End
      Begin VB.CommandButton cmdPrint_7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Print"
         Height          =   720
         Left            =   4185
         Picture         =   "frmDonnation.frx":0BF8
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   60
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
         Height          =   720
         Left            =   2070
         Picture         =   "frmDonnation.frx":17DC
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   60
         Width           =   945
      End
      Begin VB.CommandButton close 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   720
         Left            =   6255
         Picture         =   "frmDonnation.frx":1BE9
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   60
         Width           =   975
      End
      Begin VB.CommandButton REPORTCD 
         Caption         =   "&Print"
         Height          =   585
         Left            =   4860
         TabIndex        =   42
         Top             =   1020
         Width           =   1245
      End
      Begin VB.CommandButton Abandon 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   720
         Left            =   0
         Picture         =   "frmDonnation.frx":27CD
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   60
         Width           =   945
      End
      Begin VB.CommandButton Del 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   720
         Left            =   3060
         Picture         =   "frmDonnation.frx":33B1
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   60
         Width           =   1050
      End
      Begin VB.CommandButton save 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   720
         Left            =   1005
         Picture         =   "frmDonnation.frx":3F95
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   60
         Width           =   1035
      End
   End
   Begin VB.TextBox txtAmtAfterAdj 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   12195
      TabIndex        =   36
      Top             =   7164
      Width           =   990
   End
   Begin VB.TextBox txtPercentSp 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9540
      MaxLength       =   4
      TabIndex        =   15
      Top             =   7488
      Width           =   1020
   End
   Begin VB.TextBox txtFAmt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFC0&
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
      Left            =   12195
      TabIndex        =   34
      Top             =   7476
      Width           =   990
   End
   Begin VB.TextBox txtNetTotal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   12195
      TabIndex        =   32
      Top             =   6564
      Width           =   990
   End
   Begin VB.TextBox txtGTotal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9555
      TabIndex        =   30
      Top             =   6564
      Width           =   1020
   End
   Begin VB.TextBox txtReturnAdj 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9540
      MaxLength       =   4
      TabIndex        =   14
      Top             =   7164
      Width           =   1020
   End
   Begin VB.ComboBox cboSponse 
      Height          =   288
      ItemData        =   "frmDonnation.frx":4B79
      Left            =   9555
      List            =   "frmDonnation.frx":4B83
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   6876
      Width           =   1020
   End
   Begin VB.TextBox txtPrincipal 
      Height          =   315
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1656
      Width           =   5160
   End
   Begin VB.TextBox txtMob 
      Height          =   315
      Left            =   7860
      MaxLength       =   25
      TabIndex        =   8
      Top             =   1692
      Width           =   3420
   End
   Begin VB.TextBox txtSchoolName 
      Height          =   288
      Left            =   1680
      TabIndex        =   2
      Top             =   612
      Width           =   9012
   End
   Begin VB.TextBox txtWhomTobeGiven 
      Height          =   315
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1296
      Width           =   5145
   End
   Begin VB.ComboBox cboPayment 
      Appearance      =   0  'Flat
      Height          =   288
      ItemData        =   "frmDonnation.frx":4B93
      Left            =   1680
      List            =   "frmDonnation.frx":4BA0
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   972
      Width           =   2220
   End
   Begin VB.ComboBox cboSession 
      Appearance      =   0  'Flat
      Height          =   288
      ItemData        =   "frmDonnation.frx":4BC3
      Left            =   1320
      List            =   "frmDonnation.frx":4BC5
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   612
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.ComboBox cmbAgentName 
      Appearance      =   0  'Flat
      Height          =   288
      ItemData        =   "frmDonnation.frx":4BC7
      Left            =   12108
      List            =   "frmDonnation.frx":4BC9
      TabIndex        =   6
      Top             =   330
      Width           =   2340
   End
   Begin VB.TextBox txtscid 
      Enabled         =   0   'False
      Height          =   315
      Left            =   10680
      TabIndex        =   19
      Top             =   576
      Width           =   756
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   4440
      Left            =   180
      TabIndex        =   9
      Top             =   2115
      Width           =   14385
      _cx             =   25374
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmDonnation.frx":4BCB
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
      Begin VB.Frame Frame1_app 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Approval For This School :-"
         Height          =   3030
         Left            =   1485
         TabIndex        =   73
         Top             =   0
         Visible         =   0   'False
         Width           =   12750
         Begin VB.CommandButton cmdDel 
            BackColor       =   &H00FFFFFF&
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   12105
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   135
            Width           =   420
         End
         Begin VB.CheckBox Check1_EditApp 
            BackColor       =   &H8000000E&
            Caption         =   "Edit Approval For This Extra Dis"
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
            Left            =   2340
            TabIndex        =   75
            Top             =   135
            Width           =   3465
         End
         Begin VSFlex7Ctl.VSFlexGrid vs2 
            Height          =   2490
            Left            =   0
            TabIndex        =   74
            Top             =   495
            Width           =   12705
            _cx             =   22410
            _cy             =   4392
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
         Begin VB.Label Label24 
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
            Left            =   8730
            TabIndex        =   77
            Top             =   180
            Width           =   2955
         End
      End
   End
   Begin MSComCtl2.DTPicker txtDates 
      Height          =   330
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   1350
      _ExtentX        =   2392
      _ExtentY        =   593
      _Version        =   393216
      Format          =   516947969
      CurrentDate     =   39795
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Wave Off :"
      Height          =   288
      Left            =   7920
      TabIndex        =   81
      Top             =   8136
      Width           =   1812
   End
   Begin VB.Label lblCreatedBy 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5856
      TabIndex        =   80
      Top             =   8688
      Width           =   1380
   End
   Begin VB.Label lblCreatedBy1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Creaded By :"
      Height          =   240
      Left            =   4815
      TabIndex        =   79
      Top             =   8685
      Width           =   1005
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   690
      Left            =   11565
      Top             =   1395
      Width           =   3030
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "View Approval For This School:"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   11655
      TabIndex        =   70
      Top             =   1410
      Width           =   2445
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rep.-2:"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   11508
      TabIndex        =   62
      Top             =   756
      Width           =   516
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No :"
      Height          =   288
      Left            =   6888
      TabIndex        =   61
      Top             =   1368
      Width           =   1092
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "School Donnation List :-"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   300
      TabIndex        =   58
      Top             =   8550
      Width           =   2115
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty:"
      Height          =   285
      Left            =   5565
      TabIndex        =   55
      Top             =   6600
      Width           =   915
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Net amount to given :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   10632
      TabIndex        =   53
      Top             =   8172
      Width           =   1392
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
      Left            =   180
      TabIndex        =   50
      Top             =   6600
      Width           =   2955
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Round Of Amt. :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   10632
      TabIndex        =   49
      Top             =   7860
      Width           =   1392
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Adv. Amt :"
      Height          =   288
      Left            =   7920
      TabIndex        =   48
      Top             =   7848
      Width           =   1812
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Discount No :"
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   1
      Left            =   120
      TabIndex        =   45
      Top             =   60
      Width           =   1344
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks :"
      Height          =   285
      Left            =   180
      TabIndex        =   44
      Top             =   6825
      Width           =   1515
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   150
      Top             =   7455
      Width           =   7470
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Amt. After Return Adj. :"
      Height          =   288
      Left            =   10632
      TabIndex        =   37
      Top             =   7224
      Width           =   1632
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Dis (%) :"
      Height          =   288
      Left            =   7920
      TabIndex        =   35
      Top             =   7524
      Width           =   1812
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Final Amt. :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   10632
      TabIndex        =   33
      Top             =   7524
      Width           =   1392
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Total :"
      Height          =   285
      Left            =   11370
      TabIndex        =   31
      Top             =   6600
      Width           =   915
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Total :"
      Height          =   288
      Left            =   7932
      TabIndex        =   29
      Top             =   6576
      Width           =   1032
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Return Adjustment(%) :"
      Height          =   288
      Left            =   7920
      TabIndex        =   28
      Top             =   7224
      Width           =   1728
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Extra Dis On:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   14
      Left            =   7920
      TabIndex        =   27
      Top             =   6876
      Width           =   1320
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Principal Name :"
      Height          =   288
      Left            =   180
      TabIndex        =   26
      Top             =   1716
      Width           =   1512
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No :"
      Height          =   288
      Left            =   6888
      TabIndex        =   25
      Top             =   1692
      Width           =   1092
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Whom to be given :"
      Height          =   288
      Left            =   180
      TabIndex        =   24
      Top             =   1356
      Width           =   1512
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Mode :"
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   0
      Left            =   180
      TabIndex        =   23
      Top             =   1032
      Width           =   1152
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rep.-1:"
      ForeColor       =   &H80000008&
      Height          =   192
      Left            =   11508
      TabIndex        =   22
      Top             =   396
      Width           =   516
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 For Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1680
      TabIndex        =   21
      Top             =   300
      Width           =   2715
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "School Name :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   180
      TabIndex        =   20
      Top             =   648
      Width           =   1512
   End
End
Attribute VB_Name = "frmDonnation"
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
Dim db_ As String
Dim serName_ As String
Dim kk1 As Integer
Dim rss_ As New ADODB.Recordset
Dim rss_10 As New ADODB.Recordset
Dim App_serName, App_Party As String
Dim app_adjPer, app_sponsPer As Double

Dim dt_strR As String
Dim dt_strSaleNext As String
Dim dt_strSaleRNext As String
Dim CON_next As ADODB.Connection
Function fatchDate(fyear_ As String, type_ As String, inv As Long, rows_) As String

   '''================================================================
   '''Checking
   '''================================================================
    
   Select Case fyear_
       
   Case session
    
    If type_ = "I" Then
        
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
    
    If db_ <> "no" Then
    
     If type_ = "I" Then
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
     
     
    End If
    
    
    End Select


End Function
Sub searchData()

On Error GoTo search_

  bb1 = False
  vs.Clear
  
  Dim rss As ADODB.Recordset
  Set rss = New ADODB.Recordset
  
  If Not IsNumeric(txtSponsorshipNo) Then
     txtSponsorshipNo = ""
     Exit Sub
  End If

  Check1_gpSchool.value = 0
  
  If RS.State = 1 Then RS.close
  RS.Open "select * from DonnationMain where DNo=" & Val(txtSponsorshipNo) & "", con, adOpenDynamic, adLockOptimistic
  
  If RS.EOF = False Then
  
     lblCreatedBy.Caption = RS!createdby & ""
     If Not IsNull(RS!GpSchool) Then
        If RS!GpSchool = "y" Then
           Check1_gpSchool.value = 1
        End If
     End If
     
  
     Check1_manual.value = 0
     'lblCreatedBy.Caption = ""
     If rss.State = 1 Then rss.close
     rss.Open "select UserName,[vtype],[desc_] from logtbl where (No=" & Val(txtSponsorshipNo) & " and vtype='donnation')", con
     While rss.EOF = False
        s = Len(Trim(rss!desc_))
        If (rss!vtype = "donnation" And Len(Trim(rss!desc_)) = 6) Then
           ''lblCreatedBy.Caption = rss!UserName
           GoTo aaa:
        End If
        rss.MoveNext
     Wend
     
     
     
aaa:
  
     bb1 = True
     vs.Enabled = False
     
     txtManually.text = RS!manullay_Change & ""
     
     If (txtManually.text <> "") Then
        Check1_manullay.value = 1
     Else
        Check1_manullay.value = 0
     End If

     txtwave.text = RS!waveoff & ""
     
     txtDates.value = RS!DDate
     txtScId.text = RS!scid
     txtSchoolName.text = RS!scname
     cmbAgentName.text = RS!RepName
     
     If RS!RepName1 <> "" Then
        cmbAgentName1.text = RS!RepName1 & ""
     Else
        cmbAgentName1.ListIndex = -1
     End If
     
     cboPayment.text = RS!PaymentMode
     txtWhomTobeGiven.text = RS!WhomTobegiven & ""
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
     txtPercentSp.text = RS!Sponsorship_per
     txtFAmt = RS!finalAmt
     
     txtAdvAmt.text = RS!AdvAmt & ""
     txtRoundOf = RS!RoundOfAAmt & ""
     
     If Val(txtwave.text) > 0 Then
     '   txtRoundOf = (RS!waveoff + RS!RoundOfAAmt)
     End If
     
     txtNetBal = RS!NetBalance & ""
     
     txtWhomToBeGivenMob = RS!MobileWhomtoGiven & ""
     txtRoundOf = RS!RoundOfAAmt & ""

     save.Enabled = False
     Del.Enabled = False
     cmdEdit_4.Enabled = True
     
   End If
   

txtGTotal = 0
txtNetTotal = 0
txtQty = 0

vs.Cols = 14



If rs1.State = 1 Then rs1.close
rs1.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QTY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,addnew,ScId from DonnationMainDet where dno='" & txtSponsorshipNo & "' order by id", con
If rs1.EOF = False Then
vs.rows = rs1.RecordCount + 100
End If

For I = 1 To rs1.RecordCount

DoEvents
DoEvents
DoEvents


     If rs1!AddNew = "Untracable" Then
        Check1_untracableSc.value = 1
     Else
        Check1_untracableSc.value = 0
     End If
     
    vs.TextMatrix(I, 0) = rs1!fyear
    vs.TextMatrix(I, 1) = rs1!Godown
    vs.TextMatrix(I, 2) = rs1!invoiceNo & ""
    If IsDate(rs1!invoiceDate) Then
    vs.TextMatrix(I, 3) = rs1!invoiceDate
    End If
    vs.TextMatrix(I, 4) = rs1!Bookcode
    vs.TextMatrix(I, 5) = rs1!Bookname
    vs.TextMatrix(I, 6) = rs1!qty
    vs.TextMatrix(I, 7) = rs1!rate
    vs.TextMatrix(I, 8) = rs1!GrossAmt
    vs.TextMatrix(I, 9) = rs1!discount
    vs.TextMatrix(I, 10) = Round(rs1!net, 0)
    If rs1!Godown = "I" Then
        txtGTotal = Val(txtGTotal) + rs1!GrossAmt
        txtNetTotal = Val(txtNetTotal) + Round(rs1!net, 0)
    Else
        txtGTotal = Val(txtGTotal) - rs1!GrossAmt
        txtNetTotal = Val(txtNetTotal) - Round(rs1!net, 0)
    End If
    vs.TextMatrix(I, 11) = rs1!AddNew & ""
    vs.TextMatrix(I, 12) = rs1!scname
    vs.TextMatrix(I, 13) = rs1!scid
    
    txtQty = Val(txtQty) + rs1!qty
    

    rs1.MoveNext
Next

txtGTotal = Round(txtGTotal, 0)




vs.FormatString = "Session|Inv.Type|Inv. No.|Inv. Date|B.Code|B.Name|>Quantity|>Rate|>GrossAmt|>Dis.%|>NetAmt|AddNew|School Name|SCId"

vs.ColWidth(1) = 700
vs.ColWidth(2) = 700
vs.ColWidth(3) = 900
vs.ColWidth(4) = 700
vs.ColWidth(5) = 2800
vs.ColWidth(6) = 800
vs.ColWidth(7) = 700
vs.ColWidth(8) = 950

vs.ColWidth(9) = 600
vs.ColWidth(10) = 750

vs.ColWidth(11) = 750
vs.ColWidth(12) = 2050
vs.ColWidth(13) = 750

Exit Sub
search_:
Screen.MousePointer = vbDefault
MsgBox "" & err.DESCRIPTION

     
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

lastData_ = False
txtSponsorshipNo.text = inv

str11 = "SELECT distinct DonnationMain.DNo,DonnationMain.DDate,DonnationMainDet.INVOICENO,DonnationMainDet.Godown FROM  DonnationMain INNER JOIN" & _
      " DonnationMainDet ON DonnationMain.DNo = DonnationMainDet.DNo where DonnationMainDet.fyear='" & session & "' and DonnationMainDet.INVOICENO=" & inv_ledger & ""
ss1_ = ""

str11 = " SELECT  ENTRYNO AS DNo,DATES AS DDate,YRS FROM tmpDonnationnew where  substring(PARTY,1,5)='" & Mid(pname_, 1, 5) & "' and BILLNO=" & inv_ledger & " AND YRS IS NOT NULL"


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

   PopUpValue3 = ""
End If



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







  
  '======================================================================================
  
  bb1 = False
  vs.Clear

  If RS.State = 1 Then RS.close
  If lastData_ = True Then
     RS.Open "select * from DonnationMain where DNo=" & txtSponsorshipNo & "", conadj, adOpenDynamic, adLockOptimistic
  Else
     RS.Open "select * from DonnationMain where DNo=" & txtSponsorshipNo & "", con, adOpenDynamic, adLockOptimistic
  End If
  If RS.EOF = False Then
     bb1 = True
     vs.Enabled = False
     txtDates.value = RS!DDate
     txtScId.text = RS!scid
     txtSchoolName.text = RS!scname
     cmbAgentName.text = RS!RepName
     
     If RS!RepName1 <> "" Then
        cmbAgentName1.text = RS!RepName1 & ""
     Else
        cmbAgentName1.ListIndex = -1
     End If
     
     cboPayment.text = RS!PaymentMode
     txtWhomTobeGiven.text = RS!WhomTobegiven & ""
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
     txtPercentSp.text = RS!Sponsorship_per
     txtFAmt = RS!finalAmt
     
     txtAdvAmt.text = RS!AdvAmt & ""
     txtRoundOf = RS!RoundOfAAmt & ""
     txtNetBal = RS!NetBalance & ""
     
     txtWhomToBeGivenMob = RS!MobileWhomtoGiven & ""
    
     save.Enabled = False
     Del.Enabled = False
     cmdEdit_4.Enabled = True
     
   End If
   

txtGTotal = 0
txtNetTotal = 0
txtQty = 0

vs.Cols = 14

If RS.State = 1 Then RS.close
If lastData_ = True Then
   RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QTY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,addnew,ScId from DonnationMainDet where dno='" & txtSponsorshipNo & "' order by id", conadj
Else
   RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QTY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,addnew,ScId from DonnationMainDet where dno='" & txtSponsorshipNo & "' order by id", con
End If

If RS.EOF = False Then
vs.rows = RS.RecordCount + 100
End If

For I = 1 To RS.RecordCount

     If RS!AddNew = "Untracable" Then
        Check1_untracableSc.value = 1
     Else
        Check1_untracableSc.value = 0
     End If
     
    vs.TextMatrix(I, 0) = RS!fyear
    vs.TextMatrix(I, 1) = RS!Godown
    vs.TextMatrix(I, 2) = RS!invoiceNo & ""
    If IsDate(RS!invoiceDate) Then
    vs.TextMatrix(I, 3) = RS!invoiceDate
    End If
    vs.TextMatrix(I, 4) = RS!Bookcode
    vs.TextMatrix(I, 5) = RS!Bookname
    vs.TextMatrix(I, 6) = RS!qty
    vs.TextMatrix(I, 7) = RS!rate
    vs.TextMatrix(I, 8) = RS!GrossAmt
    vs.TextMatrix(I, 9) = RS!discount
    vs.TextMatrix(I, 10) = Round(RS!net, 0)
    If RS!Godown = "I" Then
        txtGTotal = Val(txtGTotal) + RS!GrossAmt
        txtNetTotal = Val(txtNetTotal) + Round(RS!net, 0)
    Else
        txtGTotal = Val(txtGTotal) - RS!GrossAmt
        txtNetTotal = Val(txtNetTotal) - Round(RS!net, 0)
    End If
    vs.TextMatrix(I, 11) = RS!AddNew & ""
    vs.TextMatrix(I, 12) = RS!scname
    vs.TextMatrix(I, 13) = RS!scid
    
    txtQty = Val(txtQty) + RS!qty
    

    RS.MoveNext
Next

txtGTotal = Round(txtGTotal, 0)




vs.FormatString = "Session|Inv.Type|Inv. No.|Inv. Date|B.Code|B.Name|>Quantity|>Rate|>GrossAmt|>Dis.%|>NetAmt|AddNew|School Name|SCId"

vs.ColWidth(1) = 700
vs.ColWidth(2) = 700
vs.ColWidth(3) = 900
vs.ColWidth(4) = 700
vs.ColWidth(5) = 2800
vs.ColWidth(6) = 800
vs.ColWidth(7) = 700
vs.ColWidth(8) = 950

vs.ColWidth(9) = 600
vs.ColWidth(10) = 750

vs.ColWidth(11) = 750
vs.ColWidth(12) = 2050
vs.ColWidth(13) = 750

     
End Sub

Private Sub ABANDON_Click()

Check1_gpSchool.value = 0
refresh_
max_sp

con.Execute "delete from tmpDDet where userid=" & UId & ""
con.Execute "delete from tmpDonnation where username='" & UserName & "'"

If checkPermission("donnation") = True Then
   con.Execute "exec tmpdata " & UId & ""
End If

cmdEdit_4.Enabled = False
Del.Enabled = False

List1_sc.Clear
List1_Sp.Visible = False


End Sub
Private Sub cboAppSer_Click()
  Screen.MousePointer = vbHourglass
   serName_ = cboAppSer.text
   PopUpValue2 = txtScId
   PopUpValue1 = txtSchoolName
   filterData
   calAmt
  Screen.MousePointer = vbDefault
End Sub

Private Sub cboPayment_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtWhomTobeGiven.SetFocus
End Sub

Private Sub cboSession_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cboPayment.SetFocus
End Sub



Private Sub cboSponse_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtReturnAdj.SetFocus
End Sub

Private Sub cboSponse_LostFocus()

If Check1_manullay.value = 0 Then
   
If Len(cboSponse1.text) > 0 Then

   If (cboSponse.text <> cboSponse1.text) Then
      MsgBox "Plz Enter Remark for this change  ...", vbCritical
      Check1_manullay.value = 1
   End If
   
End If
   
End If
End Sub

Private Sub Check1_Addapp_Click()
If Check1_Addapp.value = 1 Then
   Check1_EditApproval.Enabled = True
Else
   Check1_EditApproval.Enabled = False
End If
End Sub

Private Sub Check1_EditApproval_Click()
If Check1_EditApproval.value = 1 Then
   
   If txtScId.text <> "" Then
    Frame1_app.Visible = True
    
    Check1_EditApp.value = 1
    Check1_EditApp.Enabled = False
    'AddAproval
   End If
   
Else
   Frame1_app.Visible = False
End If

End Sub

Private Sub Check1_manullay_Click()
If Check1_manullay.value = 1 Then
  txtManually.Visible = True
  txtPercentSp.Enabled = True
Else
   txtManually.Visible = False
End If
End Sub

Private Sub Check1_untracableSc_Click()

If Check1_untracableSc.value = 1 Then
   Check1_schoolAll.value = 1
Else
   Check1_schoolAll.value = 0
End If

End Sub
Private Sub close_Click()
Unload Me
End Sub

Private Sub cmbAgentName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cboPayment.SetFocus
End If
End Sub
Sub max_sp()
If RS.State = 1 Then RS.close
RS.Open "select max(DNo) from DonnationMain", con
If Not IsNull(RS(0)) Then
   txtSponsorshipNo = RS(0) + 1
Else
   txtSponsorshipNo = 1
End If

End Sub
Private Sub cmdDel_Click()
DoEvents

txtSchoolName.SetFocus
fillSeries

'PopUpValue1 = ""
'PopUpValue2 = ""
'PopUpValue3 = ""

Frame1_app.Visible = False
DoEvents
DoEvents
End Sub
Sub fillSeries()

Dim ff As ADODB.Recordset
serName_ = ""
'substring([SUBLEDGER],1,5)
Set ff = New ADODB.Recordset
ff.Open "select pcode,adj,promo from tmpAppForDonnation where uid=" & UId & " group by pcode,adj,promo", con
If ff.EOF = False Then
    txtReturnAdj = ff!adj
    txtPercentSp = ff!Promo
End If

While ff.EOF = False
If serName_ = "" Then
   serName_ = " substring([SUBLEDGER],1,5) ='" & ff(0) & "'"
Else
   serName_ = serName_ & " and substring([SUBLEDGER],1,5) ='" & ff(0) & "'"
End If
ff.MoveNext
Wend
    
ss_ = ""
Set ff = New ADODB.Recordset
ff.Open "select serName from tmpAppForDonnation where uid=" & UId & " group by serName", con
While ff.EOF = False
If ss_ = "" Then
   ss_ = " serName ='" & ff(0) & "'"
Else
   ss_ = ss_ & " or serName ='" & ff(0) & "'"
End If
ff.MoveNext
Wend

If ss_ <> "" Then
   ss_ = "(" & ss_ & ")"
End If

If ss_ <> "" Then
serName_ = serName_ & " and " & ss_
End If

    
    
End Sub
Private Sub cmdEdit_4_Click()


   

 If rs1.State = 1 Then rs1.close
 rs1.Open "select top 100 * from DonnationMain where DNo=" & txtSponsorshipNo.text & "", con, adOpenKeyset, adLockReadOnly
 If rs1.EOF = False Then
    If RS.State = 1 Then RS.close
    RS.Open "select top 100 * from DonnationMain where DNo=" & txtSponsorshipNo.text & "", con, adOpenKeyset, adLockReadOnly
    If rs1!bAuthorized = True Then
       MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
       Exit Sub
   End If
End If




vs.Enabled = True
con.Execute "delete from tmpDonnation where Sno=" & txtSponsorshipNo & " and UserName='" & UserName & "'"
con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,sNo) select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,DNo from DonnationMainDet where dno=" & txtSponsorshipNo & ""
If Check1_Reupdate.value = 1 Then
   fatchData
   fillGrid_
   calAmt
End If

Del.Enabled = True
save.Enabled = True
Edit = True
cmdEdit_4.Enabled = False
save.SetFocus

End Sub

Private Sub cmdfatchExDis_Click()

On Error GoTo err:

Dim bcode As String
Dim sername As String
Dim rs10 As New ADODB.Recordset
Dim code_ As String
code_ = ""
inv_ = ""
vtype_ = ""

Set rs10 = New ADODB.Recordset

If (vs.TextMatrix(vs.RowSel, 4) <> "") Then
   bcode = vs.TextMatrix(vs.RowSel, 4)
   inv_ = vs.TextMatrix(vs.RowSel, 2)
   vtype_ = vs.TextMatrix(vs.RowSel, 1)
Else
   bcode = vs.TextMatrix(1, 4)
   inv_ = vs.TextMatrix(1, 2)
   vtype_ = vs.TextMatrix(1, 1)
End If

If rs10.State = 1 Then rs10.close
rs10.Open "SELECT  distinct SerName from books where BOOKCODE='" & bcode & "'", con
If rs10.EOF = False Then
   sername = rs10(0)
End If


If (vtype_ = "I") Then

    If rs10.State = 1 Then rs10.close
    rs10.Open "SELECT  subledger from invoicea where invoiceno=" & inv_ & "", con
    If rs10.EOF = False Then
       code_ = Trim(Mid(rs10!subledger, 1, 6))
    End If

Else

    If rs10.State = 1 Then rs10.close
    rs10.Open "SELECT  subledger from CREDITA where invoiceno=" & inv_ & "", con
    If rs10.EOF = False Then
       code_ = Trim(Mid(rs10!subledger, 1, 6))
    End If

End If


Dim hh1 As Integer

hh1 = 0

If (sername <> "") Then

Set rs10 = con.Execute("exec SearchAppData_SerNameExDis_New '" & txtScId.text & "','" & sername & "','" & code_ & "'")
If rs10.EOF = False Then
   cboSponse.text = rs10!Net_Gross
   cboSponse1.text = rs10!Net_Gross
   
   txtPercentSp.text = rs10!Promo
   txtPercentSp1.text = rs10!Promo
   hh1 = 1
End If

If (hh1 = 0) Then
Set rs10 = con.Execute("exec SearchAppData_SerNameExDis_New '" & code_ & "','" & sername & "','" & txtScId.text & "'")
If rs10.EOF = False Then
   cboSponse.text = rs10!Net_Gross
   cboSponse1.text = rs10!Net_Gross
   
   txtPercentSp.text = rs10!Promo
   txtPercentSp1.text = rs10!Promo
   
End If
End If

End If


Exit Sub

err:
MsgBox err.DESCRIPTION

End Sub

Private Sub cmdPartyRem_Click()

st10 = ""
str1 = "SELECT SUBLEDGER,PartyRemarks,AppNo FROM invoiceaQry where scid='" & txtScId.text & "' group by SUBLEDGER,PartyRemarks,AppNo"
If RS.State = 1 Then RS.close
RS.Open str1, con
While RS.EOF = False
  
  If st10 = "" Then
     st10 = RS(0) & " => " & RS(1) & ":" & RS!appno
  Else
     st10 = st10 & vbCrLf & RS(0) & " => " & RS(1) & ":" & RS!appno
  End If
  
  RS.MoveNext
Wend





If Len(st10) > 0 Then
   MsgBox "" & st10
Else
   MsgBox "No Party Terms Exist....", vbCritical
End If


End Sub
Sub printPRemarks()

Dim f As New ADODB.Recordset
Set f = New ADODB.Recordset
st10 = ""
Dim ss_, ss1_, pp_ As String
ss_ = ""
ss1_ = ""
pp_ = ""
con.Execute "delete from AppPrintTmp"

f.Open "select subledger from DonnationMainDet where DNo=" & txtSponsorshipNo.text & " group by subledger", con
While f.EOF = False
    If pp_ = "" Then
        pp_ = "subledger= '" & f(0) & "'"
    Else
       pp_ = pp_ & " or  subledger= '" & f(0) & "'"
    End If
f.MoveNext
Wend

 pp_ = "(" & pp_ & ")"
 
If Len(pp_) < 17 Then
   Exit Sub
End If


If (txtScId.text <> "" And txtSchoolName.text <> "") Then
   
   str1 = "SELECT SUBLEDGER,PartyRemarks,appNo,Promo,Net_Gross,Adj FROM PartyRemarksQry " & _
   " where (scid='" & txtScId.text & "') group by SUBLEDGER,PartyRemarks,appNo,Promo,Net_Gross,Adj"
End If


ss_ = ""
If f.State = 1 Then f.close
f.Open str1, con

sss = f.RecordCount

While f.EOF = False
       
       ss_ = ""
       remarks1 = ""
       If rs1.State = 1 Then rs1.close
       rs1.Open "select sername,Promo,remarks from AppForm " & _
       "where (appNo=" & f(2) & " and Promo=" & f!Promo & ") group by sername,Promo,remarks", con
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
       'End If
       
'If (f(1) <> "" Or Val(f(3)) > 0 Or Len(f(4)) > 0) Then
     con.Execute "insert into AppPrintTmp(party,Remarks,appno,PromPer,adjper,sername,remarks1,gross_) " & _
     " values('" & f.Fields("SUBLEDGER").value & "','" & f.Fields("PartyRemarks").value & "'," & _
     "'" & f.Fields("appNo").value & "','" & f.Fields("Promo").value & "'," & _
     "'" & f.Fields("Adj").value & "','" & ss_ & "','" & remarks1 & "'," & _
     "'" & f.Fields("Net_Gross").value & "')"
'End If
     ss_ = ""


    f.MoveNext
Wend




If RS.State = 1 Then RS.close
RS.Open "SELECT SUBLEDGER,PartyRemarks,appNo FROM invoiceaQry where scid='" & txtScId.text & "' group by SUBLEDGER,PartyRemarks,appNo", con, adOpenDynamic, adLockOptimistic
While RS.EOF = False
   If rs1.State = 1 Then rs1.close
   rs1.Open "SELECT party,Remarks,appNo FROM AppPrintTmp where party='" & RS!subledger & "'", con, adOpenDynamic, adLockOptimistic
   If rs1.RecordCount > 1 Then
      con.Execute "delete from AppPrintTmp where party='" & rs1!party & "' and AppNo =''"
   End If
   RS.MoveNext
Wend




End Sub
Private Sub cmdPrint_7_Click()

Dim str1 As String
str1 = ""


Dim inv_ As String
inv_ = ""

con.Execute "delete from tmpPartyRemarksQryNew"

For I = 1 To vs.rows - 1

If vs.TextMatrix(I, 1) = "I" Then
   If inv_ = "" Then
      inv_ = vs.TextMatrix(I, 2)
   Else
       inv_ = inv_ & "," & vs.TextMatrix(I, 2)
   End If
End If
Next


If rs1.State = 1 Then rs1.close
rs1.Open "SELECT ScName FROM DonnationMainDet where dno=" & txtSponsorshipNo.text & " group by ScName", con
While rs1.EOF = False

con.Execute "insert into tmpPartyRemarksQryNew " & _
"select * from PartyRemarksQryNew where scname='" & rs1!scname & "'"

rs1.MoveNext
Wend

printPRemarks

If Check1_gpSchool.value = 0 Then
    If RS.State = 1 Then RS.close
    RS.Open "select distinct(ScName) from DonnationMainDet where DNo=" & txtSponsorshipNo.text & "", con, adOpenDynamic, adLockOptimistic
    While RS.EOF = False
       If str1 = "" Then
          str1 = RS(0)
       Else
          str1 = str1 & " :: " & RS(0)
       End If
       RS.MoveNext
    Wend
End If


If str1 <> "" Then
    con.Execute "update DonnationMain set Sh_Name='" & str1 & "' where DNo=" & txtSponsorshipNo.text & ""
Else
   con.Execute "update DonnationMain set Sh_Name=ScName  where DNo=" & txtSponsorshipNo.text & ""
End If



'lblCreatedBy.Caption = ""
If rs1.State = 1 Then rs1.close
rs1.Open "select UserName,[vtype],[desc_] from logtbl where (No=" & Val(txtSponsorshipNo) & " and vtype='donnation')", con
While rs1.EOF = False
   s = Len(Trim(rs1!desc_))
   If (rs1!vtype = "donnation" And Len(Trim(rs1!desc_)) = 6) Then
      'lblCreatedBy.Caption = rs1!UserName
      GoTo aaa1:
   End If
   rs1.MoveNext
Wend

aaa1:

DSNNew

If MsgBox("Want to Print ?", vbQuestion + vbYesNo) = vbYes Then


    CR.Reset
    CR.ReportFileName = rptPath & "/ExtraDisFile.rpt"
    CR.ReplaceSelectionFormula "{donnationmain.dno}=" & txtSponsorshipNo & ""
    CR.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    CR.Formulas(0) = "fyear='" & session & "'"
    CR.Formulas(1) = "createdby='" & lblCreatedBy.Caption & "'"
    CR.WindowShowPrintSetupBtn = True
    CR.WindowShowPrintBtn = True
    CR.WindowShowExportBtn = True
    CR.WindowState = crptMaximized
    CR.Action = 1


End If

End Sub

Private Sub cmdPrint1_Click()
frmDonnationPrint.Show
End Sub

Private Sub cmdView_Click()
List1_Sp.Visible = True
List1_Sp.Clear
str1 = "SELECT DonnationMain.SCId,DonnationMain.SCName,DonnationMain.DNo,DonnationMain.DDate,DonnationMainDet.DNo FROM DonnationMain  LEFT JOIN DonnationMainDet ON DonnationMain.DNo = DonnationMainDet.DNo where DonnationMainDet.DNo is null"
If RS.State = 1 Then RS.close
RS.Open str1, con
While RS.EOF = False
   
   'If rs1.State = 1 Then rs1.close
   'rs1.Open "select ScID,ScName from invoicea where ScID='" & RS(0) & "'", CON
   'If rs1.EOF = False Then
      List1_Sp.AddItem RS(2) & " => " & RS(0) & " :: " & RS(1)
   'End If
   RS.MoveNext
Wend

End Sub

Private Sub Command1_Click()
Dim d_, amt_ret, amt_sp
d_ = 0
amt_ret = 0
amt_sp = 0


If rs1.State = 1 Then rs1.close
rs1.Open "select DNo from donation_Qry group by DNo", con

While rs1.EOF = False

If RS.State = 1 Then RS.close
RS.Open "select Sponsorship_per,ReturnAdj,GrossAmt,Net,SponsorshipOn,Godown,id from donation_Qry where DNo=" & rs1!dno & " order by BOOKCODE", con
While RS.EOF = False
   
  d_ = 0
  amt_ret = 0
  amt_sp = 0
   
   If RS!SponsorshipOn = "Gross" Then
      amt_ret = (RS!GrossAmt * RS!ReturnAdj / 100)
      amt_sp = RS!GrossAmt - amt_ret
      
   Else
      amt_ret = (RS!net * RS!ReturnAdj / 100)
      amt_sp = RS!net - amt_ret
   End If
   
   d_ = amt_sp * RS!Sponsorship_per / 100
   If RS!Godown = "C" Then
      d_ = -1 * d_
   End If
   
   con.Execute "update DonnationMainDet set DonationAmtBk=" & d_ & " where DNo=" & rs1!dno & " and id='" & RS!id & "'"
   RS.MoveNext
Wend


rs1.MoveNext

Wend


End Sub
Private Sub Del_Click()


'If rs1.State = 1 Then rs1.close
'rs1.Open "SELECT [DNo],[SCId],[ScName],[Remarks],[AdvAmt],[NetBalance],[UserName] FROM deleteDonnationMain where [DNo]='" & txtSponsorshipNo & "'", con
'If rs1.EOF = False Then
'   MsgBox "You have already deleted.. !!", vbExclamation, "Alert"
'   Exit Sub
'End If


If MsgBox("want to delete ?", vbQuestion + vbYesNo) = vbYes Then

    If rs1.State = 1 Then rs1.close
    rs1.Open "select top 100 * from DonnationMain where DNo=" & txtSponsorshipNo.text & "", con, adOpenKeyset, adLockReadOnly
    If rs1.EOF = False Then
       If RS.State = 1 Then RS.close
       RS.Open "select top 100 * from DonnationMain where DNo=" & txtSponsorshipNo.text & "", con, adOpenKeyset, adLockReadOnly
       If rs1!bAuthorized = True Then
          MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
          Exit Sub
      End If
   End If
   
   createLog UserName, txtSponsorshipNo, "donnation", " Delete : " & txtNetBal, Date

   If RS.State = 1 Then RS.close
   RS.Open "select top 10 dno,scid,scname,remarks,AdvAmt,NetBalance from DonnationMain where DNo=" & txtSponsorshipNo.text & "", con, adOpenKeyset, adLockReadOnly
   If RS.EOF = False Then
      con.Execute "insert into deleteDonnationMain(DNo,SCId,SCName,AdvAmt,NetBalance,UserName) " & _
      " values(" & RS!dno & ",'" & RS!scid & "','" & RS!scname & "'," & RS!AdvAmt & "," & RS!NetBalance & ",'" & UserName & "')"
     
      If rs1.State = 1 Then rs1.close
      rs1.Open "select top 1 * from deleteDonnationMain where DNo=" & txtSponsorshipNo.text & " and SCId='" & RS!scid & "'", con, adOpenDynamic, adLockOptimistic
      If rs1.EOF = False Then
       rs1!remarks = RS!remarks
       rs1.update
      End If
      
   
   End If
   
   con.Execute "delete from DonnationMain where dno=" & txtSponsorshipNo & ""
   con.Execute "delete from DonnationMainDet where dno=" & txtSponsorshipNo & ""
   con.Execute "delete from tmpDonnation where Sno=" & txtSponsorshipNo & " and UserName='" & UserName & "'"
   
   For I = 1 To vs.rows - 1
    If vs.TextMatrix(I, 1) <> "" Then
       con.Execute "delete FROM tmpDDet where (Fyear='" & vs.TextMatrix(I, 0) & "' and Godown='" & vs.TextMatrix(I, 1) & "'  and invoiceno='" & vs.TextMatrix(I, 2) & "')"
    End If
   Next

  'con.Execute "exec tmpdata " & UId & ""
   
   refresh_
   
End If

End Sub
Sub refresh_()
kk1 = 0

cboSponse1.text = ""

txtPercentSp.Enabled = False
Check1_manullay.value = 0

txtManually.text = ""

txtPercentSp1.text = ""
lblCreatedBy.Caption = ""

Check1_EditApproval.value = 0
Check1_Addapp.value = 0

Check1_untracableSc.value = 0
Check1_Reupdate.value = 0
vs.Enabled = True

app_adjPer = 0
app_sponsPer = 0

PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""

Check1_EditApp.value = 0
Add = True
Edit = False
vs.Clear

txtwave.text = ""
txtNetBal = ""
txtWhomToBeGivenMob = ""
txtSponsorshipNo.text = ""
txtDates.value = Format(Date, "dd/MM/yyyy")
txtScId.text = ""
txtSchoolName.text = ""
cmbAgentName.ListIndex = -1
cmbAgentName1.ListIndex = -1
cboPayment.ListIndex = -1
txtWhomTobeGiven.text = ""
txtPrincipal = ""
txtMob = ""
txtRemarks = ""
txtGTotal = 0
txtNetTotal = 0
cboSponse.ListIndex = -1
txtReturnAdj = 0
txtAmtAfterAdj = 0
txtPercentSp.text = 0
txtFAmt = 0

txtRoundOf = 0
txtAdvAmt = 0

txtSponsorshipNo.SetFocus
save.Enabled = True

Check1_manual.value = 0


vs.FormatString = "Session|Inv.Type|Inv. No.|Inv. Date|B.Code|B.Name|>Quantity|>Rate|>GrossAmt|>Dis.%|>NetAmt|AddNew|School Name|SCId"
vs.ColWidth(1) = 700
vs.ColWidth(2) = 700
vs.ColWidth(3) = 900
vs.ColWidth(4) = 700
vs.ColWidth(5) = 2800
vs.ColWidth(6) = 800
vs.ColWidth(7) = 700
vs.ColWidth(8) = 950

vs.ColWidth(9) = 600
vs.ColWidth(10) = 750

vs.ColWidth(11) = 750
vs.ColWidth(12) = 2050
vs.ColWidth(13) = 750

End Sub
Private Sub Form_Load()
On Error GoTo ee

Screen.MousePointer = vbHourglass

kk1 = 0

con.Execute "exec tmpdata_saleadj"

If RS.State = 1 Then RS.close
RS.Open "select Rep as Representative from SalesRepQry order by Rep", CON_blue
cmbAgentName.Clear
Me.cmbAgentName1.Clear
If Not RS.EOF Then
   Do While Not RS.EOF
      If IsNull(RS(0)) = False Then
        Me.cmbAgentName.AddItem RS(0)
        Me.cmbAgentName1.AddItem RS(0)
      End If
      If Not RS.EOF Then RS.MoveNext
    Loop
End If



cboSponse.ListIndex = 0


If RS.State = 1 Then RS.close
RS.Open "select * from financialyear where fyear='" & session & "'", CCON
If RS.EOF = False Then
   dt_from = RS!fromdate
   dt_to = RS!todate
End If


fdate = ""
If RS.State = 1 Then RS.close
RS.Open "select fromDate,toDate,fromDateSRet,toDateSRet from turnOverDis where Current_Next='current'", CCON
If RS.EOF = False Then
   dt_str = "(INVOICEDATE >= convert(smalldatetime,'" & RS!fromdate & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!todate & "',103))"
   dt_strR = "(INVOICEDATE >= convert(smalldatetime,'" & RS!fromDateSRet & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!toDateSRet & "',103))"
   fdate = RS!fromdate
End If



If RS.State = 1 Then RS.close
RS.Open "select fromDateSRet,toDateSRet,NotCreated from turnOverDis where (NotCreated='y' AND Current_Next='next')", CCON
If RS.EOF = False Then
If RS!NotCreated = "y" Then
   dt_strSaleNext = "(INVOICEDATE >= convert(smalldatetime,'" & RS!fromDateSRet & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!toDateSRet & "',103))"
   dt_strSaleRNext = "(INVOICEDATE >= convert(smalldatetime,'" & RS!fromDateSRet & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!toDateSRet & "',103))"
End If
End If




If Len(inviceNo) > 0 Then
     txtSponsorshipNo = inviceNo
     SearchDataNew inviceNo
     PopUpValue1 = ""
     PopUpValue2 = ""
     popupvalue5 = ""
     inviceNo = ""
     
     Me.top = 0
     Me.Left = 0
     Me.Width = 14800
     Me.Height = 10435
     BackColorFrom Me
     Screen.MousePointer = vbDefault
     Exit Sub
End If



''1Dec-2018

''''If RS.State = 1 Then RS.close
''''RS.Open "select fromDate,toDate from SaleAdj_donnationDateRange", CCON
''''If RS.EOF = False Then
''''   dt_str = "(INVOICEDATE >= convert(smalldatetime,'" & RS!FromDate & "',103) and INVOICEDATE <= convert(smalldatetime,'" & RS!toDate & "',103))"
''''End If




con.Execute "delete from tmpDonnation where username='" & UserName & "'"
txtDates.value = Format(Date, "dd/MM/yyyy")
max_sp


'------------------
'------------------
'------------------

a1 = ""
a2 = ""
db_ = "no"

If RS.State = 1 Then RS.close
RS.Open "select fyear,DataBase,Current_Next,NotCreated from turnOverDis where NotCreated='y' order by fyear", CCON
While RS.EOF = False
If a2 = "" Then
   a2 = RS!fyear
Else
   a2 = a2 & "|" & RS!fyear
End If

If RS!current_next = "next" Then
   If RS!NotCreated = "y" Then
      db_ = RS!Database
    Else
      db_ = "no"
   End If
End If


RS.MoveNext
Wend

vs.ColComboList(0) = a2

''''==================================================================================================
''''==================================================================================================
'------Fatch Data From Next Session---------------------------------------------------------------------------------------------------------------------------------


If db_ <> "no" Then
    Set CON_next = New ADODB.Connection
    If LCase(server_) = "server" Then
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & db_ & "; UID=" & sql_user & "; PWD=" & sql_pass
       CON_next.Open
       PopUpValue6 = ""
    Else
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=HP\SQL2008; DATABASE=" & db_ & "; UID=; PWD=;"
       CON_next.Open
    End If
End If



con.Execute "delete from tempLedger_net where userid=" & UId
con.Execute "insert into tempLedger_net(Billtype,des,userId) SELECT  ScID,ScName,'" & UId & "' FROM INVOICEA where (len(ScID)>0 and len(ScName)>0) group by ScID,ScName"
If rs1.State = 1 Then rs1.close
rs1.Open "select distinct Billtype,des from tempLedger_net where userid=" & UId & "", con, adOpenDynamic, adLockOptimistic

If db_ <> "no" Then

If RS.State = 1 Then RS.close
RS.Open "SELECT ScName,ScID FROM INVOICEA where len(ScName)>0 group by ScName,ScID", CON_next
While RS.EOF = False
   
    rs1.MoveFirst
    rs1.Find "Billtype='" & RS!scid & "'"
    If rs1.EOF = True Then
       con.Execute "insert into tempLedger_net(des,Billtype,userid) values('" & RS(0) & "','" & RS(1) & "'," & UId & ")"
    End If
    RS.MoveNext
Wend

End If







Screen.MousePointer = vbDefault
'------------------
Me.top = 0
Me.Left = 0
Me.Width = 14800
Me.Height = 10500
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

If Edit = False Then
   saveData
Else
   MODIFTYDATA
End If


End Sub
Sub saveData()



Dim dd1, dd2


dd1 = Date
dd2 = txtDates.value
dd3 = DateDiff("D", dd1, dd2)

If dd3 > 0 Then
   MsgBox "Date must by less then or equal to current Date...", vbInformation
   txtDates.SetFocus
   Exit Sub
End If


If (Check1_manullay.value = 1) Then
   If (txtManually.text = "") Then
      
      MsgBox "Enter Remarks...", vbInformation
      txtManually.SetFocus
      Screen.MousePointer = vbDefault
      Exit Sub
      
   End If
End If





If txtScId = "" Then
   MsgBox "Select School Name ...", vbCritical
   txtSchoolName.SetFocus
   Exit Sub
End If

If cmbAgentName.text = "" Then
   MsgBox "Select Representative Name ...", vbCritical
   cmbAgentName.SetFocus
   Exit Sub
End If

If cboPayment.text = "" Then
   MsgBox "Select Payment Mode ...", vbCritical
   cboPayment.SetFocus
   Exit Sub
End If

If txtWhomTobeGiven.text = "" Then
   MsgBox "Enter Whom To be Given ...", vbCritical
   txtWhomTobeGiven.SetFocus
   Exit Sub
End If


If rs1.State = 1 Then rs1.close
rs1.Open "select top 100 * from DonnationMain where DNo=" & txtSponsorshipNo.text & "", con, adOpenKeyset, adLockReadOnly
If rs1.EOF = False Then
   
   If RS.State = 1 Then RS.close
   RS.Open "select top 100 * from DonnationMain where DNo=" & txtSponsorshipNo.text & "", con, adOpenKeyset, adLockReadOnly
       If rs1!bAuthorized = True Then
          MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
          Exit Sub
      End If
   
End If


Dim rSave As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset


createLog UserName, txtSponsorshipNo, "donnation", " Save : " & txtNetBal, Date


If Edit = True Then
    

    
    con.Execute "delete from DonnationMainDet where (DNo=" & txtSponsorshipNo & ")"
    con.Execute "delete from DonnationMain where (DNo=" & txtSponsorshipNo & ")"
    Edit = False
    
    
Else
If Check1_manual.value = 0 Then
    max_sp
End If
End If



If RS.State = 1 Then RS.close
RS.Open "select * from DonnationMain where DNo=" & txtSponsorshipNo & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = True Then
 RS.AddNew
 
 If Check1_untracableSc.value = 0 Then
    con.Execute "insert into DonnationMainDet(Dno,INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,id,repname) select '" & txtSponsorshipNo & "' ,INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,id,repname from tmpDonnation where UserName='" & UserName & "' and addnew is null"
 Else
    
   If rSave.State = 1 Then rSave.close
   rSave.Open "select * from DonnationMainDet", con, adOpenDynamic, adLockOptimistic
    
   For k1 = 1 To vs.rows - 1
   If vs.TextMatrix(k1, 0) <> "" Then
    rSave.AddNew
    
    rSave!dno = txtSponsorshipNo
    If Val(vs.TextMatrix(k1, 2)) = 0 Then
       rSave!invoiceNo = Null
    Else
       rSave!invoiceNo = vs.TextMatrix(k1, 2)
    End If
    
    If Not IsDate(vs.TextMatrix(k1, 3)) Then
       rSave!invoiceDate = Null
    Else
       rSave!invoiceDate = vs.TextMatrix(k1, 3)
    End If
    
    
    rSave!Bookcode = vs.TextMatrix(k1, 4)
    rSave!Bookname = vs.TextMatrix(k1, 5)
    rSave!qty = vs.TextMatrix(k1, 6)
    rSave!rate = vs.TextMatrix(k1, 7)
    rSave!GrossAmt = vs.TextMatrix(k1, 8)
    rSave!discount = vs.TextMatrix(k1, 9)
    rSave!net = vs.TextMatrix(k1, 10)
    If vs.TextMatrix(k1, 1) = "" Then
       rSave!Godown = Null
    Else
       rSave!Godown = vs.TextMatrix(k1, 1)
    End If
    
    If vs.TextMatrix(k1, 0) = "" Then
    rSave!fyear = Null
    Else
    rSave!fyear = vs.TextMatrix(k1, 0)
    End If
    
    rSave!scname = vs.TextMatrix(k1, 12)
    rSave!scid = vs.TextMatrix(k1, 13)
    rSave!UserName = UserName
    rSave!createdby = UserName
    rSave!AddNew = "Untracable"
    rSave.update
    
    
   End If
  Next
 End If
 
End If



RS!waveoff = IIf(txtwave = "", 0, txtwave)
RS!dno = txtSponsorshipNo.text
RS!DDate = txtDates.value
RS!scid = txtScId.text
RS!scname = txtSchoolName.text
RS!RepName = cmbAgentName.text
RS!RepName1 = cmbAgentName1.text
RS!PaymentMode = cboPayment.text
RS!WhomTobegiven = Trim(txtWhomTobeGiven.text)
RS!Principal = Trim(txtPrincipal)
RS!mobile = Trim(txtMob)
RS!remarks = Trim(txtRemarks)
RS!GrossAmt = Val(txtGTotal)
RS!net = Val(txtNetTotal)
RS!SponsorshipOn = cboSponse.text
RS!ReturnAdj = Val(txtReturnAdj)
RS!AmtAfter_ReturnAdj = Val(txtAmtAfterAdj)
RS!Sponsorship_per = Val(txtPercentSp.text)
RS!finalAmt = Val(txtFAmt)
RS!AdvAmt = Val(txtAdvAmt.text)
RS!RoundOfAAmt = Val(txtRoundOf)
RS!NetBalance = Val(txtNetBal)
RS!MobileWhomtoGiven = Trim(txtWhomToBeGivenMob)
RS!UserName = UId
RS!createdby = UserName
RS!manullay_Change = txtManually.text


RS.update


save.Enabled = False
cmdEdit_4.Enabled = True
Add = False
Edit = False

If txtScId <> "" Then
   CON_blue.Execute "update College set Pr_Name='" & Trim(txtPrincipal) & "',Pr_Mobile='" & txtMob.text & "' where CollegeID='" & txtScId & "'"
End If


'new code
Dim d_, amt_ret, amt_sp
d_ = 0
amt_ret = 0
amt_sp = 0
d_1 = 0
'---------------------
If txtSponsorshipNo <> "" Then

If RS.State = 1 Then RS.close
RS.Open "select Sponsorship_per,ReturnAdj,GrossAmt,Net,SponsorshipOn,Godown,id from donation_Qry where DNo=" & txtSponsorshipNo & " order by BOOKCODE", con
While RS.EOF = False
   
  
  d_ = 0
  amt_ret = 0
  amt_sp = 0
   
   If RS!SponsorshipOn = "Gross" Then
      amt_ret = (RS!GrossAmt * RS!ReturnAdj / 100)
      amt_sp = RS!GrossAmt - amt_ret
      
   Else
      amt_ret = (RS!net * RS!ReturnAdj / 100)
      amt_sp = RS!net - amt_ret
   End If
   
   
   d_ = amt_sp * RS!Sponsorship_per / 100
   If RS!Godown = "C" Then
      d_ = -1 * d_
   End If
   
   
   con.Execute "update DonnationMainDet set DonationAmtBk=" & d_ & " where DNo=" & txtSponsorshipNo & " and id='" & RS!id & "'"
   
   RS.MoveNext
Wend



End If



Check1_manual.value = 0


con.Execute "exec tmpdata " & UId & ""
'=========================================================================

Dim st_ As String
If rs1.State = 1 Then rs1.close
rs1.Open "select DNo from DonnationMain where dno='" & txtSponsorshipNo.text & "'", con
If rs1.EOF = False Then
    st_ = ""
    If RS.State = 1 Then RS.close
    RS.Open "select subledger from DonnationQry_ where dno=" & rs1!dno & " and subledger is not null", con, adOpenDynamic, adLockReadOnly
    While RS.EOF = False
         If st_ = "" Then
             st_ = RS!subledger
         Else
             st_ = st_ & ", " & RS!subledger
         End If
         RS.MoveNext
     Wend
     con.Execute "update DonnationMain set party_='" & st_ & "'  where DNo=" & rs1!dno & ""
End If


con.Execute "update deleteDonnationMain set [Status]='Made'  where DNo=" & rs1!dno & ""



End Sub
Sub MODIFTYDATA()

Screen.MousePointer = vbHourglass

On Error GoTo modify_


If (Check1_manullay.value = 1) Then
   If (txtManually.text = "") Then
      
      MsgBox "Enter Remarks...", vbInformation
      txtManually.SetFocus
      Screen.MousePointer = vbDefault
      Exit Sub
      
   End If
End If


If txtScId = "" Then
   MsgBox "Select School Name ...", vbCritical
   txtSchoolName.SetFocus
   Exit Sub
End If

If cmbAgentName.text = "" Then
   MsgBox "Select Representative Name ...", vbCritical
   cmbAgentName.SetFocus
   Exit Sub
End If

If cboPayment.text = "" Then
   MsgBox "Select Payment Mode ...", vbCritical
   cboPayment.SetFocus
   Exit Sub
End If

If txtWhomTobeGiven.text = "" Then
   MsgBox "Enter Whom To be Given ...", vbCritical
   txtWhomTobeGiven.SetFocus
   Exit Sub
End If


If rs1.State = 1 Then rs1.close
rs1.Open "select top 100 * from DonnationMain where DNo=" & txtSponsorshipNo.text & "", con, adOpenKeyset, adLockReadOnly
If rs1.EOF = False Then
   
   If RS.State = 1 Then RS.close
   RS.Open "select top 100 * from DonnationMain where DNo=" & txtSponsorshipNo.text & "", con, adOpenKeyset, adLockReadOnly
       If rs1!bAuthorized = True Then
          MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
          Exit Sub
      End If
End If


Dim rSave As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset


createLog UserName, txtSponsorshipNo, "donnation", " Modify : " & txtNetBal, Date

If RS.State = 1 Then RS.close
RS.Open "select * from DonnationMain where DNo=" & txtSponsorshipNo & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
    RS!waveoff = txtwave.text
    RS!dno = txtSponsorshipNo.text
    RS!DDate = txtDates.value
    RS!scid = txtScId.text
    RS!scname = txtSchoolName.text
    RS!RepName = cmbAgentName.text
    RS!RepName1 = cmbAgentName1.text
    RS!PaymentMode = cboPayment.text
    RS!WhomTobegiven = Trim(txtWhomTobeGiven.text)
    RS!Principal = Trim(txtPrincipal)
    RS!mobile = Trim(txtMob)
    RS!remarks = Trim(txtRemarks)
    RS!GrossAmt = Val(txtGTotal)
    RS!net = Val(txtNetTotal)
    RS!SponsorshipOn = cboSponse.text
    RS!ReturnAdj = Val(txtReturnAdj)
    RS!AmtAfter_ReturnAdj = Val(txtAmtAfterAdj)
    RS!Sponsorship_per = Val(txtPercentSp.text)
    RS!finalAmt = Val(txtFAmt)
    RS!AdvAmt = Val(txtAdvAmt.text)
    RS!RoundOfAAmt = Val(txtRoundOf)
    RS!NetBalance = Val(txtNetBal)
    RS!MobileWhomtoGiven = Trim(txtWhomToBeGivenMob)
    RS!manullay_Change = txtManually.text
    RS!createdby = UserName
    RS.update
    
    
    st_ = ""
    If rs1.State = 1 Then rs1.close
    rs1.Open "select subledger from DonnationQry_ where dno=" & txtSponsorshipNo.text & " and subledger is not null", con, adOpenDynamic, adLockReadOnly
    While rs1.EOF = False
         If st_ = "" Then
             st_ = rs1!subledger
         Else
             st_ = st_ & ", " & rs1!subledger
         End If
         rs1.MoveNext
     Wend
     con.Execute "update DonnationMain set party_='" & st_ & "'  where DNo=" & txtSponsorshipNo.text & ""

End If


If txtScId <> "" Then
   CON_blue.Execute "update College set Pr_Name='" & Trim(txtPrincipal) & "',Pr_Mobile='" & txtMob.text & "' where CollegeID='" & txtScId & "'"
End If


'============================================================================
'Donnation Details===========================================================
Dim f As New ADODB.Recordset

'============================================================================
For k1 = 1 To vs.rows - 1
If vs.TextMatrix(k1, 0) <> "" Then
  
  If rSave.State = 1 Then rSave.close
  rSave.Open "select * from DonnationMainDet where (dno=" & txtSponsorshipNo & " and invoiceNo=" & vs.TextMatrix(k1, 2) & " and bookcode='" & vs.TextMatrix(k1, 4) & "')", con, adOpenDynamic, adLockOptimistic
  If rSave.EOF = True Then
   rSave.AddNew
  End If
  
  rSave!dno = txtSponsorshipNo
  If Val(vs.TextMatrix(k1, 2)) = 0 Then
     rSave!invoiceNo = Null
  Else
     rSave!invoiceNo = vs.TextMatrix(k1, 2)
  End If
   
  If Not IsDate(vs.TextMatrix(k1, 3)) Then
     rSave!invoiceDate = Null
  Else
     rSave!invoiceDate = vs.TextMatrix(k1, 3)
  End If
  rSave!Bookcode = vs.TextMatrix(k1, 4)
  rSave!Bookname = vs.TextMatrix(k1, 5)
  rSave!qty = vs.TextMatrix(k1, 6)
  rSave!rate = vs.TextMatrix(k1, 7)
  rSave!GrossAmt = vs.TextMatrix(k1, 8)
  rSave!discount = vs.TextMatrix(k1, 9)
  rSave!net = vs.TextMatrix(k1, 10)
  If vs.TextMatrix(k1, 1) = "" Then
     rSave!Godown = Null
  Else
     rSave!Godown = vs.TextMatrix(k1, 1)
  End If
   
  If vs.TextMatrix(k1, 0) = "" Then
     rSave!fyear = Null
  Else
     rSave!fyear = vs.TextMatrix(k1, 0)
  End If
   
  rSave!scname = vs.TextMatrix(k1, 12)
  rSave!scid = vs.TextMatrix(k1, 13)
  rSave!UserName = UserName
  'rSave!AddNew = "Untracable"
  
  Set f = New ADODB.Recordset
  f.Open "select SUBLEDGER,states,DESCFORINVOICE,address1,address2,DISTCODE,mobile,RepName,id from tmpDonnation where (fyear='" & vs.TextMatrix(k1, 0) & "' and Godown='" & vs.TextMatrix(k1, 1) & "' and INVOICENO=" & vs.TextMatrix(k1, 2) & " and bookcode='" & vs.TextMatrix(k1, 4) & "' and sno=" & txtSponsorshipNo & ")", con
  If f.EOF = False Then
    rSave!subledger = f!subledger
    rSave!states = f!states
    rSave!DESCFORINVOICE = f!DESCFORINVOICE
    rSave!address1 = f!address1
    rSave!address2 = f!address2
    rSave!distcode = f!distcode
    rSave!mobile = f!mobile
    rSave!RepName = f!RepName
    rSave!id = f!id
  End If
  
  rSave.update
   
   
End If
Next

'============================================================================
'============================================================================



Dim d_, amt_ret, amt_sp
d_ = 0
amt_ret = 0
amt_sp = 0
d_1 = 0

If txtSponsorshipNo <> "" Then


If RS.State = 1 Then RS.close
RS.Open "select Sponsorship_per,ReturnAdj,GrossAmt,Net,SponsorshipOn,Godown,id from donation_Qry where DNo=" & txtSponsorshipNo & " order by BOOKCODE", con
While RS.EOF = False
   
  
  d_ = 0
  amt_ret = 0
  amt_sp = 0
   
   If RS!SponsorshipOn = "Gross" Then
      amt_ret = (RS!GrossAmt * RS!ReturnAdj / 100)
      amt_sp = RS!GrossAmt - amt_ret
      
   Else
      amt_ret = (RS!net * RS!ReturnAdj / 100)
      amt_sp = RS!net - amt_ret
   End If
   
   
   
   d_ = amt_sp * RS!Sponsorship_per / 100
   If RS!Godown = "C" Then
      d_ = -1 * d_
   End If
   
   con.Execute "update DonnationMainDet set DonationAmtBk=" & d_ & " where DNo=" & txtSponsorshipNo & " and id='" & RS!id & "'"
   RS.MoveNext
Wend



End If








'====================================================================================
'End Code============================================================================
'====================================================================================
Check1_manual.value = 0
con.Execute "exec tmpdata " & UId & ""

save.Enabled = False
cmdEdit_4.Enabled = True
Add = False
Edit = False
'===================================================================================
Screen.MousePointer = vbDefault

MsgBox "Modify Data....", vbInformation


Exit Sub
modify_:
Screen.MousePointer = vbDefault
MsgBox "" & err.DESCRIPTION

End Sub
Private Sub txtAdvAmt_Change()
calAmt
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
 txtRoundOf = Val(txtFAmt) - (Val(txtNetBal) + Val(txtAdvAmt))
End Sub

Private Sub txtPercentSp_Change()
calAmt
End Sub

Private Sub txtPercentSp_GotFocus()
HIT
End Sub

Private Sub txtPercentSp_LostFocus()

'Set rs1 = con.Execute("exec SearchAppData_ExDis '" & txtscid.Text & "'")
'If rs1.EOF = False Then
'   cboSponse.Text = rs1!Net_Gross
'   txtPercentSp.Text = rs1!Promo
   
   
If Check1_manullay.value = 0 Then
   
   If (Val(txtPercentSp.text) > Val(txtPercentSp1.text)) Then
      MsgBox "Enter only Less Extra Discount ...", vbCritical
      txtPercentSp.text = txtPercentSp1.text
   End If
   
   
   
End If

   
'End If

End Sub

Private Sub txtPrincipal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtMob.SetFocus

End Sub

Private Sub txtPrincipal_LostFocus()
txtPrincipal = UCase(txtPrincipal)
End Sub

Private Sub txtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cboSponse.SetFocus
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

'On Error Resume Next

sum1 = 0
AmtretAdj = 0
AmtSp = 0

If cboSponse.text = "Gross" Then
   sum1 = txtGTotal
Else
   sum1 = IIf(txtNetTotal = "", 0, txtNetTotal)
End If



AmtretAdj = Round((sum1 * Val(txtReturnAdj) / 100), 0)
txtAmtAfterAdj.text = sum1 - AmtretAdj
txtFAmt = Round((Val(txtAmtAfterAdj.text) * Val(txtPercentSp) / 100), 0)

txtRoundOf = Val(txtFAmt) - (Val(txtNetBal) + Val(txtAdvAmt))


'txtRoundOf = Round(Val(txtFAmt), 0)
'txtNetBal = (Val(txtFAmt) - Val(txtAdvAmt))



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



If PopUpValue1 <> "" Then

vs.Clear


If Check1_gpSchool.value = 0 Then

    If Check1_schoolAll.value = 0 Then
        txtSchoolName.text = PopUpValue1
        txtScId.text = PopUpValue2
    Else
        txtSchoolName.text = PopUpValue1 & "," & PopUpValue2
        txtScId.text = PopUpValue3
    End If

Else

        txtSchoolName.text = PopUpValue1
        If rs1.State = 1 Then rs1.close
        rs1.Open "select top 1 CollegeID from collegeView_ind where GroupOfSchool='" & PopUpValue1 & "'", CON_blue
        If rs1.EOF = False Then
           txtScId.text = rs1!collegeid
        End If

End If


End If

'=====================================================================
cboAppSer.Clear




If Check1_Addapp.value = 1 Then
    If txtScId.text <> "" Then
       If Check1_EditApp.value = 0 Then
          AddAproval
          fillSeries
       End If
       
    End If
Else
    con.Execute "delete from tmpAppForDonnation where uid=" & UId & ""
End If


If Check1_gpSchool.value = 0 Then
   filterData
   
    If rs1.State = 1 Then rs1.close
    rs1.Open "SELECT [DNo],[SCId],[ScName],[Remarks],[AdvAmt],[NetBalance],[UserName] FROM deleteDonnationMain " & _
    " where [SCId]='" & txtScId.text & "' and Status='Delete'", con
    If rs1.EOF = False Then
       txtSchoolName.text = rs1!scname
       txtScId.text = rs1!scid
       txtRemarks.text = rs1!remarks & ""
       txtAdvAmt.text = rs1!AdvAmt
       txtNetBal.text = rs1!NetBalance
    End If
   
Else
   
   
If txtSchoolName.text <> "" Then
   
Screen.MousePointer = vbHourglass

   vs.Clear
   con.Execute "delete from tmpDonnation where username='" & UserName & "'"
   If rss_.State = 1 Then rss_.close
   rss_.Open "SELECT CollegeID as ScID,GroupOfSchool FROM collegeView_ind  where GroupOfSchool= '" & txtSchoolName.text & "'", CON_blue
   While rss_.EOF = False
     Group_filterData rss_(0)
     rss_.MoveNext
   Wend
   
   
   
   
Screen.MousePointer = vbDefault

End If

   
End If


End Sub
Sub AddAproval()

App_serName = ""
con.Execute "delete from tmpAppForDonnation where uid=" & UId & ""

con.Execute "insert into tmpAppForDonnation(SerName,Adj,Promo,NetOrGross,PCode,Party,Scid,school,schoolOrParty,Appno,uid) " & _
" SELECT SerName,Adj,Promo,Net_Gross,Code,PName,id,School_PartyName,School_party,appno,'" & UId & "' " & _
" FROM AppForm where id='" & txtScId & "' group by SerName,Adj,Promo,Net_Gross,Code,PName,id,School_PartyName,School_party,appno"

con.Execute "insert into tmpAppForDonnation(SerName,Adj,Promo,NetOrGross,PCode,Party,Scid,school,schoolOrParty,Appno,uid) " & _
" select SerName,Adj,Promo,Net_Gross,id,School_PartyName,Code,PName,School_party,appno,'" & UId & "' " & _
" FROM AppForm where code='" & txtScId & "' group by SerName,Adj,Promo,Net_Gross,Code,PName,id,School_PartyName,School_party,appno"
    
DoEvents
DoEvents

Dim rsfill As New ADODB.Recordset
vs2.rows = 1
vs2.Cols = 11

If RS.State = 1 Then RS.close
RS.Open "SELECT count(SerName) as value_,SerName,Scid,School FROM tmpAppForDonnation where scid='" & txtScId & "' group by  SerName,Scid,School", con
If RS.EOF = False Then
   If RS(0) > 1 Then
      
      Frame1_app.Visible = True
      Check1_EditApp.value = 1
      Check1_EditApp.Enabled = False
   End If
   
      k11 = 1
      If rsfill.State = 1 Then rsfill.close
      rsfill.Open "select SerName,Adj,Promo as PromoPer,NetOrGross,PCode,Party,Scid,School,SchoolOrParty,Appno,UID from tmpAppForDonnation where (scid='" & txtScId & "' and uid=" & UId & ")", con
      While rsfill.EOF = False
      vs2.rows = vs2.rows + 1
      vs2.TextMatrix(k11, 0) = rsfill!sername
      vs2.TextMatrix(k11, 1) = rsfill!adj
      vs2.TextMatrix(k11, 2) = rsfill!PromoPer
      vs2.TextMatrix(k11, 3) = rsfill!netorgross
      vs2.TextMatrix(k11, 4) = rsfill!pcode
      vs2.TextMatrix(k11, 5) = rsfill!party
      vs2.TextMatrix(k11, 6) = rsfill!scid
      vs2.TextMatrix(k11, 7) = rsfill!school
      vs2.TextMatrix(k11, 8) = rsfill!schoolOrParty
      vs2.TextMatrix(k11, 9) = rsfill!appno
      vs2.TextMatrix(k11, 10) = rsfill!UId
      
      
      k11 = k11 + 1
      
      rsfill.MoveNext
      Wend
      
      vs2.FormatString = "SerName|AdjPer|PromoPer|NetorGross|PCode|Party|Scid|School|SchoolOrParty|AppNo|UID"
      vs2.ColWidth(0) = 1200
      vs2.ColWidth(1) = 900
      vs2.ColWidth(2) = 900
      vs2.ColWidth(3) = 1000
      vs2.ColWidth(4) = 1000
      vs2.ColWidth(5) = 2000
      vs2.ColWidth(6) = 1000
      vs2.ColWidth(7) = 2000
      vs2.ColWidth(8) = 1000
      vs2.ColWidth(9) = 600
      vs2.ColWidth(10) = 600
      
   Else
      Frame1_app.Visible = False
   End If
   

   
    
End Sub
Sub filterData()

On Error GoTo err1

HIT


If PopUpValue1 <> "" Then

    If Check1_schoolAll.value = 0 Then
        txtSchoolName.text = PopUpValue1
        txtScId.text = PopUpValue2
    Else
        txtSchoolName.text = PopUpValue1 & "," & PopUpValue2
        txtScId.text = PopUpValue3
    End If



Set rs1 = con.Execute("exec SearchAppData_ExDis '" & txtScId.text & "'")
If rs1.EOF = False Then
   
   If rs1!Net_Gross <> "" Then
      cboSponse.text = rs1!Net_Gross
      cboSponse1.text = rs1!Net_Gross
   End If
   
   txtPercentSp.text = rs1!Promo
   txtPercentSp1.text = rs1!Promo
End If



End If

If txtScId.text <> "" Then

vs.Clear



PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""


'====================================================================
Dim gps As String
Dim Qty_ As Long
bb_2 = False



If Check1_gp.value = 0 Then
    con.Execute "delete from tmpDonnation where username='" & UserName & "' and sno=" & txtSponsorshipNo & ""
    gps = "n"
Else
    gps = "gps"


'=====================================================================
If rs1.State = 1 Then rs1.close

rs1.Open "select top 1 * from tmpDonnation where Scid='" & txtScId.text & "' and username='" & UserName & "'", con, adOpenDynamic, adLockOptimistic
If rs1.EOF = False Then

    bb_2 = True
    GoTo dinesh:
    
End If

'=====================================================================
End If

'===============New Coding Sale============================================

If RS.State = 1 Then RS.close
If serName_ = "" Then
   RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,agentname from PartyWiseItemWiseQtySales where Scid='" & txtScId & "' and " & dt_str & " order by INVOICENO", con
Else
   RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,agentname from PartyWiseItemWiseQtySales where scid='" & txtScId & "' and " & dt_str & " and " & serName_ & " order by INVOICENO", con
End If
While RS.EOF = False
 

'-----------------------------------------------------------
If Check1_incAdj.value = 0 Then
     '''''check Adj
        If rs1.State = 1 Then rs1.close
        rs1.Open "select Qty FROM tmpSAdjDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='I' and fyear='" & RS!fyear & "'", con
        If rs1.EOF = False Then
           Qty_ = RS!qty - rs1(0)
        Else
           Qty_ = RS!qty
        End If
Else
     Qty_ = RS!qty
End If
'-----------------------------------------------------------


If rs1.State = 1 Then rs1.close
rs1.Open "select Qty FROM tmpDDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='I' and fyear='" & RS!fyear & "'", con
If rs1.EOF = False Then
   Qty_ = Qty_ - rs1(0)
End If


If Qty_ > 0 Then
  con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & Qty_ & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!net & "','" & RS!discount & "','I','" & RS!fyear & "','" & UserName & "','" & txtScId & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "')"
End If
RS.MoveNext



Wend


Qty_ = 0
'===============New Coding Sale Ret============================================
If RS.State = 1 Then RS.close
If serName_ = "" Then
RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Quantity,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,netamount,agentname from PartyWiseItemWiseQtySales_Return where Scid='" & txtScId & "' and " & dt_strR & " order by INVOICENO", con
Else
RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Quantity,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,netamount,agentname from PartyWiseItemWiseQtySales_Return where Scid='" & txtScId & "' and " & dt_strR & " and " & serName_ & "  order by INVOICENO", con
End If

While RS.EOF = False
 

'-----------------------------------------------------------
If Check1_incAdj.value = 0 Then
     '''''check Adj
        If rs1.State = 1 Then rs1.close
        rs1.Open "select Qty FROM tmpSAdjDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='C' and fyear='" & RS!fyear & "'", con
        If rs1.EOF = False Then
           Qty_ = RS!QUANTITY - rs1(0)
        Else
           Qty_ = RS!QUANTITY
        End If
Else
     Qty_ = RS!QUANTITY
End If
'-----------------------------------------------------------



 
If rs1.State = 1 Then rs1.close
'rs1.Open "select Qty FROM DonnationMainDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and fyear='" & RS!fyear & "'", con
rs1.Open "select Qty FROM tmpDDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='C' and fyear='" & RS!fyear & "'", con
If rs1.EOF = False Then
    Qty_ = Qty_ - rs1(0)
'Else
'   Qty_ = RS!Quantity
End If

   
 If Qty_ > 0 Then
    con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & RS!QUANTITY & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!netamount & "','" & RS!discount & "','C','" & RS!fyear & "','" & UserName & "','" & txtScId & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "')"
 End If
 RS.MoveNext
Wend


'=========================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=========================================================================



If db_ <> "no" Then

   
    
    Set CON_next = New ADODB.Connection
    If LCase(server_) = "server" Then
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & db_ & "; UID=" & sql_user & "; PWD=" & sql_pass
       CON_next.Open
    Else
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=HP\SQL2008; DATABASE=" & db_ & "; UID=; PWD=;"
       CON_next.Open
    End If

    '---------------------------'
    
    If RS.State = 1 Then RS.close
    If serName_ = "" Then
    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,agentname from PartyWiseItemWiseQtySales where Scid='" & txtScId & "' and " & dt_strSaleNext & " order by INVOICENO", CON_next
    Else
    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,agentname from PartyWiseItemWiseQtySales where Scid='" & txtScId & "' and " & dt_strSaleNext & " and " & serName_ & " order by INVOICENO", CON_next
    End If
    While RS.EOF = False
     
     
 
Qty_ = 0
'-----------------------------------------------------------
If Check1_incAdj.value = 0 Then
     '''''check Adj
        If rs1.State = 1 Then rs1.close
        rs1.Open "select Qty FROM tmpSAdjDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='I' and fyear='" & RS!fyear & "'", con
        If rs1.EOF = False Then
           Qty_ = RS!qty - rs1(0)
        Else
           Qty_ = RS!qty
        End If
Else
     Qty_ = RS!qty
End If
'-----------------------------------------------------------



       
    
    
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select Qty FROM tmpDDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='I' and fyear='" & RS!fyear & "'", con
    If rs1.EOF = False Then
        Qty_ = Qty_ - rs1(0)
    'Else
    ''Qty_ = RS!qty
    End If
       
       
     If Qty_ > 0 Then
       con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & Qty_ & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!net & "','" & RS!discount & "','I','" & RS!fyear & "','" & UserName & "','" & txtScId & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "')"
     End If
    
    
    
    RS.MoveNext
    Wend
    
    'SaleReturn
    If RS.State = 1 Then RS.close
    If serName_ = "" Then
    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QUANTITY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,NETAMOUNT,DISCOUNT,Godown,fyear,agentname from PartyWiseItemWiseQtySales_Return where Scid='" & txtScId & "' and " & dt_strSaleRNext & " order by INVOICENO", CON_next
    Else
    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QUANTITY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,NETAMOUNT,DISCOUNT,Godown,fyear,agentname from PartyWiseItemWiseQtySales_Return where Scid='" & txtScId & "' and " & dt_strSaleRNext & " and " & serName_ & " order by INVOICENO", CON_next
    End If
    While RS.EOF = False
    
    
    
    Qty_ = 0
    
    
    '-----------------------------------------------------------
If Check1_incAdj.value = 0 Then
     '''''check Adj
        If rs1.State = 1 Then rs1.close
        rs1.Open "select Qty FROM tmpSAdjDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='C' and fyear='" & RS!fyear & "'", con
        If rs1.EOF = False Then
           Qty_ = RS!QUANTITY - rs1(0)
        Else
           Qty_ = RS!QUANTITY
        End If
Else
     Qty_ = RS!QUANTITY
End If
'-----------------------------------------------------------

    
    
    
    If rs1.State = 1 Then rs1.close
    ''rs1.Open "select Qty FROM DonnationMainDet where INVOICENO='" & RS!INVOICENO & "' and BOOKCODE='" & RS!Bookcode & "' and fyear='" & RS!fyear & "'", con
    rs1.Open "select Qty FROM tmpDDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "'  and Godown='C' and fyear='" & RS!fyear & "'", con
    If rs1.EOF = False Then
       Qty_ = Qty_ - rs1(0)
   ' Else
   '    Qty_ = RS!Quantity
    End If
    
    
    
       
    If Qty_ > 0 Then
       con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & RS!QUANTITY & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!netamount & "','" & RS!discount & "','C','" & RS!fyear & "','" & UserName & "','" & txtScId & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "')"
    End If
    RS.MoveNext
    
    
    Wend


End If '''' Next Data

'------end code-------------'

End If

'cboSponse.ListIndex = 0
If cboSponse.text = "Gross" Then
   sum1 = Val(txtGTotal)
Else
   sum1 = Val(txtNetTotal)
End If


'=========================================================================

If txtScId <> "" Then

    List1_sc.Clear
    If RS.State = 1 Then RS.close
    RS.Open "SELECT DNO,SCId,ScName FROM DonnationMain where scid='" & txtScId & "'", con
    While RS.EOF = False
          List1_sc.AddItem RS(0) & "=>" & RS(1) & ":" & RS(2)
    RS.MoveNext
    Wend
    
    
    '=++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++========
    'Fatch Data from School Data----------------------------------------------
    txtPrincipal = ""
    txtMob = ""
    If RS.State = 1 Then RS.close
    RS.Open "select top 1 Pr_Name,Pr_mobile from collegeView_ind where CollegeID='" & txtScId & "'", CON_blue
    If RS.EOF = False Then
       txtPrincipal = RS!Pr_Name
       txtMob = RS!Pr_mobile
    End If
    '************************************************************************
    
    If txtSponsorshipNo <> "" Then
        If RS.State = 1 Then RS.close
        RS.Open "select distinct repname from tmpDonnation where sno='" & txtSponsorshipNo & "' and username='" & UserName & "' ", con
        If RS.RecordCount > 1 Then
           cmbAgentName.text = RS(0)
           RS.MoveNext
           cmbAgentName1.text = RS(0)
        ElseIf RS.RecordCount = 1 Then
           cmbAgentName.text = RS(0)
        End If
        
        
        
    End If
    
End If



dinesh:

fillGrid_


If bb_2 = True Then
   MsgBox "This School Alreay Added...", vbCritical
   txtSchoolName.text = ""
   txtScId.text = ""
   txtSchoolName.SetFocus
End If




Exit Sub
err1:

MsgBox "" & err.DESCRIPTION

End Sub
Sub Group_filterData(scid As String)

On Error GoTo err1

HIT


If kk1 = 0 Then


If PopUpValue1 <> "" Then

'    If Check1_schoolAll.value = 0 Then
'        txtSchoolName.text = PopUpValue1
'        txtscid.text = PopUpValue2
'    Else
'        txtSchoolName.text = PopUpValue1 & "," & PopUpValue2
'        txtscid.text = PopUpValue3
'    End If



Set rs1 = con.Execute("exec SearchAppData_ExDis '" & scid & "'")
If rs1.EOF = False Then
   
   If rs1!Net_Gross <> "" Then
      cboSponse.text = rs1!Net_Gross
   End If
   
   txtPercentSp.text = rs1!Promo
   txtPercentSp1.text = rs1!Promo
End If




End If

vs.Clear
kk1 = kk1 + 1


End If


If txtScId.text <> "" Then



PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""


'====================================================================
Dim gps As String
Dim Qty_ As Long
bb_2 = False


'===============New Coding Sale============================================

If RS.State = 1 Then RS.close
If serName_ = "" Then
   RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,agentname from ExtraDiscountQry where Godown='I' and Scid='" & scid & "' and " & dt_str & " order by INVOICENO", con
Else
   RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,agentname from ExtraDiscountQry where Godown='I' and scid='" & scid & "' and " & dt_str & " and " & serName_ & " order by INVOICENO", con
End If
While RS.EOF = False
 

'-----------------------------------------------------------
If Check1_incAdj.value = 0 Then
     '''''check Adj
        If rs1.State = 1 Then rs1.close
        rs1.Open "select Qty FROM ExtraDiscountSalesAdjQry where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='I' and fyear='" & RS!fyear & "'", con
        If rs1.EOF = False Then
           Qty_ = RS!qty - rs1(0)
        Else
           Qty_ = RS!qty
        End If
Else
     Qty_ = RS!qty
End If
'-----------------------------------------------------------


If rs1.State = 1 Then rs1.close
rs1.Open "select Qty FROM ExtraDiscountDonnationQry where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='I' and fyear='" & RS!fyear & "'", con
If rs1.EOF = False Then
   Qty_ = Qty_ - rs1(0)
End If


If Qty_ > 0 Then
  con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & Qty_ & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!net & "','" & RS!discount & "','I','" & RS!fyear & "','" & UserName & "','" & txtScId & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "')"
End If
RS.MoveNext



Wend


Qty_ = 0
'===============New Coding Sale Ret============================================
If RS.State = 1 Then RS.close
If serName_ = "" Then
RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty as QUANTITY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,netamount,agentname from ExtraDiscountQry where Godown='C' and Scid='" & txtScId & "' and " & dt_strR & " order by INVOICENO", con
Else
RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty as QUANTITY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,netamount,agentname from ExtraDiscountQry where Godown='C' and Scid='" & txtScId & "' and " & dt_strR & " and " & serName_ & "  order by INVOICENO", con
End If

While RS.EOF = False
 

'-----------------------------------------------------------
If Check1_incAdj.value = 0 Then
     '''''check Adj
        If rs1.State = 1 Then rs1.close
        rs1.Open "select Qty FROM ExtraDiscountSalesAdjQry where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='C' and fyear='" & RS!fyear & "'", con
        If rs1.EOF = False Then
           Qty_ = RS!QUANTITY - rs1(0)
        Else
           Qty_ = RS!QUANTITY
        End If
Else
     Qty_ = RS!QUANTITY
End If
'-----------------------------------------------------------



 
If rs1.State = 1 Then rs1.close
rs1.Open "select Qty FROM ExtraDiscountDonnationQry where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='C' and fyear='" & RS!fyear & "'", con
If rs1.EOF = False Then
    Qty_ = Qty_ - rs1(0)
End If

   
 If Qty_ > 0 Then
    con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & RS!QUANTITY & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!netamount & "','" & RS!discount & "','C','" & RS!fyear & "','" & UserName & "','" & txtScId & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "')"
 End If
 RS.MoveNext
Wend


'=========================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=========================================================================



'If db_ <> "no" Then

   
    
'    Set CON_next = New ADODB.Connection
'    If LCase(server_) = "server" Then
'       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & db_ & "; UID=" & sql_user & "; PWD=" & sql_pass
'       CON_next.Open
'    Else
'       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=HP\SQL2008; DATABASE=" & db_ & "; UID=; PWD=;"
'       CON_next.Open
'    End If

    '---------------------------'
    
    If RS.State = 1 Then RS.close
    If serName_ = "" Then
    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,agentname from ExtraDiscountQry where Godown='I' and Scid='" & scid & "' and " & dt_strSaleNext & " order by INVOICENO", con
    Else
    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,agentname from ExtraDiscountQry where Godown='I' and Scid='" & scid & "' and " & dt_strSaleNext & " and " & serName_ & " order by INVOICENO", con
    End If
    While RS.EOF = False
     
     
 
    Qty_ = 0
    '-----------------------------------------------------------
    If Check1_incAdj.value = 0 Then
         '''''check Adj
            If rs1.State = 1 Then rs1.close
            rs1.Open "select Qty FROM ExtraDiscountSalesAdjQry where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='I' and fyear='" & RS!fyear & "'", con
            If rs1.EOF = False Then
               Qty_ = RS!qty - rs1(0)
            Else
               Qty_ = RS!qty
            End If
    Else
         Qty_ = RS!qty
    End If
    '-----------------------------------------------------------
    

    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select Qty FROM ExtraDiscountDonnationQry where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='I' and fyear='" & RS!fyear & "'", con
    If rs1.EOF = False Then
        Qty_ = Qty_ - rs1(0)

    End If
       
       
     If Qty_ > 0 Then
       con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & Qty_ & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!net & "','" & RS!discount & "','I','" & RS!fyear & "','" & UserName & "','" & scid & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "')"
     End If
    
    
    
    RS.MoveNext
    Wend
    
    
    'SaleReturn
    If RS.State = 1 Then RS.close
    If serName_ = "" Then
    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty as QUANTITY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,NETAMOUNT,DISCOUNT,Godown,fyear,agentname from ExtraDiscountQry where Godown='C' and Scid='" & scid & "' and " & dt_strSaleRNext & " order by INVOICENO", con
    Else
    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty as QUANTITY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,NETAMOUNT,DISCOUNT,Godown,fyear,agentname from ExtraDiscountQry where Godown='C' and Scid='" & scid & "' and " & dt_strSaleRNext & " and " & serName_ & " order by INVOICENO", con
    End If
    While RS.EOF = False
    
    
    
    Qty_ = 0
    
    
    '-----------------------------------------------------------
   If Check1_incAdj.value = 0 Then
     '''''check Adj
        If rs1.State = 1 Then rs1.close
        rs1.Open "select Qty FROM ExtraDiscountSalesAdjQry where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and Godown='C' and fyear='" & RS!fyear & "'", con
        If rs1.EOF = False Then
           Qty_ = RS!QUANTITY - rs1(0)
        Else
           Qty_ = RS!QUANTITY
        End If
    Else
         Qty_ = RS!QUANTITY
    End If
'-----------------------------------------------------------

    
    

    If rs1.State = 1 Then rs1.close
    rs1.Open "select Qty FROM ExtraDiscountDonnationQry where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "'  and Godown='C' and fyear='" & RS!fyear & "'", con
    If rs1.EOF = False Then
       Qty_ = Qty_ - rs1(0)
    End If

    
    
       
    If Qty_ > 0 Then
       con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & RS!QUANTITY & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!netamount & "','" & RS!discount & "','C','" & RS!fyear & "','" & UserName & "','" & scid & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "')"
    End If
    RS.MoveNext
    
    
    Wend


'''End If '''' Next Data

'------end code-------------'

End If

cboSponse.ListIndex = 0
If cboSponse.text = "Gross" Then
   sum1 = Val(txtGTotal)
Else
   sum1 = Val(txtNetTotal)
End If


'=========================================================================

If scid <> "" Then

    List1_sc.Clear
    If RS.State = 1 Then RS.close
    RS.Open "SELECT DNO,SCId,ScName FROM DonnationMain where scid='" & scid & "'", con
    While RS.EOF = False
          List1_sc.AddItem RS(0) & "=>" & RS(1) & ":" & RS(2)
    RS.MoveNext
    Wend
    
    
    '=++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++========
    'Fatch Data from School Data----------------------------------------------
    txtPrincipal = ""
    txtMob = ""
    If RS.State = 1 Then RS.close
    RS.Open "select top 1 Pr_Name,Pr_mobile from collegeView_ind where CollegeID='" & scid & "'", CON_blue
    If RS.EOF = False Then
       txtPrincipal = RS!Pr_Name
       txtMob = RS!Pr_mobile
    End If
    '************************************************************************
    
    If txtSponsorshipNo <> "" Then
        If RS.State = 1 Then RS.close
        RS.Open "select distinct repname from tmpDonnation where sno='" & txtSponsorshipNo & "' and username='" & UserName & "' ", con
        If RS.RecordCount > 1 Then
           cmbAgentName.text = RS(0)
           RS.MoveNext
           cmbAgentName1.text = RS(0)
        ElseIf RS.RecordCount = 1 Then
           cmbAgentName.text = RS(0)
        End If
        
        
        
    End If
    
End If



dinesh:

fillGrid_


If bb_2 = True Then
   MsgBox "This School Alreay Added...", vbCritical
   txtSchoolName.text = ""
   txtScId.text = ""
   txtSchoolName.SetFocus
End If




Exit Sub
err1:

MsgBox "" & err.DESCRIPTION

End Sub


Sub fatchData()


'==========================================================================
Dim gps As String
Dim Qty_ As Long
bb_2 = False
'===============New Coding Sale============================================

If RS.State = 1 Then RS.close
RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,agentname from PartyWiseItemWiseQtySales where Scid='" & txtScId & "' and " & dt_str & " order by INVOICENO", con
While RS.EOF = False
 

If rs1.State = 1 Then rs1.close
rs1.Open "select Qty FROM tmpDDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and fyear='" & RS!fyear & "'", con
If rs1.EOF = False Then
   Qty_ = RS!qty - rs1(0)
Else
   Qty_ = RS!qty
End If



If Qty_ > 0 Then
  con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & Qty_ & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!net & "','" & RS!discount & "','I','" & RS!fyear & "','" & UserName & "','" & txtScId & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "')"
End If
RS.MoveNext


Wend


Qty_ = 0
'===============New Coding Sale Ret============================================
If RS.State = 1 Then RS.close
RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Quantity,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,netamount,agentname from PartyWiseItemWiseQtySales_Return where Scid='" & txtScId & "' and " & dt_strR & " order by INVOICENO", con
While RS.EOF = False
 
If rs1.State = 1 Then rs1.close
rs1.Open "select Qty FROM tmpDDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and fyear='" & RS!fyear & "'", con
If rs1.EOF = False Then
   Qty_ = RS!QUANTITY - rs1(0)
Else
   Qty_ = RS!QUANTITY
End If
   
 If Qty_ > 0 Then
    con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & RS!QUANTITY & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!netamount & "','" & RS!discount & "','C','" & RS!fyear & "','" & UserName & "','" & txtScId & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "')"
 End If
 RS.MoveNext
Wend


'=========================================================================
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'=========================================================================



If db_ <> "no" Then

    Set CON_next = New ADODB.Connection
    If LCase(server_) = "server" Then
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=" & serverName_ & "; DATABASE=" & db_ & "; UID=" & sql_user & "; PWD=" & sql_pass
       CON_next.Open
    Else
       CON_next.ConnectionString = "Provider=MSDASQL; DRIVER=Sql Server; SERVER=HP\SQL2008; DATABASE=" & db_ & "; UID=; PWD=;"
       CON_next.Open
    End If



    '---------------------------'
    If RS.State = 1 Then RS.close
    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,agentname from PartyWiseItemWiseQtySales where Scid='" & txtScId & "' and " & dt_strSaleNext & " order by INVOICENO", CON_next
    While RS.EOF = False
     
       
    If rs1.State = 1 Then rs1.close
    rs1.Open "select Qty FROM tmpDDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and fyear='" & RS!fyear & "'", con
    If rs1.EOF = False Then
        Qty_ = Qty_ - rs1(0)
    Else
        Qty_ = RS!qty
    End If
       
       
     If Qty_ > 0 Then
       con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & Qty_ & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!net & "','" & RS!discount & "','I','" & RS!fyear & "','" & UserName & "','" & txtScId & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "')"
     End If
    
    
    
    RS.MoveNext
    Wend
    
    'SaleReturn
    If RS.State = 1 Then RS.close
    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QUANTITY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,NETAMOUNT,DISCOUNT,Godown,fyear,agentname from PartyWiseItemWiseQtySales_Return where Scid='" & txtScId & "' and " & dt_strSaleRNext & " order by INVOICENO", CON_next
    While RS.EOF = False
    
    
    
    Qty_ = 0
    
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select Qty FROM tmpDDet where INVOICENO='" & RS!invoiceNo & "' and BOOKCODE='" & RS!Bookcode & "' and fyear='" & RS!fyear & "'", con
    If rs1.EOF = False Then
       Qty_ = RS!QUANTITY - rs1(0)
    Else
       Qty_ = RS!QUANTITY
    End If
    
    
    
       
    If Qty_ > 0 Then
       con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno,repname) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & RS!QUANTITY & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!netamount & "','" & RS!discount & "','C','" & RS!fyear & "','" & UserName & "','" & txtScId & "','" & gps & "'," & txtSponsorshipNo & ",'" & RS!agentname & "')"
    End If
    RS.MoveNext
    
    
    Wend


End If '''' Next Data

'------end code-------------'



'=========================================================================

Exit Sub
err1:

MsgBox "" & err.DESCRIPTION


End Sub
Sub fillGrid_()

Dim fy As Integer
Dim rrs As New ADODB.Recordset

Set rrs = New ADODB.Recordset

fy = 0

fy = Right(session, 2) - 1


txtGTotal = 0
txtNetTotal = 0
txtQty = 0
vs.Cols = 15

If RS.State = 1 Then RS.close
RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QTY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,scid from tmpDonnation where userNAME='" & UserName & "' and sno=" & txtSponsorshipNo & " order by id", con
If RS.EOF = False Then
vs.rows = RS.RecordCount + 100
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

''vs.TextMatrix(I, 9) = RS!discount
''vs.TextMatrix(I, 10) = Round(RS!net, 0)
''If RS!Godown = "I" Then
''    txtGTotal = Val(txtGTotal) + RS!GrossAmt
''    txtNetTotal = Val(txtNetTotal) + Round(RS!net, 0)
''Else
''    txtGTotal = Val(txtGTotal) - RS!GrossAmt
''    txtNetTotal = Val(txtNetTotal) - Round(RS!net, 0)
''End If

If fy < 23 Then

    vs.TextMatrix(I, 9) = RS!discount
    vs.TextMatrix(I, 10) = Round(RS!net, 0)
    If RS!Godown = "I" Then
        txtGTotal = Val(txtGTotal) + RS!GrossAmt
        txtNetTotal = Val(txtNetTotal) + Round(RS!net, 0)
    Else
        txtGTotal = Val(txtGTotal) - RS!GrossAmt
        txtNetTotal = Val(txtNetTotal) - Round(RS!net, 0)
    End If

Else

   If cboSponse.text = "Net" Then
        If rrs.State = 1 Then rrs.close
        Set rrs = con.Execute("exec SearchAppDataExDis '" & txtScId.text & "','" & RS!Bookcode & "'")
        If rrs.EOF = False Then
            vs.TextMatrix(I, 9) = Round(rrs(0), 2)
            vs.TextMatrix(I, 10) = Round((RS!GrossAmt * rrs(0) / 100), 2)    'Round(RS!net, 0)
            If RS!Godown = "I" Then
                txtGTotal = Val(txtGTotal) + RS!GrossAmt
                txtNetTotal = Val(txtNetTotal) + Round(vs.TextMatrix(I, 10), 2)
            Else
                txtGTotal = Val(txtGTotal) - RS!GrossAmt
                txtNetTotal = Val(txtNetTotal) - Round(vs.TextMatrix(I, 10), 2)
            End If

        Else

            vs.TextMatrix(I, 9) = RS!discount
            vs.TextMatrix(I, 10) = Round(RS!net, 0)
            If RS!Godown = "I" Then
            txtGTotal = Val(txtGTotal) + RS!GrossAmt
            txtNetTotal = Val(txtNetTotal) + Round(RS!net, 0)
            Else
            txtGTotal = Val(txtGTotal) - RS!GrossAmt
            txtNetTotal = Val(txtNetTotal) - Round(RS!net, 0)
            End If

        End If

   Else

        vs.TextMatrix(I, 9) = RS!discount
        vs.TextMatrix(I, 10) = Round(RS!net, 0)
        If RS!Godown = "I" Then
            txtGTotal = Val(txtGTotal) + RS!GrossAmt
            txtNetTotal = Val(txtNetTotal) + Round(RS!net, 0)
        Else
            txtGTotal = Val(txtGTotal) - RS!GrossAmt
            txtNetTotal = Val(txtNetTotal) - Round(RS!net, 0)
        End If

   End If

End If



vs.TextMatrix(I, 12) = RS!scname
vs.TextMatrix(I, 13) = RS!scid

txtQty = Val(txtQty) + RS!qty

RS.MoveNext

Next

txtGTotal = Round(txtGTotal, 0)

vs.FormatString = "Session|Inv.Type|Inv. No.|Inv. Date|B.Code|B.Name|>Quantity|>Rate|>GrossAmt|>Dis.%|>NetAmt|AddNew|School Name|SCId"

vs.ColWidth(1) = 700
vs.ColWidth(2) = 700
vs.ColWidth(3) = 900
vs.ColWidth(4) = 700
vs.ColWidth(5) = 2800
vs.ColWidth(6) = 800
vs.ColWidth(7) = 700
vs.ColWidth(8) = 950

vs.ColWidth(9) = 600
vs.ColWidth(10) = 750

vs.ColWidth(11) = 750
vs.ColWidth(12) = 2050
vs.ColWidth(13) = 750

''''''==============================================================
''''''==============================================================
'''''If rs1.State = 1 Then rs1.close
'''''rs1.Open "SELECT distinct  AppForm.AppPer,Promo,Net_Gross FROM AppForm INNER JOIN BOOKS ON AppForm.SerName = BOOKS.SerName " & _
'''''" where (AppForm.PName='" & txtSchoolName.Text & "' or School_PartyName='" & txtSchoolName.Text & "') and BOOKCODE='" & vs.TextMatrix(1, 4) & "' and Promo>0", con
'''''If rs1.EOF = False Then
'''''       a11 = IIf(rs1!promo = Null, 0, rs1!promo)
'''''       txtPercentSp.Text = a11
'''''       If (Not IsNull(rs1!Net_Gross)) Then
'''''       If Len(rs1!Net_Gross) > 1 Then
'''''          cboSponse.Text = rs1!Net_Gross
'''''        End If
'''''       End If
'''''End If
''''''==============================================================

End Sub
Private Sub txtSchoolName_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    
    
    
    
 If Check1_gpSchool.value = 0 Then
    
    
    If Check1_schoolAll.value = 0 Then
        
        searchType = "party"
        value = "SELECT des as ScName,Billtype as ScID FROM tempLedger_net group by des,Billtype"
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
        searchType = ""
        value = "SELECT GroupOfSchool FROM collegeView_ind  where GroupOfSchool<> '' group by GroupOfSchool"
        popuplist10 value, CON_blue
        set_focus = True
        Screen.MousePointer = vbDefault


    
End If
    
    
End If


If KeyCode = 13 Then



  cmbAgentName.SetFocus
End If

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
    popuplist10 "select DNo,DDate,ScName,RoundOfAAmt from DonnationMain order by DNo", con
End If

If KeyCode = 13 Then
     'If MsgBox("Want to Edid... ", vbQuestion + vbYesNo) = vbYes Then
     '   CON.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,sNo) select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,Fyear,UserName,scid,DNo from DonnationMainDet where dno=" & txtSponsorshipNo & ""
     'End If

   searchData
   
    If rs1.State = 1 Then rs1.close
    rs1.Open "SELECT [DNo],[SCId],[ScName],[Remarks],[AdvAmt],[NetBalance],[UserName] FROM deleteDonnationMain " & _
    " where [DNo]='" & txtSponsorshipNo & "' and Status='Delete'", con
    If rs1.EOF = False Then
       txtSchoolName.text = rs1!scname
       txtScId.text = rs1!scid
       txtRemarks.text = rs1!remarks & ""
       txtAdvAmt.text = rs1!AdvAmt
       txtNetBal.text = rs1!NetBalance
    End If
   
   txtDates.SetFocus
End If
End Sub
Private Sub txtSponsorshipNo_LostFocus()
 
''' If txtSponsorshipNo <> "" Then
'''     PopUpValue1 = txtSponsorshipNo
'''     searchData
'''     PopUpValue1 = ""
''' End If

End Sub

Private Sub txtwave_LostFocus()
'' If txtRoundOf.Text = "" Then
    Dim ramt As Double
    ramt = Val(txtFAmt) - (Val(txtNetBal) + Val(txtAdvAmt))


    If Val(txtRoundOf.text) > 0 Then
       MsgBox "You can not wave off ", vbCritical
       ''txtwave.SetFocus
    Else
    
       txtRoundOf.text = Val(ramt) + Val(txtwave.text)
       
    End If
''End If
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

For I = 1 To vs.rows - 1
If vs.TextMatrix(I, 1) <> "" Then
If vs.TextMatrix(I, 1) = "I" Then
    txtGTotal = Val(txtGTotal) + vs.TextMatrix(I, 8)
    txtNetTotal = Val(txtNetTotal) + vs.TextMatrix(I, 10)
    txtQty = Val(txtQty) + Val(vs.TextMatrix(I, 7))
Else
    txtGTotal = Val(txtGTotal) - vs.TextMatrix(I, 8)
    txtNetTotal = Val(txtNetTotal) - vs.TextMatrix(I, 10)
    txtQty = Val(txtQty) + Val(vs.TextMatrix(I, 7))
End If

End If
Next


txtGTotal = Round(txtGTotal, 0)

End Sub

Private Sub txtWhomToBeGivenMob_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtPrincipal.SetFocus
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)
    
On Error Resume Next


If KeyCode = 115 Then
   
   con.Execute "update tmpDonnation set addnew='n' where sNo=" & txtSponsorshipNo & " and BOOKCODE='" & vs.TextMatrix(vs.RowSel, 4) & "' and INVOICENO=" & vs.TextMatrix(vs.RowSel, 2) & " and username='" & UserName & "'"
   
   If rss_10.State = 1 Then rss_10.close
   rss_10.Open "select * from DonnationMainDet where DNo=" & txtSponsorshipNo & " and fyear='" & vs.TextMatrix(vs.RowSel, 0) & "' and BOOKCODE='" & vs.TextMatrix(vs.RowSel, 4) & "' and INVOICENO=" & vs.TextMatrix(vs.RowSel, 2) & "", con
   If rss_10.EOF = False Then
     If MsgBox("want to delete ?", vbQuestion + vbYesNo) = vbYes Then
       If RS.State = 1 Then RS.close
       RS.Open "select top 100 * from DonnationMain where DNo=" & txtSponsorshipNo.text & "", con, adOpenKeyset, adLockReadOnly
       If RS!bAuthorized = True Then
          MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
          Exit Sub
      End If

       con.Execute "delete from DonnationMainDet where DNo=" & txtSponsorshipNo & " and fyear='" & vs.TextMatrix(vs.RowSel, 0) & "' and BOOKCODE='" & vs.TextMatrix(vs.RowSel, 4) & "' and INVOICENO=" & vs.TextMatrix(vs.RowSel, 2) & ""
       vs.SetFocus
     End If
   End If
   
   vs.RemoveItem vs.RowSel
   Total
   calAmt
   vs.SetFocus
   
''''==============================================================
''''==============================================================
'''If rs1.State = 1 Then rs1.close
'''rs1.Open "SELECT distinct  AppForm.AppPer,Promo,Net_Gross FROM AppForm INNER JOIN BOOKS ON AppForm.SerName = BOOKS.SerName " & _
'''" where (AppForm.PName='" & txtSchoolName.Text & "' or School_PartyName='" & txtSchoolName.Text & "') and BOOKCODE='" & vs.TextMatrix(1, 4) & "'", con
'''If rs1.EOF = False Then
'''       a11 = IIf(rs1!promo = Null, 0, rs1!promo)
'''       txtPercentSp.Text = a11
'''       If Not IsNull(rs1!Net_Gross) Then
'''          cboSponse.Text = rs1!Net_Gross
'''       End If
'''End If
''''===============================================================
   
End If
End Sub
Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
  
On Error GoTo abc1:
  
  
  Dim dis_ As Double
  Dim tblName As String
  
  
  If KeyCode = 13 Then
  
    'Untraceable----------------------------------------------------------------------------------
    
    If Check1_untracableSc.value = 1 Then
       
       If (vs.Col = 0 Or vs.Col = 1 Or vs.Col = 2 Or vs.Col = 3) Then
           sendkeys "{right}"
       ElseIf (vs.Col = 4) Then
           If RS.State = 1 Then RS.close
           RS.Open "select BOOKCODE,BOOKNAME,rate,discount from BOOKS where BOOKCODE='" & vs.TextMatrix(vs.RowSel, 4) & "'", con
           If RS.EOF = False Then
              vs.TextMatrix(vs.RowSel, 4) = UCase(RS!Bookcode)
              vs.TextMatrix(vs.RowSel, 5) = UCase(RS!Bookname)
              vs.TextMatrix(vs.RowSel, 7) = UCase(RS!rate)
              vs.TextMatrix(vs.RowSel, 9) = UCase(RS!discount)
              vs.TextMatrix(vs.RowSel, 12) = txtSchoolName.text
              vs.TextMatrix(vs.RowSel, 13) = txtScId.text

              sendkeys "{right}"
              sendkeys "{right}"
           End If
       ElseIf (vs.Col = 6) Then
           vs.TextMatrix(vs.RowSel, 8) = Round((Val(vs.TextMatrix(vs.RowSel, 6)) * Val(vs.TextMatrix(vs.RowSel, 7))), 0)
            
            dis_ = Round((vs.TextMatrix(vs.RowSel, 8) * vs.TextMatrix(vs.RowSel, 9) / 100), 0)
            vs.TextMatrix(vs.RowSel, 10) = Round(Val(vs.TextMatrix(vs.RowSel, 8)) - dis_, 0)
            
            vs.TextMatrix(vs.RowSel, 12) = txtSchoolName.text
            vs.TextMatrix(vs.RowSel, 13) = txtScId.text
            
            sendkeys "{down}"
            sendkeys "{home}"
            Total

       End If
       
       
       
       Exit Sub
    End If
    '---------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------
    '---------------------------------------------------------------------------------------------
    
    
     
     If (vs.Col = 0 Or vs.Col = 1 Or vs.Col = 2 Or vs.Col = 3) Then
        
           
       If (Len(vs.TextMatrix(vs.RowSel, 0)) > 0 And Len(vs.TextMatrix(vs.RowSel, 1)) > 0 And Len(vs.TextMatrix(vs.RowSel, 2)) > 0) Then
         tblName = fatchDate(vs.TextMatrix(vs.RowSel, 0), vs.TextMatrix(vs.RowSel, 1), vs.TextMatrix(vs.RowSel, 2), 0)
         
         If MsgBox("Want to Add Bill", vbQuestion + vbYesNo) = vbYes Then
            
          If vs.TextMatrix(vs.RowSel, 1) = "I" Then
            
                If current_next = "current" Then
                    con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno) select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,'I','" & session & "','" & UserName & "',scid,'" & gps & "','" & txtSponsorshipNo & "' from PartyWiseItemWiseQtySales where INVOICENO=" & vs.TextMatrix(vs.RowSel, 2) & ""
                Else
                
                
                 If db_ <> "no" Then
                 
                    If RS.State = 1 Then RS.close
                    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,Qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,Net,DISCOUNT,Godown,fyear,scid from PartyWiseItemWiseQtySales where INVOICENO='" & vs.TextMatrix(vs.RowSel, 2) & "' and " & dt_str & "", CON_next
                    While RS.EOF = False
                      con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & RS!qty & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!net & "','" & RS!discount & "','I','" & RS!fyear & "','" & UserName & "','" & RS!scid & "','" & gps & "'," & txtSponsorshipNo & ")"
                      RS.MoveNext
                    Wend
                    
                 End If
                    

                End If
            
                fillGrid_
           Else
            
                If vs.TextMatrix(vs.RowSel, 0) = session Then
                    con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno) select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QUANTITY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,NETAMOUNT,DISCOUNT,'C','" & session & "','" & UserName & "',scid,'" & gps & "','" & txtSponsorshipNo & "' from PartyWiseItemWiseQtySales_Return where INVOICENO=" & vs.TextMatrix(vs.RowSel, 2) & ""
                 Else
                 
                 If db_ <> "no" Then
                    If RS.State = 1 Then RS.close
                    RS.Open "select INVOICENO,SUBLEDGER,BOOKCODE,BOOKNAME,states,QUANTITY,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,mobile,INVOICEDATE,GrossAmt,RATE,ScName,NETAMOUNT,DISCOUNT,Godown,fyear,scid from PartyWiseItemWiseQtySales_Return where INVOICENO='" & vs.TextMatrix(vs.RowSel, 2) & "' and " & dt_str & "", CON_next
                    While RS.EOF = False
                     con.Execute "insert into tmpDonnation(INVOICENO,SUBLEDGER,BOOKCODE,BOOKName,states,qty,DESCFORINVOICE,ADDRESS1,ADDRESS2,DISTCODE,Mobile,INVOICEDATE,GrossAmt,rate,ScName,net,DISCOUNT,Godown,Fyear,UserName,ScId,gps,sno) values(" & RS!invoiceNo & ",'" & RS!subledger & "','" & RS!Bookcode & "','" & RS!Bookname & "','" & RS!states & "','" & RS!QUANTITY & "','" & RS!DESCFORINVOICE & "','" & RS!address1 & "','" & RS!address2 & "','" & RS!distcode & "','" & RS!mobile & "',Convert(smalldatetime,'" & RS!invoiceDate & "', 103),'" & RS!GrossAmt & "','" & RS!rate & "','" & RS!scname & "','" & RS!netamount & "','" & RS!discount & "','C','" & RS!fyear & "','" & UserName & "','" & RS!scid & "','addnew','" & txtSponsorshipNo & "')"
                     RS.MoveNext
                    Wend
                  End If

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
    If vs.TextMatrix(vs.RowSel, 1) = "I" Then
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
    
    If db_ <> "no" Then
    
        If vs.TextMatrix(vs.RowSel, 1) = "I" Then
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
           con.Execute "update tmpDonnation set Qty=" & Val(vs.TextMatrix(vs.RowSel, 6)) & " where SNO=" & txtSponsorshipNo & " and invoiceno=" & Val(vs.TextMatrix(vs.RowSel, 2)) & " and BOOKCODE='" & vs.TextMatrix(vs.RowSel, 4) & "' and UserName='" & UserName & "'"
          End If
           
           vs.TextMatrix(vs.RowSel, 8) = Round((Val(vs.TextMatrix(vs.RowSel, 6)) * Val(vs.TextMatrix(vs.RowSel, 7))), 0)
           dis_ = Round((vs.TextMatrix(vs.RowSel, 8) * vs.TextMatrix(vs.RowSel, 9) / 100), 0)
           vs.TextMatrix(vs.RowSel, 10) = Round(Val(vs.TextMatrix(vs.RowSel, 8)) - dis_, 0)
           sendkeys "{right}"
           Total
      ElseIf (vs.Col = 9) Then
      
            dis_ = Round((vs.TextMatrix(vs.RowSel, 8) * vs.TextMatrix(vs.RowSel, 9) / 100), 0)
            vs.TextMatrix(vs.RowSel, 10) = Round(Val(vs.TextMatrix(vs.RowSel, 8)) - dis_, 0)

            If dis_ > 0 Then
            con.Execute "update tmpDonnation set DISCOUNT=" & Val(vs.TextMatrix(vs.RowSel, 9)) & ",net=" & Val(vs.TextMatrix(vs.RowSel, 10)) & " where SNO=" & txtSponsorshipNo & " and invoiceno=" & Val(vs.TextMatrix(vs.RowSel, 2)) & " and BOOKCODE='" & vs.TextMatrix(vs.RowSel, 4) & "' and UserName='" & UserName & "'"
            End If

            
            sendkeys "{down}"
            Total
      
      
     End If
     
  End If
  
Exit Sub

abc1:

MsgBox "" & err.DESCRIPTION
  
End Sub
Private Sub vs_LostFocus()

'Dim sername As String
'
'If rs1.State = 1 Then rs1.close
'rs1.Open "select SerName from BOOKS where bookcode='" & vs.TextMatrix(1, 1) & "'"
'
'
'Set rs1 = con.Execute("exec SearchAppData_SerNameExDis '" & txtScId.Text & "'")
'If rs1.EOF = False Then
'   cboSponse.Text = rs1!Net_Gross
'   txtPercentSp.Text = rs1!Promo
'   txtPercentSp1.Text = rs1!Promo
'End If

End Sub
Private Sub vs2_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 115 Then
  If vs2.TextMatrix(vs2.RowSel, 0) <> "" Then
     con.Execute "delete from tmpAppForDonnation where (SerName='" & vs2.TextMatrix(vs2.RowSel, 0) & "' and appno='" & vs2.TextMatrix(vs2.RowSel, 9) & "')"
     vs2.RemoveItem (vs2.RowSel)
     fillSeries
  End If
  End If
End Sub
