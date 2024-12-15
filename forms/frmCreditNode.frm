VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCreditNode 
   Caption         =   "Credit Node"
   ClientHeight    =   9672
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   15240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9672
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   10950
      TabIndex        =   82
      Top             =   8550
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   11445
      TabIndex        =   38
      Text            =   "0"
      Top             =   8550
      Width           =   1155
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   11445
      TabIndex        =   37
      Text            =   "0"
      Top             =   8250
      Width           =   1155
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   10950
      TabIndex        =   80
      Top             =   8250
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox txtino 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1140
      TabIndex        =   51
      Top             =   120
      Width           =   1245
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   3
      Top             =   120
      Width           =   3510
   End
   Begin VB.TextBox txtCenteral 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "4/2006 CE Dt. 01/03/06 S.No.97"
      Top             =   1800
      Width           =   4125
   End
   Begin VB.TextBox txtAdd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   690
      Left            =   6420
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   49
      Top             =   420
      Width           =   3510
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   6735
      Width           =   1155
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      TabIndex        =   41
      Top             =   7980
      Width           =   8835
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
         Height          =   840
         Left            =   45
         Picture         =   "frmCreditNode.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   135
         Width           =   1230
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
         Height          =   840
         Left            =   1260
         Picture         =   "frmCreditNode.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   1230
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
         Height          =   840
         Left            =   2520
         Picture         =   "frmCreditNode.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   135
         Width           =   1230
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
         Height          =   840
         Left            =   3750
         Picture         =   "frmCreditNode.frx":23AC
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   135
         Width           =   1230
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
         Height          =   840
         Left            =   7500
         Picture         =   "frmCreditNode.frx":27B9
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   120
         Width           =   1230
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
         Height          =   840
         Left            =   6255
         Picture         =   "frmCreditNode.frx":339D
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4980
         Picture         =   "frmCreditNode.frx":3F81
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   120
         Width           =   1230
      End
   End
   Begin VB.TextBox txtDest 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   4
      Top             =   1125
      Width           =   2205
   End
   Begin VB.TextBox txtModePay 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   5
      Text            =   "By Road"
      Top             =   1425
      Width           =   3465
   End
   Begin VB.TextBox txtTrans 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   6
      Top             =   1725
      Width           =   3465
   End
   Begin VB.TextBox txtBoxes 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   7
      Top             =   2025
      Width           =   1245
   End
   Begin VB.TextBox txtRR 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   9
      Top             =   2325
      Width           =   1245
   End
   Begin VB.TextBox txtWagon 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   11
      Top             =   2625
      Width           =   1245
   End
   Begin VB.TextBox txtFrieght 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   12
      Top             =   2925
      Width           =   1245
   End
   Begin VB.TextBox txtWeight 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8760
      TabIndex        =   8
      Top             =   2085
      Width           =   1125
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   11400
      TabIndex        =   40
      Text            =   "0"
      Top             =   7020
      Width           =   1155
   End
   Begin VB.TextBox txtNet 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11445
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   8820
      Width           =   1155
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   10980
      TabIndex        =   35
      Top             =   7020
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   11445
      TabIndex        =   34
      Text            =   "0"
      Top             =   7320
      Width           =   1155
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   10980
      TabIndex        =   33
      Top             =   7320
      Width           =   435
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   2685
      Left            =   10080
      TabIndex        =   20
      Top             =   540
      Width           =   5775
      Begin VB.OptionButton Option_with 
         BackColor       =   &H00FFC0C0&
         Caption         =   "With Form - C"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   29
         Top             =   660
         Width           =   1995
      End
      Begin VB.OptionButton Option_without 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Without Form - C"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   28
         Top             =   1020
         Width           =   1995
      End
      Begin VB.CheckBox Check_Net 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Net Rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   27
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton Option_Check_SaleAgainstFormH 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sale Against Form-H"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   26
         Top             =   1740
         Width           =   2355
      End
      Begin VB.OptionButton Option_NonTaxable 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Non Taxable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   60
         TabIndex        =   25
         Top             =   1380
         Width           =   1995
      End
      Begin VB.ListBox List_SaleSubledger 
         BackColor       =   &H8000000D&
         ForeColor       =   &H00FFFFFF&
         Height          =   1776
         Left            =   2460
         Style           =   1  'Checkbox
         TabIndex        =   24
         Top             =   540
         Width           =   2955
      End
      Begin VB.CheckBox check_AccSet 
         BackColor       =   &H8000000D&
         Caption         =   "Account Setting"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   1755
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Height          =   315
         Left            =   3420
         TabIndex        =   22
         Top             =   2385
         Width           =   855
      End
      Begin VB.ComboBox cboUp_Exup 
         Height          =   315
         ItemData        =   "frmCreditNode.frx":4B65
         Left            =   3600
         List            =   "frmCreditNode.frx":4B72
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label withForm 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   6720
         TabIndex        =   32
         Top             =   2400
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label withoutForm 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   6720
         TabIndex        =   31
         Top             =   2340
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "U.P./Ex. U.P."
         Height          =   255
         Left            =   2460
         TabIndex        =   30
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   11445
      TabIndex        =   36
      Text            =   "0"
      Top             =   7620
      Width           =   1155
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   10980
      TabIndex        =   19
      Top             =   7620
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   10980
      TabIndex        =   18
      Top             =   7920
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   11445
      TabIndex        =   15
      Text            =   "0"
      Top             =   7920
      Width           =   1155
   End
   Begin VB.ComboBox cboPrint 
      Height          =   315
      ItemData        =   "frmCreditNode.frx":4B8B
      Left            =   1260
      List            =   "frmCreditNode.frx":4B98
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   7560
      Width           =   2415
   End
   Begin VB.ComboBox txtOrdered 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7740
      TabIndex        =   13
      Top             =   2925
      Width           =   2175
   End
   Begin Crystal.CrystalReport CR 
      Left            =   13740
      Top             =   6540
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3450
      Left            =   60
      TabIndex        =   14
      Top             =   3255
      Width           =   12585
      _cx             =   22199
      _cy             =   6085
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.4
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
      BackColorSel    =   13888387
      ForeColorSel    =   16711680
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
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
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCreditNode.frx":4BCB
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      Begin VB.Frame VsFrame 
         Height          =   2310
         Left            =   0
         TabIndex        =   52
         Top             =   240
         Visible         =   0   'False
         Width           =   4155
         Begin MSDataListLib.DataCombo cboItem 
            Height          =   2310
            Left            =   0
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   0
            Width           =   4125
            _ExtentX        =   7281
            _ExtentY        =   3958
            _Version        =   393216
            Appearance      =   0
            Style           =   1
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin MSMask.MaskEdBox dateRR 
      Height          =   315
      Left            =   8760
      TabIndex        =   10
      Top             =   2385
      Width           =   1125
      _ExtentX        =   1990
      _ExtentY        =   550
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox dateInv 
      Height          =   315
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   1125
      _ExtentX        =   1990
      _ExtentY        =   550
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox dateIssue 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   660
      Width           =   1125
      _ExtentX        =   1990
      _ExtentY        =   550
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox dateDispatch 
      Height          =   315
      Left            =   3000
      TabIndex        =   2
      Top             =   1020
      Width           =   1125
      _ExtentX        =   1990
      _ExtentY        =   550
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label tax 
      BackColor       =   &H80000013&
      Height          =   255
      Index           =   5
      Left            =   9225
      TabIndex        =   83
      Top             =   8550
      Width           =   1755
   End
   Begin VB.Label tax 
      BackColor       =   &H80000013&
      Height          =   255
      Index           =   4
      Left            =   9225
      TabIndex        =   81
      Top             =   8250
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Credit Note No:"
      Height          =   270
      Index           =   0
      Left            =   90
      TabIndex        =   79
      Top             =   120
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Dealer/Buyer :"
      Height          =   300
      Index           =   2
      Left            =   4440
      TabIndex        =   78
      Top             =   120
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "Date "
      Height          =   270
      Index           =   1
      Left            =   2520
      TabIndex        =   77
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Central Exise Exemption Notification No &&  Date :"
      Height          =   300
      Index           =   4
      Left            =   60
      TabIndex        =   76
      Top             =   1440
      Width           =   4305
   End
   Begin VB.Label Label1 
      Caption         =   "Total :"
      Height          =   195
      Index           =   6
      Left            =   2160
      TabIndex        =   75
      Top             =   6480
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "Date &&  Time of issue of invoice"
      Height          =   270
      Index           =   7
      Left            =   60
      TabIndex        =   74
      Top             =   660
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Date &&  Time of Dispatch"
      Height          =   270
      Index           =   8
      Left            =   60
      TabIndex        =   73
      Top             =   1020
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Destination :"
      Height          =   270
      Index           =   9
      Left            =   4440
      TabIndex        =   72
      Top             =   1125
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "Mode of Transport"
      Height          =   270
      Index           =   10
      Left            =   4440
      TabIndex        =   71
      Top             =   1425
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "Transporter's Name :"
      Height          =   270
      Index           =   11
      Left            =   4440
      TabIndex        =   70
      Top             =   1725
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "No. of Boxes :"
      Height          =   270
      Index           =   12
      Left            =   4440
      TabIndex        =   69
      Top             =   2025
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "RR/LR No. :"
      Height          =   270
      Index           =   13
      Left            =   4440
      TabIndex        =   68
      Top             =   2325
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "Wagon/Truck No. :"
      Height          =   270
      Index           =   14
      Left            =   4440
      TabIndex        =   67
      Top             =   2625
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "Frieght :"
      Height          =   270
      Index           =   15
      Left            =   4440
      TabIndex        =   66
      Top             =   2925
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "Weight No. :"
      Height          =   270
      Index           =   16
      Left            =   7740
      TabIndex        =   65
      Top             =   2085
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "Date "
      Height          =   270
      Index           =   17
      Left            =   7740
      TabIndex        =   64
      Top             =   2445
      Width           =   555
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "Total :"
      Height          =   195
      Left            =   7200
      TabIndex        =   63
      Top             =   6720
      Width           =   795
   End
   Begin VB.Label tax 
      BackColor       =   &H80000013&
      Caption         =   "VAT"
      Height          =   255
      Index           =   0
      Left            =   9225
      TabIndex        =   62
      Top             =   7020
      Width           =   1755
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000013&
      Caption         =   "Net Amount"
      Height          =   255
      Left            =   9225
      TabIndex        =   61
      Top             =   8820
      Width           =   1755
   End
   Begin VB.Label tax 
      BackColor       =   &H80000013&
      Caption         =   "Less Special Dis."
      Height          =   255
      Index           =   1
      Left            =   9225
      TabIndex        =   60
      Top             =   7320
      Width           =   1755
   End
   Begin VB.Label ptype 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   10080
      TabIndex        =   59
      Top             =   120
      Width           =   2715
   End
   Begin VB.Label tax 
      BackColor       =   &H80000013&
      Height          =   255
      Index           =   2
      Left            =   9225
      TabIndex        =   58
      Top             =   7620
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Ordered By :"
      Height          =   270
      Index           =   3
      Left            =   7740
      TabIndex        =   57
      Top             =   2685
      Width           =   1050
   End
   Begin VB.Label tax 
      BackColor       =   &H80000013&
      Height          =   255
      Index           =   3
      Left            =   9225
      TabIndex        =   56
      Top             =   7920
      Width           =   1755
   End
   Begin VB.Label Label4 
      Caption         =   "Print Option"
      Height          =   255
      Left            =   120
      TabIndex        =   55
      Top             =   7620
      Width           =   1155
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Height          =   255
      Left            =   8100
      TabIndex        =   54
      Top             =   6720
      Width           =   1095
   End
End
Attribute VB_Name = "frmCreditNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itemgp As String
Dim rates As Double
Dim I As Integer
Dim Status As String
Dim Item_Name As String

Dim unit As String
Dim qty As Integer

Dim iitem1 As String
Dim StockFlag As String

Dim Edit As Boolean
Const gl As String = "SUNDRY DEBTORS"

Const VAT As Double = 4.5
Const VAT1 As Double = 2

Dim total1 As Double
Dim VAT_less As Double

Dim VAT_Add As Double
Dim Add_Amt, Less_Amt As Double

Dim ST As String
Dim nontax As String

Dim header As String
Dim gledger() As String

Dim Debit_Credit() As String
Dim taxhead As String

Dim Sale_SLedger As String
Dim rss As New ADODB.Recordset


Const dis1 As Double = 25
Const dis2 As Double = 9.09
Const dis3 As Double = 7.5
Const dis4 As Double = 5



Private Sub cmdMain_Click()
Unload Me
End Sub
Sub Total()

On Error Resume Next

txtTotal.Text = 0

Dim f1, f2 As Double

total1 = 0
f1 = 0
f2 = 0


For J = 1 To vs.rows - 1
If vs.TextMatrix(J, 1) <> "" Then
txtTotal.Text = (Val(txtTotal.Text) + Val(vs.TextMatrix(J, 6)))
End If
Next

'-------------------------------------------------------------
txtTotal = Format(txtTotal, ".00")
total1 = (Val(txtTotal) - Val(txtAmount(0)))


'-------------------------------------------------------------
txtAmount(1) = Round((total1 * Val(txtRate(1)) / 100), 2)
txtAmount(1) = Format(txtAmount(1), ".00")

'-------------------------------------------------------------
Add_Amt = Val(txtAmount(2))
Less_Amt = Val(txtAmount(3))

f1 = Val(txtAmount(4))
f2 = Val(txtAmount(5))


'-------------------------------------------------------------
txtNet = Format((Val(total1) + Val(txtAmount(1))), ".00")

txtNet = (Val(txtNet) + Add_Amt)
txtNet = (Val(txtNet) - Less_Amt)
txtNet = (Val(txtNet) + f1)
txtNet = (Val(txtNet) - f2)



txtNet = Format((txtNet), ".00")


End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdModify_Click()
End Sub
Private Sub cmdref_Click()
      txtHeating.Text = ""
      txtParty.Text = ""
      
      txtRemarks.Text = ""
      
      
      txtTotal1.Text = 0
      txtTotal2.Text = 0
      txtTotal3.Text = 0
      txtTotal4.Text = 0
      
      txtSize.Text = ""
      txtGrade.Text = ""
      txtRawAndCasting.Text = 0
      
      vs.Clear
      vs1.Clear
      vs2.Clear
      vs3.Clear
      
      setWidth
      txtHeating.SetFocus
      cmdDelete.Enabled = False
      cmdModify.Enabled = False
      cmdSave.Enabled = True
      
      Record = ""
      
End Sub


Private Sub Command4_Click()
   Unload Me
End Sub

Private Sub cboaddLess_Change()

End Sub

Private Sub check_AccSet_Click()
  If check_AccSet.value = 0 Then
     Frame1.Width = ((5775 / 2) - 430)
  Else
     Frame1.Width = 5775
  End If
End Sub

Private Sub cmdAdd_1_Click()
   
   
Dim o As Object
For Each o In Me
If TypeOf o Is textbox Then
o.Text = ""
End If

If TypeOf o Is MaskEdBox Then
o.Text = "__/__/____"
End If

Next


dateRR.Text = Format(Date, "dd/MM/yyyy")

'Check_Nontaxable.Value = False
Option_with.value = True
Option_NonTaxable.value = False
txtParty.Enabled = True
ptype.Caption = ""

dateInv = Format(Date, "dd/MM/yyyy")

'cboaddLess.ListIndex = 0

dateInv.SetFocus
txtRate(0) = VAT
txtCenteral = "4/2006 CE Dt. 01/03/06 S.No.97"
txtModePay = "BY Road"
   
   vs.Clear
   setWidth
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = True
   'cmdPrint_7.Enabled = False
   cmdSave_2.Enabled = True
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
   txtino.Text = MaxSNo_Tax("credita", "INVOICENO")
End Sub
Function MaxSNo_Tax(tbl As String, fld As String) As Double
    Dim RS As New Recordset
    If RS.State = 1 Then RS.close
    RS.Open "Select max(" & fld & ") from " & tbl & " where typeofinvoice = 'tax' and " & stringyear & "", con
    If IsNull(RS(0)) Then
        MaxSNo_Tax = 1
    Else
        MaxSNo_Tax = Val(RS(0)) + 1
    End If
    RS.close
End Function


Private Sub cmdDelete_3_Click()

  If checkAuthentication("credita", "INVOICENO", txtino) = True Then
      MsgBox "You are Not Authorised ...", vbCritical
      Exit Sub
   End If
   


If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
   
   con.BeginTrans
   con.Execute "delete from credita where " & stringyear & " invoiceNo =" & txtino & ""
   con.Execute "delete from creditb where " & stringyear & " invoiceNo =" & txtino & ""
   con.Execute "delete from creditc where " & stringyear & " invoiceNo =" & txtino & ""
   con.CommitTrans
   
   Call cmdAdd_1_Click
End If
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False
End Sub

Private Sub cmdEdit_4_Click()
   
   If checkAuthentication("credita", "INVOICENO", txtino) = True Then
      MsgBox "You are Not Authorised ...", vbCritical
      Exit Sub
   End If
   
   cmdDelete_3.Enabled = True
   cmdEdit_4.Enabled = False
   cmdSave_2.Enabled = True
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
   Edit = True
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub
Private Sub cmdPrint_7_Click()

DSNNew

cr.Reset
cr.ReportFileName = rptPath & "/CHALLAN.rpt"
cr.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
cr.ReplaceSelectionFormula "{credita.invoiceno}=" & txtHeating.Text & ""
cr.WindowShowPrintSetupBtn = True
cr.WindowState = crptMaximized
cr.Action = 1

End Sub


''Sub searchData()
''
''If rs.State = 1 Then rs.Close
''rs.Open "select * from credita where " & stringyear & " INVOICENO=" & txtHeating.Text & "", CON, adOpenDynamic, adLockOptimistic
''If rs.EOF = False Then
''txtParty.Text = rs.Fields("SUBLEDGER").Value
''txtRemarks.Text = rs.Fields("Remarks").Value & ""
''Text1.Text = rs.Fields("add1").Value & ""
''Text2.Text = rs.Fields("add2").Value & ""
''If Not IsNull(rs.Fields("godown").Value) Then
''cboGodown = rs.Fields("godown").Value & ""
''Else
''cboGodown.ListIndex = -1
''End If
''
''End If
''
''
''
''If rs.State = 1 Then rs.Close
''rs.Open "select * from creditb where " & stringyear & " INVOICENO=" & txtHeating.Text & "", CON, adOpenDynamic, adLockOptimistic
''For i = 1 To rs.RecordCount
''If rs.EOF = False Then
''vs.TextMatrix(i, 0) = rs.Fields("BOOKCODE").Value
''vs.TextMatrix(i, 1) = rs.Fields("TBook").Value
''vs.TextMatrix(i, 2) = rs.Fields("LoosBook").Value
''vs.TextMatrix(i, 3) = rs.Fields("TotalBook").Value
''vs.TextMatrix(i, 4) = rs.Fields("NetBook").Value
''vs.TextMatrix(i, 5) = rs.Fields("remarks").Value & ""
''vs.TextMatrix(i, 6) = rs.Fields("Book_Code").Value & ""
''rs.MoveNext
''End If
''Next
''
''Total
''
''End Sub
Private Sub cmdUndo_5_Click()
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = False
   cmdPrint_7.Enabled = True
   cmdSave_2.Enabled = False
   cmdUndo_5.Enabled = False
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
End Sub



Private Sub dates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtParty.SetFocus
End Sub

Private Sub cmdok_Click()
    saveDataForSaleSubledger
End Sub
Sub saveDataForSaleSubledger()
    Dim s As String
    
  con.Execute "delete from InvoiceSubLedger where " & stringyear & " TaxHead='" & taxhead & "' and Up_Exup='" & cboUp_Exup.Text & "'"
    
  For I = 0 To List_SaleSubledger.ListCount - 1
    
  If List_SaleSubledger.Selected(I) = True Then
  
    If RS.State = 1 Then RS.close
    s = "SELECT [TaxHead],[SubLedger] from [ExportData].[dbo].[InvoiceSubLedger] where " & stringyear & " [TaxHead]='" & taxhead & "' and [SubLedger]='" & List_SaleSubledger.List(I) & "'"
    RS.Open s, con
    If RS.EOF = True Then
       con.Execute "insert into InvoiceSubLedger(TaxHead,SubLedger,Up_Exup) values('" & taxhead & "','" & List_SaleSubledger.List(I) & "','" & cboUp_Exup.Text & "')"
    End If
    
  End If
    
  Next
    
  MsgBox "Data Saved ...", vbInformation
  
End Sub
Sub SearchDataForSaleSubledger()
  Dim s As String
   
  
   If ST = "" Then
      MsgBox "Plz. Select Dealer/Buyer ...", vbCritical
      Exit Sub
   End If
    
   If (Option_without.value = True Or Option_with.value = True) Then
    If ST = "U.P." Then
       cboUp_Exup.Text = "U.P."
    Else
       cboUp_Exup.Text = "Ex. U.P."
    End If
   
   ElseIf (Option_NonTaxable.value = True Or Option_Check_SaleAgainstFormH = True) Then
      cboUp_Exup.Text = "Non"
   End If
    
    
    
    For I = 0 To List_SaleSubledger.ListCount - 1
      List_SaleSubledger.Selected(I) = False
    Next
  
  
  
   
  If rs1.State = 1 Then rs1.close
  s = "SELECT [TaxHead],[SubLedger],up_exup from [ExportData].[dbo].[InvoiceSubLedger] where " & stringyear & " [TaxHead]='" & taxhead & "' and up_exup='" & cboUp_Exup.Text & "' order by TaxHead"
  rs1.Open s, con
  If rs1.EOF = False Then
   
    cboUp_Exup.Text = rs1(2)
  
    For I = 0 To List_SaleSubledger.ListCount - 1
      If rs1(1) = List_SaleSubledger.List(I) Then
         List_SaleSubledger.Selected(I) = True
         Sale_SLedger = List_SaleSubledger.List(I)
      End If
    Next
  
  End If
 
 
End Sub

Private Sub cmdPrint_Click()
    
On Error GoTo aa1:
    
    DSNNew
    
    cr.Reset
    cr.Connect = constr
    cr.ReportFileName = rptPath & "\Exinvoice_ret.rpt"
    cr.ReplaceSelectionFormula "{invoicea.invoiceno} = " & txtino & " AND {invoicea.setupid} = " & main.setupid & " AND {invoicea.fyear} = '" & main.session & "'"
    
    cr.Formulas(0) = "tax_label='" & tax(0) & "'"
    cr.Formulas(1) = "tax_rate='" & txtRate(0) & "'"
    cr.Formulas(2) = "tax_amount='" & txtAmount(0) & "'"

    cr.Formulas(3) = "tax_label1='" & tax(1) & "'"
    cr.Formulas(4) = "tax_rate1='" & txtRate(1) & "'"
    cr.Formulas(5) = "tax_amount1='" & Format(txtAmount(1), ".00") & "'"
    
    cr.Formulas(6) = "header='" & "CREDIT NOTE" & "'"
    
    If RS.State = 1 Then RS.close
    RS.Open "select tpt,contphone,tinno from [SLEDGER] where " & stringyear & " subledger='" & txtParty & "'", con
    If RS.EOF = False Then
      cr.Formulas(7) = "cst='" & RS(0) & "'"
      cr.Formulas(8) = "uptt='" & RS(1) & "'"
      cr.Formulas(9) = "tin='" & RS(2) & "'"
    End If
    
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "SELECT [State] from [invoiceQuery] where " & stringyear & " [SUBLEDGER]='" & txtParty & "'", con
    If rs1.EOF = False Then
    If Option_with.value = True Then
       cr.Formulas(10) = "WITH_WITHOUT_FORM='" & Option_with.Caption & "'"
    ElseIf Option_without.value = True Then
       cr.Formulas(10) = "WITH_WITHOUT_FORM='" & Option_without.Caption & "'"
    ElseIf Option_NonTaxable.value = True Then
       cr.Formulas(10) = "WITH_WITHOUT_FORM='" & Option_NonTaxable.Caption & "'"
    ElseIf Option_Check_SaleAgainstFormH.value = True Then
       cr.Formulas(10) = "WITH_WITHOUT_FORM='" & Option_Check_SaleAgainstFormH.Caption & "'"
    End If
    End If
   
    
    cr.Formulas(11) = "add_lesshead='" & cboaddLess & "'"
    cr.Formulas(12) = "add_lessAmt='" & txtAmount(2) & "'"
    cr.Formulas(13) = "p_status='" & ptype & "'"
    cr.Formulas(14) = "copyinv='" & cboPrint.Text & "'"
    cr.Formulas(15) = "toword='" & toword(txtNet.Text) & "'"
    
    
    
    cr.WindowState = crptMaximized
    
   'If MsgBox("Want To Print ?", vbQuestion + vbYesNo) = vbYes Then
   '   CR.Destination = crptToPrinter
   'Else
       cr.Action = 1
   'End If

Exit Sub
aa1:
MsgBox "" & err.DESCRIPTION


End Sub

Private Sub cmdSave_2_Click()


'On Error GoTo save:

Dim n As Date
Dim I As Integer
Dim netrate As String
Dim with_without As String


If Check_Net.value = 1 Then
   netrate = "y"
   Else
   netrate = "n"
End If

'If Option_with.Value = True Then
'with_without = 1
'Else
'with_without = 2
'End If



If Option_with.value = True Then
     with_without = 1
ElseIf Option_without.value = True Then
     with_without = 2
ElseIf Option_NonTaxable.value = True Then
     with_without = 3
ElseIf Option_Check_SaleAgainstFormH.value = True Then
     with_without = 4
End If
  



If ptype.Caption = "" Then
  MsgBox "Plz. Select Party Name ...", vbCritical
  Exit Sub
End If

If Not IsDate(dateIssue.Text) Then
  MsgBox "Enter Issue Date ...", vbCritical
  dateIssue.SetFocus
  Exit Sub
End If

If Not IsDate(dateDispatch.Text) Then
  MsgBox "Enter Dispatch Date ...", vbCritical
  dateDispatch.SetFocus
  Exit Sub
End If


If Not IsDate(dateRR.Text) Then
  MsgBox "Enter RR/LR Date ...", vbCritical
  dateRR.SetFocus
  Exit Sub
End If

Dim nontax As String

'If Check_Nontaxable.Value = 0 Then
If Option_NonTaxable.value = True Then
    nontax = "nontax"
ElseIf Option_Check_SaleAgainstFormH.value = True Then
    nontax = "formh"
End If

Dim amt1 As Double
Dim amt2 As Double

amt1 = Val(txtAmount(4))
amt2 = Val(txtAmount(5))

I = 1

If Edit = False Then

            txtino = MaxSNo_Tax("credita", "INVOICENO")
            con.BeginTrans
            
            con.Execute "exec insertData_credita " & txtino & ",'" & dateInv & "','" & dateIssue & "'," & _
            "'" & dateDispatch & "','" & txtCenteral.Text & "','SUNDRY DEBTORS','" & txtParty & "','" & txtDest & "'," & _
            "'" & txtModePay & "','" & txtTrans & "','" & txtBoxes & "','" & txtWeight & "','" & txtRR & "','" & dateRR.Text & "'," & _
            "'" & txtFrieght & "','" & txtWagon & "','" & Val(txtTotal) & "','" & Val(txtNet) & "','" & with_without & "','" & netrate & "','tax','" & main.UserName & "','" & main.UserName & "','" & main.session & "'," & _
            "" & main.setupid & ""
            
            
            
            con.Execute "update credita  set NonTaxable = '" & nontax & "',orderby='" & IIf(txtOrdered = "", "Direct Office", txtOrdered) & "'," & _
            " aexp1='" & tax(0) & "',aexp1rate=" & txtRate(0) & ",aexp1am='" & txtAmount(0) & "'," & _
            " aexp2='" & tax(1) & "',aexp2rate=" & txtRate(1) & ",aexp2am='" & txtAmount(1) & "'," & _
            " aexp3='" & tax(2) & "',aexp3rate=" & txtRate(2) & ",aexp3am='" & txtAmount(2) & "'," & _
            " aexp4='" & tax(3) & "',aexp4rate=" & txtRate(3) & ",aexp4am='" & txtAmount(3) & "'," & _
            " aexp5='" & (tax(4)) & "',aexp5rate=" & Val(txtRate(4)) & ",aexp5am='" & amt1 & "'," & _
            " aexp6='" & (tax(5)) & "',aexp6rate=" & Val(txtRate(5)) & ",aexp6am='" & amt2 & "'" & _
            " where invoiceno=" & txtino & " and where " & stringyear & ""
            
             
             
            For I = 1 To vs.rows - 1
            
            If vs.TextMatrix(I, 1) <> "" Then
            
            con.Execute "exec insertData_creditb " & txtino & ",'" & dateInv & "','" & txtParty & "'," & _
            "" & Val(vs.TextMatrix(I, 0)) & ",'" & vs.TextMatrix(I, 1) & "'," & Val(vs.TextMatrix(I, 3)) & "," & _
            "" & Val(vs.TextMatrix(I, 4)) & "," & Val(vs.TextMatrix(I, 5)) & "," & Val(vs.TextMatrix(I, 6)) & "," & _
            "'tax','" & main.UserName & "','" & main.UserName & "','" & main.session & "','" & vs.TextMatrix(I, 2) & "'," & _
            "" & main.setupid & ""
            
            End If
            
            Next
            
            For J = 0 To txtAmount.Count - 1
            
                con.Execute "insert into creditc" & _
                "(INVOICENO,INVOICEDate,GENLEDGER,Subledger,GAmount,Rate,Amount,typeofinvoice,DebitOrCredit," & _
                "text,fyear,createdby,createdon,updatedby,updatedon,setupid) values(" & _
                "" & txtino & ",'" & Format(dateInv, "MM/dd/yyyy") & "','" & gledger(J) & "','" & Sale_SLedger & "'," & Val(txtNet) & "," & Val(txtRate(J)) & "," & _
                "" & Val(txtAmount(J)) & ",'tax','" & Debit_Credit(J) & "','" & tax(J) & "','" & main.session & "','" & main.UserName & "'," & _
                "'" & Format(Date, "MM/DD/yyyy") & "','" & main.UserName & "','" & Format(Date, "MM/dd/yyyy") & "'," & main.setupid & ")"
            
            Next
            
            
            con.CommitTrans




Else

            '-------------------------------------------
            
            
            
            con.BeginTrans
            
            
            con.Execute "delete from credita where " & stringyear & " invoiceNo =" & txtino & ""
            con.Execute "delete from creditb where " & stringyear & " invoiceNo =" & txtino & ""
            con.Execute "delete from creditc where " & stringyear & " invoiceNo =" & txtino & ""
            
            
            
            con.Execute "exec insertData_credita " & txtino & ",'" & dateInv & "','" & dateIssue & "'," & _
            "'" & dateDispatch & "','" & txtCenteral.Text & "','SUNDRY DEBTORS','" & txtParty & "','" & txtDest & "'," & _
            "'" & txtModePay & "','" & txtTrans & "','" & txtBoxes & "','" & txtWeight & "','" & txtRR & "','" & dateRR.Text & "'," & _
            "'" & txtFrieght & "','" & txtWagon & "','" & Val(txtTotal) & "','" & Val(txtNet) & "','" & with_without & "','" & netrate & "','tax','" & main.UserName & "','" & main.UserName & "','" & main.session & "'," & _
            "" & main.setupid & ""
            
            'CON.Execute "update credita  set NonTaxable = '" & nontax & "',orderby='" & IIf(txtOrdered = "", "Direct Office", txtOrdered) & "' where invoiceno=" & txtino & ""
            
            con.Execute "update credita  set NonTaxable = '" & nontax & "',orderby='" & IIf(txtOrdered = "", "Direct Office", txtOrdered) & "'," & _
            " aexp1='" & tax(0) & "',aexp1rate=" & txtRate(0) & ",aexp1am='" & txtAmount(0) & "'," & _
            " aexp2='" & tax(1) & "',aexp2rate=" & txtRate(1) & ",aexp2am='" & txtAmount(1) & "'," & _
            " aexp3='" & tax(2) & "',aexp3rate=" & txtRate(2) & ",aexp3am='" & txtAmount(2) & "'," & _
            " aexp4='" & tax(3) & "',aexp4rate=" & txtRate(3) & ",aexp4am='" & txtAmount(3) & "'," & _
            " aexp5='" & (tax(4)) & "',aexp5rate=" & Val(txtRate(4)) & ",aexp5am='" & amt1 & "'," & _
            " aexp6='" & (tax(5)) & "',aexp6rate=" & Val(txtRate(5)) & ",aexp6am='" & amt2 & "'" & _
            " where invoiceno=" & txtino & " and " & stringyear & ""
             
            
            For I = 1 To vs.rows - 1
            
            If vs.TextMatrix(I, 1) <> "" Then
            
            con.Execute "exec insertData_creditb " & txtino & ",'" & dateInv & "','" & txtParty & "'," & _
            "" & Val(vs.TextMatrix(I, 0)) & ",'" & vs.TextMatrix(I, 1) & "'," & Val(vs.TextMatrix(I, 3)) & "," & _
            "" & Val(vs.TextMatrix(I, 4)) & "," & Val(vs.TextMatrix(I, 5)) & "," & Val(vs.TextMatrix(I, 6)) & "," & _
            "'tax','" & main.UserName & "','" & main.UserName & "','" & main.session & "','" & vs.TextMatrix(I, 2) & "'," & _
            "" & main.setupid & ""
            
            End If
            
            Next
            
            
            
            For J = 0 To txtAmount.Count - 1
            
            If tax(J) <> "-" Then
            
                 
                con.Execute "insert into creditc" & _
                "(INVOICENO,INVOICEDate,GENLEDGER,Subledger,GAmount,Rate,Amount,typeofinvoice,DebitOrCredit," & _
                "text,fyear,createdby,createdon,updatedby,updatedon,setupid) values(" & _
                "" & txtino & ",'" & Format(dateInv, "MM/dd/yyyy") & "','" & gledger(J) & "','" & Sale_SLedger & "'," & Val(txtNet) & "," & Val(txtRate(J)) & "," & _
                "" & Val(txtAmount(J)) & ",'tax','" & Debit_Credit(J) & "','" & tax(J) & "','" & main.session & "','" & main.UserName & "'," & _
                "'" & Format(Date, "MM/DD/yyyy") & "','" & main.UserName & "','" & Format(Date, "MM/dd/yyyy") & "'," & main.setupid & ")"
    
                
            
            End If
            
            
            Next
            
            
            
            
            con.CommitTrans
            
            
            Edit = False

'---------------------------------------------

End If


cmdEdit_4.Enabled = True
cmdDelete_3.Enabled = True
cmdSave_2.Enabled = False
'Call cmdAdd_1_Click



''Exit Sub
''
''
''save:
''
''CON.RollbackTrans
'''If err.Number = "-2147217900" Then
''   MsgBox "" & err.DESCRIPTION, vbCritical
''   'txtCode.SetFocus
'''End If




End Sub

Private Sub cmdSearch_Click()
popuplist10 "select [INVOICENO], [INVOICEDATE], [SUBLEDGER] from credita where " & stringyear & " order by cast(INVOICENO as int)", con
End Sub

Private Sub cmdSearch_GotFocus()
txtino = PopUpValue1
searchData
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
End Sub

Private Sub dateInv_LostFocus()
'dateIssue.Text = dateInv.Text
'dateDispatch.Text = dateInv.Text
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

 
If KeyCode = 27 Then
     If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
        Unload Me
     End If
 End If
 
 If KeyCode = 13 Then
 
 If UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("vs")) Then
    SendKeys "{tab}"
    HIT
 End If
 
 End If
 
 
 
 
 
End Sub
Sub searchData()

vs.Clear
setWidth


If RS.State = 1 Then RS.close

st1 = "select INVOICENO,INVOICEDATE,IssueDate,DisPatchDate,CentralExise,SUBLEDGER,STATION," & _
"ModeOfPayment,THROUGH,BUNDLES,WEIGHT,BILTYNO,BILTYDate,FREIGHT,TXT1,[with_withoutFormc],[NetRate]," & _
"NonTaxable,Orderby,gamount,netAmount,netrate from credita where invoiceno=" & txtino & " and  " & stringyear & ""
     
     
On Error Resume Next



     
RS.Open st1, con

If RS.EOF = True Then
   Exit Sub
End If

If RS.EOF = False Then
  
  
  
  txtParty.Enabled = False
  
  dateInv = RS!invoiceDate
  dateIssue = Format(RS!IssueDate, "dd/MM/yyyy")
  dateDispatch = Format(RS!DisPatchDate, "dd/MM/yyyy")
  txtCenteral = RS!CentralExise
  txtParty = RS!subledger
  txtDest = RS!station
  txtModePay = RS!ModeOfPayment
  txtTrans = RS!through
  txtBoxes = RS!bundles
  txtWeight = RS!weight
  dateRR = RS!BILTYDATE
  txtRR = RS!biltyno
  txtWagon = RS!txt1
  txtFrieght = RS!freight
  txtOrdered = RS!orderby & ""
  
  If RS![with_withoutFormc] = 1 Then
     Option_with.value = True
  ElseIf RS![with_withoutFormc] = 2 Then
     Option_without.value = True
  ElseIf RS![with_withoutFormc] = 3 Then
     Option_NonTaxable.value = True
  ElseIf RS![with_withoutFormc] = 4 Then
     Option_Check_SaleAgainstFormH.value = True
  End If
  
  
  
  
  If RS![netrate] = "y" Then
     Check_Net.value = 1
  Else
     Check_Net.value = 0
  End If
  
  
  
  txtTotal = Format(RS!gamount, "0.00")
  txtNet = Format(RS!netamount, "0.00")
  
  
  
  If RS.State = 1 Then RS.close
   RS.Open "SELECT [SUBLEDGER],address1,address2,[City],[District],[State],[bcityid],typeofcust " & _
   "from [ExportData].[dbo].[SubledgerQry] WHERE SUBLEDGER ='" & txtParty & "' and  " & stringyear & "", con
   If RS.EOF = False Then
     txtadd = RS![subledger] & " " & vbCrLf & RS!address1 & " " & RS!address1 & vbCrLf & RS![city] + vbCrLf + RS![District] & "," & RS![State]
     Me.ptype.Caption = RS!TypeOfCust
   End If
  
End If

lblQty = 0

If RS.State = 1 Then RS.close
RS.Open "select * from creditb where INVOICENO=" & txtino.Text & " and  " & stringyear & " order by printorder", con, adOpenDynamic, adLockOptimistic
For I = 1 To RS.RecordCount
If RS.EOF = False Then
   
    vs.TextMatrix(I, 0) = RS.Fields("printorder").value
    vs.TextMatrix(I, 1) = RS.Fields("BOOKCODE").value
    
    vs.TextMatrix(I, 3) = RS.Fields("Quantity").value
    vs.TextMatrix(I, 4) = RS.Fields("Rate").value
    vs.TextMatrix(I, 5) = Format(RS.Fields("NetRate").value, ".00")
    vs.TextMatrix(I, 6) = Format(RS.Fields("Amount").value, ".00")
    
    lblQty = (Val(lblQty) + RS.Fields("Quantity").value)
    
    If rs1.State = 1 Then rs1.close
    rs1.Open "select ProductQuality,TypeofProduct,rulling,rate,NoofPages from copymaster " & _
    "where " & stringyear & " bookno='" & vs.TextMatrix(I, 1) & "'", con                     ' and  " & stringyear & "", CON
    If rs1.EOF = False Then
       vs.TextMatrix(I, 2) = rs1!TypeofProduct + " (" + rs1!rulling + ")" + Str(rs1!NoOfPages) + " " + rs1!ProductQuality
    End If
    
RS.MoveNext
End If
Next


'Total

            
  

If RS.State = 1 Then RS.close
RS.Open "select [aexp1],[aexp1rate],[aexp1am],[aexp2],[aexp2rate],[aexp2am]," & _
"[aexp3],[aexp3rate],[aexp3am],[aexp4],[aexp4rate],[aexp4am] from credita " & _
" where INVOICENO=" & txtino.Text & " and  " & stringyear & "", con, adOpenDynamic, adLockOptimistic
If RS.EOF = False Then
   
   
   tax(0) = RS![aexp1]
   txtRate(0) = RS![aexp1rate]
   txtAmount(0) = Format(RS![aexp1am], "0.00")
   
   tax(1) = RS![aexp2]
   txtRate(1) = RS![aexp2rate]
   txtAmount(1) = Format(RS![aexp2am], "0.00")
   
   
   tax(2) = RS![aexp3]
   txtRate(2) = RS![aexp3rate]
   txtAmount(2) = Format(RS![aexp3am], "0.00")
   
   tax(3) = RS![aexp4]
   txtRate(3) = RS![aexp4rate]
   txtAmount(3) = Format(RS![aexp4am], "0.00")
     
   cmdSave_2.Enabled = False

End If






thead = ""

If Option_NonTaxable.value = True Then
   thead = Option_NonTaxable.Caption
ElseIf Option_Check_SaleAgainstFormH.value = True Then
   thead = Option_Check_SaleAgainstFormH.Caption
ElseIf Option_with.value = True Then
   thead = Option_with.Caption
ElseIf Option_without.value = True Then
   thead = Option_without.Caption
   
End If


If rs1.State = 1 Then rs1.close
rs1.Open "select subledger from InvoiceSubLedger where " & stringyear & " taxhead='" & thead & "'", con
If rs1.EOF = False Then
   Sale_SLedger = rs1!subledger
End If



'-----------------------

If RS.State = 1 Then RS.close
RS.Open "select State,tinno from SubledgerQry where " & stringyear & " subledger='" & txtParty & "'", con
If RS.EOF = False Then
    ST = RS(0)
    If LCase(ST) = "u.p." Then
     If Len(RS!tinno) > 0 Then
       header = "TAX-INVOICE"
       Sale_SLedger = "Tax Invoice Sales"
     Else
       header = "SALE-INVOICE"
       Sale_SLedger = "Sales Invoice"
     End If
       
    Else
       header = "SALE-INVOICE"
    '   Sale_SLedger = "Sales Invoice"
    End If
End If

''=================================




k1 = 0
ReDim gledger(txtAmount.Count)
ReDim Debit_Credit(txtAmount.Count)

If RS.State = 1 Then RS.close
RS.Open "select text,Rate,GENLEDGER,DEBITORCREDIT,amount from creditc " & _
" where  " & stringyear & " and invoiceNo=" & txtino & " order by auto", con
For kk1 = 1 To txtAmount.Count + 1
    If RS.EOF = False Then
       tax(k1) = RS(0)
       txtRate(k1) = RS(1)
       gledger(k1) = RS!Genledger
       Debit_Credit(k1) = RS!DebitorCredit
       txtAmount(k1) = RS!amount
       k1 = k1 + 1
       RS.MoveNext
    End If
Next





End Sub

''Sub TotalFinal()
''   If txtTotal3.Text = "" Then
''      txtTotal3.Text = 0
''   End If
''
''   If txtTotal2.Text = "" Then
''      txtTotal2.Text = 0
''   End If
''
''
''    txtRawAndCasting.Text = (CDbl(txtTotal2.Text) + CDbl(txtTotal3.Text))
''    txtRawAndCasting.Text = Format(txtRawAndCasting.Text, "#,###.000")
''End Sub
Private Sub Form_Load()
 
 setWidth
 
 dateInv.Text = Format(Date, "dd/MM/yyyy")
 
 txtino = MaxSNo_Tax("credita", "INVOICENO")
 
 'txtRate(0) = VAT
 
 dateRR.Text = Format(Date, "dd/MM/yyyy")
 
 
 
 withForm = VAT1
 withoutForm = VAT
 
 Frame1.Width = ((5775 / 2) - 430)
 
 
 
If RS.State = 1 Then RS.close
RS.Open "select subledger from sledger where " & stringyear & " and gledger='SALES'", con
While RS.EOF = False
  List_SaleSubledger.AddItem RS(0)
  RS.MoveNext
Wend
 
cboPrint.ListIndex = 0

If RS.State = 1 Then RS.close
RS.Open "select distinct agentname from agentmaster where " & stringyear & "", con, adOpenKeyset, adLockReadOnly
While RS.EOF = False
  txtOrdered.AddItem RS(0)
  RS.MoveNext
Wend
 
If main.UserName = "admin" Then
   check_AccSet.Enabled = True
Else
  check_AccSet.Enabled = False
End If
 
 
ButtonPermission cmdSave_2, cmdDelete_3, cmdEdit_4
 
End Sub
Sub setWidth()
vs.Cols = 7
vs.FormatString = "S.No.|^Item Code|<Item Name|>Quantity|>MRP|>Net Rate|>Net Amount"
vs.ColWidth(0) = 300
vs.ColWidth(1) = 1000
vs.ColWidth(2) = 6500
vs.ColWidth(3) = 1000
vs.ColWidth(4) = 800
vs.ColWidth(5) = 1000
vs.ColWidth(6) = 1200
End Sub
''Private Sub fromdate_KeyDown(KeyCode As Integer, Shift As Integer)
''   If KeyCode = 13 Then todate.SetFocus
''End Sub
''Private Sub HeatingDate_KeyDown(KeyCode As Integer, Shift As Integer)
''   If KeyCode = 13 Then txtParty.SetFocus
''End Sub

Private Sub Option_Check_SaleAgainstFormH_Click()
 
 taxhead = Option_Check_SaleAgainstFormH.Caption
' SearchDataForSaleSubledger

End Sub

Private Sub Option_NonTaxable_Click()
taxhead = Option_NonTaxable.Caption
'SearchDataForSaleSubledger

End Sub

''Private Sub ListHeatingNo_Click()
''  Call cmdref_Click
''  SearchData
''  TotalFinal
''  'Frame1.Visible = False
''End Sub
'Private Sub txtGrade_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Then txtRemarks.SetFocus
'End Sub

''Private Sub txtHeating_GotFocus()
''If PopUpValue1 <> "" Then
''txtHeating.Text = PopUpValue1
''Dates.Value = PopUpValue2
''vs.Clear
''setwidth
''SearchData
''PopUpValue1 = ""
''PopUpValue2 = ""
''End If
''End Sub

''Private Sub txtHeating_KeyDown(KeyCode As Integer, Shift As Integer)
'''If KeyCode = 113 Then
'''popuplist2 "select INVOICENO,INVOICEDATE,SUBLEDGER from credita order by INVOICENO", CON
'''End If
'''
'''If KeyCode = 13 Then
'''SearchData
'''End If
''
''End Sub

''Private Sub txtHeating_KeyPress(KeyAscii As Integer)
''
''   If KeyAscii = 13 Then
''
''   Dates.SetFocus
''
''
''  End If
''
''
''End Sub

Private Sub Option_with_Click()
'If Option_with.Value = True Then
'  txtRate(0) = withForm
'Else
'  txtRate(0) = withoutForm
'End If

'FatchTaxFromSate

taxhead = Option_with.Caption


End Sub

Private Sub Option_without_Click()
'If Option_with.Value = True Then
'  txtRate(0) = withForm
'Else
'  txtRate(0) = withoutForm
'End If

taxhead = Option_without.Caption


'FatchTaxFromSate

'SearchDataForSaleSubledger
End Sub

Private Sub txtamount_GotFocus(Index As Integer)
HIT
End Sub

Private Sub txtamount_LostFocus(Index As Integer)

txtAmount(Index) = Format(txtAmount(Index), ".00")
Total


End Sub

Private Sub txtino_GotFocus()
HIT
End Sub

Private Sub txtino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  searchData
  
End If
End Sub

Private Sub txtParty_GotFocus()
HIT
If PopUpValue1 <> "" Then
   txtParty = PopUpValue2
   
   
  If RS.State = 1 Then RS.close
   RS.Open "SELECT [SUBLEDGER],address1,address2,[City],[District],[State],[bcityid],typeofcust " & _
   "from [ExportData].[dbo].[SubledgerQry] where " & stringyear & " SUBLEDGER ='" & txtParty & "'", con
   If RS.EOF = False Then
     txtadd = RS![subledger] & " " & vbCrLf & RS!address1 & " " & RS!address1 & vbCrLf & RS![city] + vbCrLf + RS![District] & "," & RS![State]
     Me.ptype.Caption = RS!TypeOfCust
   End If
   
   
   
   FatchTaxFromSate
   
   txtDest.SetFocus
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   
End If

End Sub

Private Sub txtParty_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Exit Sub
tblNo = 1
headData = "SUNDRY DEBTORS"
frmSearchItem.Show
End Sub

Private Sub txtParty_LostFocus()
   Record = ""
End Sub
Private Sub txtQty_GotFocus()
     txtQty.SelLength = 10
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub
Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys "{tab}"
End Sub
Private Sub txtSize_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtGrade.SetFocus
End Sub

Private Sub txtRate_GotFocus(Index As Integer)
 HIT
End Sub

Private Sub txtRate_LostFocus(Index As Integer)
   
   s = InStr(tax(1).Caption, "@")
   
   tax(1).Caption = Mid(tax(1).Caption, 1, s) & txtRate(1) & "%"
End Sub

Private Sub vs_KeyDown(KeyCode As Integer, Shift As Integer)

  
  
  If KeyCode = 115 Then
  If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
    vs.RemoveItem (vs.RowSel)
    Total
  End If
  End If
  

  
  
  
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)

Dim Item As String


If KeyCode = 13 Then
        
 If vs.Col = 1 Then
    
    
 If vs.TextMatrix(vs.RowSel, 1) = "" Then
    Exit Sub
 End If
    
    
    If RS.State = 1 Then RS.close
    RS.Open "select ProductQuality,TypeofProduct,rulling,rate,NoofPages from copymaster " & _
    "where " & stringyear & " bookno='" & vs.TextMatrix(vs.RowSel, 1) & "'", con
    If RS.EOF = False Then
          
          Item = RS!TypeofProduct + " (" + RS!rulling + ")" + Str(RS!NoOfPages) + " " + RS!ProductQuality
       
          vs.TextMatrix(vs.RowSel, 2) = Item
          
          vs.TextMatrix(vs.RowSel, 4) = RS.Fields("Rate").value
          
    End If
    
    SendKeys "{right}"
    SendKeys "{right}"
 
 End If
 
 
 If Check_Net.value = 1 Then
 
 
 If vs.Col = 3 Then
    vs.TextMatrix(vs.RowSel, 6) = Format((Val(vs.TextMatrix(vs.RowSel, 3)) * Val(vs.TextMatrix(vs.RowSel, 5))), ".00")
    
    SendKeys "{right}"
    SendKeys "{right}"
    Total
 End If
 
 If vs.Col = 5 Then
   
    vs.TextMatrix(vs.RowSel, 6) = Format((Val(vs.TextMatrix(vs.RowSel, 3)) * Val(vs.TextMatrix(vs.RowSel, 5))), ".00")
    SendKeys "{home}"
    SendKeys "{down}"
    
    vs.TextMatrix(vs.RowSel, 0) = vs.Row
    
    Total
 
 End If



Else
 
 
 
 
 If vs.Col = 3 Then
   
    
    
    Dim amt1, amt2, amt3, amt4 As Double
    
    amt1 = (Val(vs.TextMatrix(vs.RowSel, 4)) - ((Val(vs.TextMatrix(vs.RowSel, 4)) * dis1) / 100))
    amt2 = (amt1 - ((amt1 * dis2) / 100))
    amt3 = (amt2 - ((amt2 * dis3) / 100))
     
    If ptype.Caption = "Supper Stockist" Then
       amt4 = (amt3 - ((amt3 * dis4) / 100))
    Else
       amt4 = amt3
    End If
    
    ss = 0
     
    
    FatchTaxFromSate
    If Option_NonTaxable.value = False Then
    'If Check_Nontaxable.Value = 0 Then
        If Option_with.value = True Then
           ss = ((amt4 * VAT_less) / 100)
           amt4 = Round((amt4 - ss), 2)
        Else
           ss = ((amt4 * VAT_less) / 100)
           amt4 = Round((amt4 - ss), 2)
        End If
    End If
     
     
     
    vs.TextMatrix(vs.RowSel, 5) = Format(Round(amt4, 2), ".00")
    vs.TextMatrix(vs.RowSel, 6) = (Val(vs.TextMatrix(vs.RowSel, 3)) * Val(vs.TextMatrix(vs.RowSel, 5)))
    
    vs.TextMatrix(vs.RowSel, 6) = Format(vs.TextMatrix(vs.RowSel, 6), ".00")
    
    vs.TextMatrix(vs.RowSel, 0) = vs.Row
    SendKeys "{home}"
    SendKeys "{down}"

    
    Total
 End If
    
    
    
 End If


       





End If

End Sub

Sub FatchTaxFromSate()

On Error Resume Next


Dim with_without As String
Dim B_CH  As Boolean

B_CH = False

If Option_with.value = True Then
   with_without = Option_with.Caption
Else
   with_without = Option_without.Caption
End If




If rs1.State = 1 Then rs1.close
rs1.Open "select [State],tinno from SubledgerQry where subledger='" & txtParty & "' and  " & stringyear & "", con
If rs1.EOF = False Then
   ST = rs1(0)
   B_CH = True
End If






If rs1.State = 1 Then rs1.close
rs1.Open "select add_val,less_val from [state_tax_list] where " & stringyear & " statename='" & ST & "'" & _
" and with_without='" & with_without & "'", con
If rs1.EOF = False Then
   VAT_Add = rs1(0)
   VAT_less = rs1(1)
   
   con.Execute "update CREDITEND set ryn='y'  where  " & stringyear
   
If (Option_NonTaxable.value = True Or Option_Check_SaleAgainstFormH.value = True) Then
   con.Execute "update CREDITEND set ryn='n'  where [Rate] >0 and  " & stringyear & ""
Else
   con.Execute "update CREDITEND set ryn='n'  where  ([Rate] <> " & VAT_Add & " and [Rate] >0) and  " & stringyear & ""
End If

End If






If Option_NonTaxable.value = True Then
   txtRate(0) = 0
End If


k1 = 0

If (Option_NonTaxable.value = True Or Option_Check_SaleAgainstFormH.value = True) Then
    B_CH = True

    If Option_NonTaxable = True Then
       nontax = Option_NonTaxable.Caption
    Else
       nontax = Option_Check_SaleAgainstFormH.Caption
    End If
Else
   B_CH = False
End If

'============================================================
kk1 = 0

If rs1.State = 1 Then rs1.close
rs1.Open "select printorder from CREDITEND where  " & stringyear & " and RYN='y' order by printorder", con
While rs1.EOF = False
  con.Execute "update CREDITEND set id=" & kk1 & "  where  printorder=" & rs1(0) & " and " & stringyear
  kk1 = kk1 + 1
  rs1.MoveNext
Wend

'============================================================



k1 = 0

ReDim gledger(txtAmount.Count)
ReDim Debit_Credit(txtAmount.Count)


For kk1 = 0 To txtAmount.Count - 1

If rs1.State = 1 Then rs1.close
rs1.Open "select text,Rate,GENLEDGER,DEBITORCREDIT from CREDITEND where  " & stringyear & " and id=" & kk1 & " and RYN='y' order by printorder", con
If rs1.EOF = False Then
   tax(k1) = rs1(0)
   txtRate(k1) = rs1(1)
   gledger(k1) = rs1!Genledger
   Debit_Credit(k1) = rs1!DebitorCredit
   k1 = k1 + 1


If B_CH = True Then
   tax(k1) = nontax
   txtRate(k1) = 0
   gledger(k1) = "-"
   Debit_Credit(k1) = "-"
   B_CH = False
   k1 = k1 + 1
End If


End If


Next


SearchDataForSaleSubledger




End Sub



