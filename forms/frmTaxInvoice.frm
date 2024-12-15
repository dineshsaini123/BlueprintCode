VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTaxInvoice 
   ClientHeight    =   9780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15960
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   15960
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtDays 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7740
      TabIndex        =   101
      Text            =   "30"
      Top             =   3120
      Width           =   1245
   End
   Begin VB.CommandButton cmdCopy 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Copy Invoice"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   360
      Width           =   1275
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   11400
      TabIndex        =   93
      Text            =   "0"
      Top             =   8505
      Width           =   1155
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   10950
      TabIndex        =   92
      Top             =   8505
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   11400
      TabIndex        =   90
      Text            =   "0"
      Top             =   8190
      Width           =   1155
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   10950
      TabIndex        =   89
      Top             =   8190
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtParty1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4455
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   1890
   End
   Begin VB.TextBox txtFormNo 
      Height          =   285
      Left            =   1200
      TabIndex        =   81
      Top             =   2715
      Width           =   2940
   End
   Begin VB.TextBox txtPermitNo 
      Height          =   285
      Left            =   1200
      TabIndex        =   80
      Top             =   2400
      Width           =   2940
   End
   Begin VB.TextBox txtPaymentRem 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   83
      Top             =   3180
      Width           =   5895
   End
   Begin VB.ComboBox cboWarehouse 
      Height          =   315
      Left            =   1290
      Style           =   2  'Dropdown List
      TabIndex        =   79
      Top             =   2025
      Width           =   2940
   End
   Begin VB.ComboBox txtOrdered 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7740
      TabIndex        =   78
      Top             =   2760
      Width           =   2175
   End
   Begin VB.ComboBox cboPrint 
      Height          =   315
      ItemData        =   "frmTaxInvoice.frx":0000
      Left            =   1260
      List            =   "frmTaxInvoice.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   75
      Top             =   7500
      Width           =   2415
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   11400
      TabIndex        =   19
      Text            =   "0"
      Top             =   7905
      Width           =   1155
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   10935
      TabIndex        =   69
      Top             =   7905
      Width           =   480
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   10935
      TabIndex        =   66
      Top             =   7290
      Width           =   480
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   11400
      TabIndex        =   18
      Text            =   "0"
      Top             =   7290
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   2895
      Left            =   10080
      TabIndex        =   56
      Top             =   480
      Width           =   5775
      Begin VB.CheckBox Check1_nontaxable 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Non Taxable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   98
         Top             =   2520
         Width           =   2295
      End
      Begin VB.ComboBox cboUp_Exup 
         Height          =   315
         ItemData        =   "frmTaxInvoice.frx":0040
         Left            =   3600
         List            =   "frmTaxInvoice.frx":004D
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   120
         Width           =   1155
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Height          =   315
         Left            =   4560
         TabIndex        =   72
         Top             =   2460
         Width           =   855
      End
      Begin VB.CheckBox check_AccSet 
         BackColor       =   &H8000000D&
         Caption         =   "Account Setting"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
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
         TabIndex        =   71
         Top             =   120
         Width           =   1755
      End
      Begin VB.ListBox List_SaleSubledger 
         BackColor       =   &H8000000D&
         ForeColor       =   &H00FFFFFF&
         Height          =   1860
         Left            =   2460
         Style           =   1  'Checkbox
         TabIndex        =   70
         Top             =   540
         Width           =   2955
      End
      Begin VB.OptionButton Option_NonTaxable 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Non Taxable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   59
         Top             =   1380
         Width           =   1995
      End
      Begin VB.OptionButton Option_Check_SaleAgainstFormH 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sale Against Form-H"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   60
         Top             =   1740
         Width           =   2355
      End
      Begin VB.CheckBox Check_Net 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Net Rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   61
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton Option_without 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Without Form - C"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   58
         Top             =   1020
         Width           =   1995
      End
      Begin VB.OptionButton Option_with 
         BackColor       =   &H00FFC0C0&
         Caption         =   "With Form - C"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   57
         Top             =   660
         Width           =   1995
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "U.P./Ex. U.P."
         Height          =   255
         Left            =   2460
         TabIndex        =   74
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label withoutForm 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   6720
         TabIndex        =   63
         Top             =   2340
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label withForm 
         BackColor       =   &H00C0FFC0&
         Height          =   255
         Left            =   6720
         TabIndex        =   62
         Top             =   2400
         Visible         =   0   'False
         Width           =   315
      End
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   10935
      TabIndex        =   54
      Top             =   6990
      Width           =   480
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   11400
      TabIndex        =   17
      Text            =   "0"
      Top             =   6990
      Width           =   1155
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   10935
      TabIndex        =   52
      Top             =   6690
      Width           =   480
   End
   Begin VB.TextBox txtNet 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11400
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   8835
      Width           =   1155
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   11400
      TabIndex        =   16
      Text            =   "0"
      Top             =   6690
      Width           =   1155
   End
   Begin VB.TextBox txtWeight 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8760
      TabIndex        =   10
      Top             =   1980
      Width           =   1125
   End
   Begin VB.TextBox txtFrieght 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   14
      Top             =   2820
      Width           =   1245
   End
   Begin VB.TextBox txtWagon 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   13
      Top             =   2520
      Width           =   1245
   End
   Begin VB.TextBox txtRR 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   11
      Top             =   2220
      Width           =   1245
   End
   Begin VB.TextBox txtBoxes 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   9
      Top             =   1920
      Width           =   1245
   End
   Begin VB.TextBox txtTrans 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   8
      Top             =   1620
      Width           =   3465
   End
   Begin VB.TextBox txtModePay 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   7
      Text            =   "By Road"
      Top             =   1320
      Width           =   3465
   End
   Begin VB.TextBox txtDest 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6420
      TabIndex        =   6
      Top             =   1020
      Width           =   2205
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   0
      TabIndex        =   34
      Top             =   7920
      Width           =   8835
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4980
         Picture         =   "frmTaxInvoice.frx":0066
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   120
         Width           =   1230
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   6255
         Picture         =   "frmTaxInvoice.frx":0C4A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   7500
         Picture         =   "frmTaxInvoice.frx":182E
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   120
         Width           =   1230
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   3750
         Picture         =   "frmTaxInvoice.frx":2412
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdDelete_3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2520
         Picture         =   "frmTaxInvoice.frx":281F
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   1260
         Picture         =   "frmTaxInvoice.frx":3403
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   1230
      End
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   45
         Picture         =   "frmTaxInvoice.frx":3FE7
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11385
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   6390
      Width           =   1155
   End
   Begin Crystal.CrystalReport CR 
      Left            =   12780
      Top             =   3540
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtAdd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   645
      Left            =   6420
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   30
      Top             =   360
      Width           =   3510
   End
   Begin VB.TextBox txtCenteral 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "16/2012 CE  Dt. 17/03/2012"
      Top             =   1725
      Width           =   4125
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6435
      TabIndex        =   4
      Top             =   45
      Width           =   3465
   End
   Begin VB.TextBox txtino 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1140
      TabIndex        =   24
      Top             =   45
      Width           =   1305
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   2925
      Left            =   60
      TabIndex        =   15
      Top             =   3480
      Width           =   12585
      _cx             =   22199
      _cy             =   5159
      _ConvInfo       =   1
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      FormatString    =   $"frmTaxInvoice.frx":4BCB
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
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   4155
         Begin MSDataListLib.DataCombo cboItem 
            Height          =   2310
            Left            =   0
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   0
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   4075
            _Version        =   393216
            Appearance      =   0
            Style           =   1
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
      TabIndex        =   12
      Top             =   2280
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
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
      Top             =   60
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
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
      Top             =   840
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
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
      Top             =   1140
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
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
   Begin VB.Label Label11 
      Caption         =   "F4 Delete Grid Raw ....."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3900
      TabIndex        =   102
      Top             =   7620
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "No Of Days :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   6420
      TabIndex        =   100
      Top             =   3180
      Width           =   1290
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000013&
      Caption         =   "Total :"
      Height          =   240
      Left            =   8865
      TabIndex        =   97
      Top             =   6435
      Width           =   2130
   End
   Begin VB.Label lblGTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10890
      TabIndex        =   96
      Top             =   7605
      Width           =   1650
   End
   Begin VB.Label tax 
      BackColor       =   &H80000013&
      Caption         =   "Gross Total :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   8865
      TabIndex        =   95
      Top             =   7605
      Width           =   2055
   End
   Begin VB.Label tax 
      BackColor       =   &H80000013&
      Height          =   255
      Index           =   5
      Left            =   8865
      TabIndex        =   94
      Top             =   8505
      Width           =   2100
   End
   Begin VB.Label tax 
      BackColor       =   &H80000013&
      Height          =   255
      Index           =   4
      Left            =   8865
      TabIndex        =   91
      Top             =   8190
      Width           =   2100
   End
   Begin VB.Label lblinv 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3780
      TabIndex        =   88
      Top             =   6795
      Width           =   3030
   End
   Begin VB.Label lbltin 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   90
      TabIndex        =   87
      Top             =   6795
      Width           =   3570
   End
   Begin VB.Label Label9 
      Caption         =   "Form- C No."
      Height          =   240
      Left            =   120
      TabIndex        =   86
      Top             =   2715
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Road Permit No."
      Height          =   195
      Left            =   75
      TabIndex        =   85
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Payemt Remarks :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   84
      Top             =   2970
      Width           =   2595
   End
   Begin VB.Label Label5 
      Caption         =   "Warehouse "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   82
      Top             =   2055
      Width           =   1590
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      Height          =   255
      Left            =   7635
      TabIndex        =   77
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Print Option"
      Height          =   255
      Left            =   180
      TabIndex        =   76
      Top             =   7560
      Width           =   1155
   End
   Begin VB.Label tax 
      BackColor       =   &H80000013&
      Height          =   255
      Index           =   3
      Left            =   8850
      TabIndex        =   68
      Top             =   7905
      Width           =   2100
   End
   Begin VB.Label Label1 
      Caption         =   "Ordered By :"
      Height          =   270
      Index           =   3
      Left            =   7740
      TabIndex        =   67
      Top             =   2580
      Width           =   1050
   End
   Begin VB.Label tax 
      BackColor       =   &H80000013&
      Height          =   255
      Index           =   2
      Left            =   8850
      TabIndex        =   65
      Top             =   7290
      Width           =   2100
   End
   Begin VB.Label ptype 
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   10080
      TabIndex        =   64
      Top             =   60
      Width           =   2715
   End
   Begin VB.Label tax 
      BackColor       =   &H80000013&
      Height          =   255
      Index           =   1
      Left            =   8850
      TabIndex        =   53
      Top             =   6990
      Width           =   2100
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000013&
      Caption         =   "Net Amount"
      Height          =   255
      Left            =   8850
      TabIndex        =   51
      Top             =   8835
      Width           =   2100
   End
   Begin VB.Label tax 
      BackColor       =   &H80000013&
      Height          =   255
      Index           =   0
      Left            =   8850
      TabIndex        =   50
      Top             =   6690
      Width           =   2100
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Caption         =   "Total :"
      Height          =   255
      Left            =   6915
      TabIndex        =   49
      Top             =   6840
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Date "
      Height          =   270
      Index           =   17
      Left            =   7740
      TabIndex        =   47
      Top             =   2340
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Weight No. :"
      Height          =   270
      Index           =   16
      Left            =   7740
      TabIndex        =   46
      Top             =   1980
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "Frieght :"
      Height          =   210
      Index           =   15
      Left            =   4440
      TabIndex        =   45
      Top             =   2940
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "Wagon/Truck No. :"
      Height          =   210
      Index           =   14
      Left            =   4440
      TabIndex        =   44
      Top             =   2580
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "RR/LR No. :"
      Height          =   210
      Index           =   13
      Left            =   4440
      TabIndex        =   43
      Top             =   2280
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "No. of Boxes :"
      Height          =   210
      Index           =   12
      Left            =   4440
      TabIndex        =   42
      Top             =   1980
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "Transporter's Name :"
      Height          =   210
      Index           =   11
      Left            =   4440
      TabIndex        =   41
      Top             =   1680
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "Mode of Transport"
      Height          =   210
      Index           =   10
      Left            =   4440
      TabIndex        =   40
      Top             =   1380
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "Destination :"
      Height          =   210
      Index           =   9
      Left            =   4440
      TabIndex        =   39
      Top             =   1080
      Width           =   1770
   End
   Begin VB.Label Label1 
      Caption         =   "Date &&  Time of Dispatch"
      Height          =   270
      Index           =   8
      Left            =   60
      TabIndex        =   38
      Top             =   1140
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Date &&  Time of issue of invoice"
      Height          =   270
      Index           =   7
      Left            =   60
      TabIndex        =   37
      Top             =   840
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Total :"
      Height          =   195
      Index           =   6
      Left            =   2160
      TabIndex        =   29
      Top             =   6420
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "Central Exise Exemption Notification No &&  Date :"
      Height          =   300
      Index           =   4
      Left            =   60
      TabIndex        =   28
      Top             =   1455
      Width           =   4305
   End
   Begin VB.Label Label1 
      Caption         =   "Date "
      Height          =   270
      Index           =   1
      Left            =   2520
      TabIndex        =   27
      Top             =   60
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Dealer/Buyer :"
      Height          =   300
      Index           =   2
      Left            =   4440
      TabIndex        =   26
      Top             =   60
      Width           =   2205
   End
   Begin VB.Label Label1 
      Caption         =   "Invoice No:"
      Height          =   270
      Index           =   0
      Left            =   90
      TabIndex        =   25
      Top             =   60
      Width           =   1110
   End
End
Attribute VB_Name = "frmTaxInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itemgp As String
Dim rates As Double
Dim i As Integer
Dim Status As String
Dim Item_Name As String
Dim unit As String
Dim qty As Integer
Dim iitem1 As String
Dim StockFlag As String
Dim edit As Boolean
Const gl As String = "SUNDRY DEBTORS"
Const VAT As Double = 4.5
Const VAT1 As Double = 2

Dim total1 As Double
Dim VAT_less As Double
Dim VAT_Add As Double

Dim copy_inv As Boolean

Dim Add_Amt, Less_Amt As Double
Dim ST As String

Dim nontax As String
Dim sledger1, gledger1 As String

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
txtNet = 0


total1 = 0



For J = 1 To vs.Rows - 1
If vs.TextMatrix(J, 1) <> "" Then
txtTotal.Text = (Val(txtTotal.Text) + Val(vs.TextMatrix(J, 6)))
End If
Next

'-------------------------------------------------------------
txtTotal = Format(txtTotal, ".00")
total1 = (Val(txtTotal) - Val(txtamount(0)))


subtot = (Val(txtTotal.Text) + Val(txtamount(0)) + Val(txtamount(1)) + Val(txtamount(2)))
lblGTotal.Caption = Format(subtot, ".00")


'-------------------------------------------------------------
txtamount(1) = Round(((total1 * Val(txtRate(1)) / 100) + 0.0001), 2)
txtamount(1) = Format(txtamount(1), ".00")

'-------------------------------------------------------------
Add_Amt = Val(txtamount(4))
Less_Amt = Val(txtamount(5))

'''-------------------------------------------------------------
txtNet = (Val(txtamount(3)) + subtot)


txtNet = (Val(txtNet) + Add_Amt)
txtNet = (Val(txtNet) + Less_Amt)


txtNet = Format((txtNet), ".00")


End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdModify_Click()
End Sub
Private Sub cmdRef_Click()
      txtHeating.Text = ""
      txtParty1.Text = ""
      
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
      
      setwidth
      txtHeating.SetFocus
      cmdDelete.Enabled = False
      cmdmodify.Enabled = False
      cmdSave.Enabled = True
      
      Record = ""
      
End Sub


Private Sub Command4_Click()
   Unload Me
End Sub

Private Sub cboaddLess_Change()

End Sub

Private Sub check_AccSet_Click()
  If check_AccSet.Value = 0 Then
     Frame1.Width = ((5775 / 2) - 430)
  Else
     Frame1.Width = 5775
  End If
End Sub

Private Sub Check_Net_Click()
 If Check_Net.Value = 1 Then
 '   Check1_nontaxable.Value = 0
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


copy_inv = False

lblinv.Caption = ""
Check1_nontaxable.Value = 0

dateRR.Text = Format(Date, "dd/MM/yyyy")

lbltin.Caption = ""

'Check_Nontaxable.Value = False
'Option_with.Value = True
'Option_NonTaxable.Value = False
txtParty1.Enabled = True
ptype.Caption = ""

dateInv = Format(Date, "dd/MM/yyyy")

'cboaddLess.ListIndex = 0

dateInv.SetFocus
txtRate(0) = VAT
txtCenteral = "4/2006 CE Dt. 01/03/06 S.No.97"
txtModePay = "BY Road"
   
   vs.Clear
   setwidth
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = True
   'cmdPrint_7.Enabled = False
   cmdSave_2.Enabled = True
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
   txtino.Text = MaxSNo_Tax("invoicea", "INVOICENO")




Option_Check_SaleAgainstFormH.Value = False
Option_NonTaxable.Value = False
Option_with.Value = False
Option_without.Value = False

Check_Net.Value = 0


End Sub
Function MaxSNo_Tax(tbl As String, fld As String) As Double
    Dim rs As New Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "Select max(" & fld & ") from " & tbl & " where " & stringyear & " and  typeofinvoice = 'tax'", CON
    If IsNull(rs(0)) Then
        MaxSNo_Tax = 1
    Else
        MaxSNo_Tax = Val(rs(0)) + 1
    End If
    rs.Close
End Function


Private Sub cmdCopy_Click()
   
   Dim inv
   
   inv = ""
   
   inv = InputBox("Enter Invoice No :", "Message")
   
   If inv <> "" Then
      
      copy_inv = True
      txtino = inv
      Call cmdSave_2_Click
      
      copy_inv = True
      
   Else
      Exit Sub
   End If
   
End Sub

Private Sub cmdDelete_3_Click()

  If checkAuthentication("invoicea", "INVOICENO", txtino) = True Then
      MsgBox "You are Not Authorised ...", vbCritical
      Exit Sub
   End If
   


If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
   
   CON.BeginTrans
   CON.Execute "delete from invoicea where " & stringyear & " and  invoiceNo =" & txtino & ""
   CON.Execute "delete from invoiceb where " & stringyear & " and  invoiceNo =" & txtino & ""
   CON.Execute "delete from invoicec where " & stringyear & " and  invoiceNo =" & txtino & ""
   CON.CommitTrans
   
   Call cmdAdd_1_Click
End If
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False
End Sub

Private Sub cmdEdit_4_Click()
   
   If checkAuthentication("invoicea", "INVOICENO", txtino) = True Then
      MsgBox "You are Not Authorised ...", vbCritical
      Exit Sub
   End If
   
   cmdDelete_3.Enabled = True
   cmdEdit_4.Enabled = False
   cmdSave_2.Enabled = True
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
   edit = True
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdPrint_7_Click()

cr.Reset
cr.ReportFileName = App.Path & "/Reports/CHALLAN.rpt"
cr.ReplaceSelectionFormula "{invoiceA.invoiceno}=" & txtHeating.Text & ""
cr.WindowShowPrintSetupBtn = True
cr.WindowState = crptMaximized
cr.Action = 1

End Sub


''Sub searchData()
''
''If rs.State = 1 Then rs.Close
''rs.Open "select * from invoicea where " & stringyear & " and  INVOICENO=" & txtHeating.Text & "", CON, adOpenDynamic, adLockOptimistic
''If rs.EOF = False Then
''txtParty1.Text = rs.Fields("SUBLEDGER").Value
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
''rs.Open "select * from invoiceb where " & stringyear & " and  INVOICENO=" & txtHeating.Text & "", CON, adOpenDynamic, adLockOptimistic
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



Private Sub Dates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then txtParty1.SetFocus
End Sub

Private Sub cmdOK_Click()
    saveDataForSaleSubledger
End Sub
Sub saveDataForSaleSubledger()
    Dim s As String
    
  CON.Execute "delete from InvoiceSubLedger where " & stringyear & " and  TaxHead='" & taxhead & "' and Up_Exup='" & cboUp_Exup.Text & "'"
    
  For i = 0 To List_SaleSubledger.ListCount - 1
    
  If List_SaleSubledger.Selected(i) = True Then
  
    If rs.State = 1 Then rs.Close
    s = "SELECT [TaxHead],[SubLedger] from [ExportData].[dbo].[InvoiceSubLedger] where " & stringyear & " and  [TaxHead]='" & taxhead & "' and [SubLedger]='" & List_SaleSubledger.List(i) & "'"
    rs.Open s, CON
    If rs.EOF = True Then
       CON.Execute "insert into InvoiceSubLedger(TaxHead,SubLedger,Up_Exup) values('" & taxhead & "','" & List_SaleSubledger.List(i) & "','" & cboUp_Exup.Text & "')"
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
    
   If (Option_without.Value = True Or Option_with.Value = True) Then
    If ST = "U.P." Then
       cboUp_Exup.Text = "U.P."
    Else
       cboUp_Exup.Text = "Ex. U.P."
    End If
   
   ElseIf (Option_NonTaxable.Value = True Or Option_Check_SaleAgainstFormH = True) Then
      cboUp_Exup.Text = "Non"
   End If
    
    
    
    For i = 0 To List_SaleSubledger.ListCount - 1
      List_SaleSubledger.Selected(i) = False
    Next
  
  
  
   
  If rs1.State = 1 Then rs1.Close
  s = "SELECT [TaxHead],[SubLedger],up_exup from [ExportData].[dbo].[InvoiceSubLedger] where " & stringyear & " and  [TaxHead]='" & taxhead & "' and up_exup='" & cboUp_Exup.Text & "' order by TaxHead"
  rs1.Open s, CON
  If rs1.EOF = False Then
   
    cboUp_Exup.Text = rs1(2)
  
    For i = 0 To List_SaleSubledger.ListCount - 1
      If rs1(1) = List_SaleSubledger.List(i) Then
         List_SaleSubledger.Selected(i) = True
         Sale_SLedger = List_SaleSubledger.List(i)
      End If
    Next
  
  End If
  
 
  
 
 
End Sub
Sub showCat()

Dim K1 As Integer
Dim subtot As Double

K1 = 0
subtot = 0

'------------------------------------------------------
If rs.State = 1 Then rs.Close
'rs.Open "select head,rate from ExsieDetail order by orderid", CON
rs.Open "select TEXT,rate from INVOICEEND WHERE (TEXT LIKE 'CENVAT%' OR TEXT LIKE 'EDU%' OR TEXT LIKE 'S.H.EDU%') and " & stringyear & "  order by PRINTORDER", CON
If rs.EOF = False Then
While rs.EOF = False

   tax(K1) = rs!Text
   
   
  If Check1_nontaxable.Value = 0 Then
     txtRate(K1) = rs!rate
     
  Else
     txtRate(K1) = 0
     txtamount(K1) = 0
  End If
   
   K1 = K1 + 1
   
rs.MoveNext

Wend
End If

'---cenvat duty--------------------------------------------------

If Check1_nontaxable.Value = 1 Then
   txtRate(3) = 0
   txtamount(3) = 0
End If


txtRate(4) = 0
txtRate(5) = 0


txtamount(0).Text = Round(((Val(txtTotal.Text) * Val(txtRate(0)) / 100) + 0.0001), 0)

If txtRate(0).Text <> 0 Then
If txtamount(0).Text < 1 Then txtamount(0).Text = 1
End If

'''txtamount(1).Text = Round((Val(txtTotal.Text) * Val(txtRate(1)) / 100), 0)
txtamount(1).Text = Round(((Val(txtamount(0).Text) * Val(txtRate(1)) / 100) + 0.0001), 0)


If txtRate(1).Text <> 0 Then
If txtamount(1).Text < 1 Then txtamount(1).Text = 1
End If

txtamount(2).Text = Round(((Val(txtamount(0).Text) * Val(txtRate(2)) / 100) + 0.0001), 0)

If txtRate(2).Text <> 0 Then
If txtamount(2).Text < 1 Then txtamount(2).Text = 1
End If


subtot = (Val(txtTotal.Text) + Val(txtamount(0)) + Val(txtamount(1)) + Val(txtamount(2)))
lblGTotal.Caption = Format(subtot, ".00")

txtamount(3).Text = Round((Val(subtot) * Val(txtRate(3)) / 100), 2)

txtNet.Text = ((Val(txtTotal.Text) + Val(txtamount(0).Text) + Val(txtamount(1).Text) + Val(txtamount(2).Text) + Val(txtamount(3).Text) + Val(txtamount(4).Text)) - Val(txtamount(5).Text))


'------------------------------------------------------

If rs.State = 1 Then rs.Close
rs.Open "select State,tinno from SubledgerQry where " & stringyear & " and  subledger='" & txtParty1 & "'", CON
If rs.EOF = False Then


If Len(rs!tinno) > 0 Then
   lblinv.Caption = "TAX INVOICE"
   lbltin.Caption = "Tin No : " & rs!tinno
Else
   lblinv.Caption = "SALE INVOICE"
   lbltin.Caption = ""
End If


If ST = "U.P." Then

    If Len(rs!tinno) = 0 Then
       Sale_SLedger = "sales Invoice"
     Else
       Sale_SLedger = "Tax Invoice Sales"
    End If
    
   If Option_NonTaxable.Value = True Then
      Sale_SLedger = "sales Non - Taxable"
   End If

Else



If Option_with.Value = True Then
   Sale_SLedger = "Sales Agst. Form-C"
ElseIf Option_without.Value = True Then
   Sale_SLedger = "Sales Without Form -C"
ElseIf Option_NonTaxable.Value = True Then
   Sale_SLedger = "sales Non - Taxable"
ElseIf Option_Check_SaleAgainstFormH.Value = True Then
   Sale_SLedger = "sales Agst.Form - H"
End If


End If
End If
  
  


End Sub
Private Sub cmdPrint_Click()
    cr.Reset
    
    cr.Connect = constr
    cr.ReportFileName = strrptpath & "\REPORTS\Exinvoice_Excise.rpt"
    cr.ReplaceSelectionFormula "{invoicea.invoiceno} = " & txtino & " AND {invoicea.setupid} = " & main.setupid & " AND {invoicea.fyear} = '" & main.session & "'"
    
    cr.Formulas(0) = "tax_label='" & tax(0) & "'"
    cr.Formulas(1) = "tax_rate='" & txtRate(0) & "'"
    cr.Formulas(2) = "tax_amount='" & txtamount(0) & "'"

    cr.Formulas(3) = "tax_label1='" & tax(1) & "'"
    cr.Formulas(4) = "tax_rate1='" & txtRate(1) & "'"
    cr.Formulas(5) = "tax_amount1='" & Format(txtamount(1), ".00") & "'"
    
    
    '-----------------
    '-----------------------
    
    If rs.State = 1 Then rs.Close
    rs.Open "select State,tinno from SubledgerQry where " & stringyear & " and  subledger='" & txtParty1 & "'", CON
    If rs.EOF = False Then
    ST = rs(0)
    'PopUpValue3 = rs!tinno
    
    If LCase(ST) = "u.p." Then
     If Len(rs!tinno) > 0 Then
       header = "TAX-INVOICE"
       Sale_SLedger = "Tax Invoice Sales"
     Else
       header = "SALE-INVOICE"
       Sale_SLedger = "Sales Invoice"
     End If
       
    Else
       header = "SALE-INVOICE"
     End If
    End If
    
    ''=================================

    If Check1_nontaxable.Value = 1 Then
       header = "SALE-INVOICE"
    End If
    
    
    
    cr.Formulas(6) = "header='" & header & "'"
    
    If rs.State = 1 Then rs.Close
    rs.Open "select tpt,contphone,tinno from [SLEDGER] where " & stringyear & " and  subledger='" & txtParty1 & "'", CON
    If rs.EOF = False Then
      cr.Formulas(7) = "cst='" & rs(0) & "'"
      cr.Formulas(8) = "uptt='" & rs(1) & "'"
      cr.Formulas(9) = "tin='" & rs(2) & "'"
    End If
    
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "SELECT [State] from [invoiceQuery] where " & stringyear & " and  [SUBLEDGER]='" & txtParty1 & "'", CON
    If rs1.EOF = False Then
    If Option_with.Value = True Then
       cr.Formulas(10) = "WITH_WITHOUT_FORM='Against Form - C'"
    ElseIf Option_without.Value = True Then
       cr.Formulas(10) = "WITH_WITHOUT_FORM='" & Option_without.Caption & "'"
    ElseIf Option_NonTaxable.Value = True Then
       cr.Formulas(10) = "WITH_WITHOUT_FORM='" & Option_NonTaxable.Caption & "'"
    ElseIf Option_Check_SaleAgainstFormH.Value = True Then
       cr.Formulas(10) = "WITH_WITHOUT_FORM='" & Option_Check_SaleAgainstFormH.Caption & "'"
    End If
    End If
    
    
    cr.Formulas(11) = "add_lesshead='" & cboaddLess & "'"
    cr.Formulas(12) = "add_lessAmt='" & txtamount(2) & "'"
    cr.Formulas(13) = "p_status='" & ptype & "'"
    cr.Formulas(14) = "copyinv='" & cboPrint.Text & "'"
    cr.Formulas(15) = "toword='" & toword(txtNet.Text) & "'"
    'sum1 = (Val(txtamount(0)) + Val(txtamount(1)) + Val(txtamount(2)))
    
    cr.Formulas(16) = "toword1='" & toword((Val(txtamount(0)) + Val(txtamount(1)) + Val(txtamount(2)))) & "'"
    
    'If CR.PageCount > 1 Then
    '   CR.Formulas(19) = "Pagecontinue='Pagecontinue'"
    'End If

    
    
    cr.WindowState = crptMaximized
    
   'If MsgBox("Want To Print ?", vbQuestion + vbYesNo) = vbYes Then
   'CR.Destination = crptToPrinter
   'Else
    cr.Action = 1
    
    
   'End If


End Sub

Private Sub cmdSave_2_Click()


On Error GoTo save:

Dim n As Date
Dim i As Integer
Dim netrate As String
Dim with_without As String

Dim non_taxable As String

non_taxable = ""










If Check_Net.Value = 1 Then
   netrate = "y"
   non_taxable = "0"
   Else
   netrate = "n"
End If




If Option_with.Value = True Then
     with_without = 1
ElseIf Option_without.Value = True Then
     with_without = 2
ElseIf Option_NonTaxable.Value = True Then
     with_without = 3
ElseIf Option_Check_SaleAgainstFormH.Value = True Then
     with_without = 4
     
Else

'===========================================================
'--------------------------------------------------------
    If rs.State = 1 Then rs.Close
    rs.Open "select State,tinno from SubledgerQry where " & stringyear & " and  subledger='" & txtParty1 & "'", CON
    If rs.EOF = False Then
    
    ST = rs(0)
     
    If LCase(ST) = "u.p." Then
     If Len(rs!tinno) > 0 Then
          with_without = 6
     Else
       'header = "SALE-INVOICE"
       'Sale_SLedger = "Sales Invoice"
       with_without = 5
     End If
       
    Else
       'header = "SALE-INVOICE"
       with_without = 5
     End If
    End If

'=======================================================


     
End If


If Check1_nontaxable.Value = 1 Then
   non_taxable = "1"
   with_without = 3
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
If Option_NonTaxable.Value = True Then
    nontax = "nontax"
ElseIf Option_Check_SaleAgainstFormH.Value = True Then
    nontax = "formh"
End If



'---------------------------------------------------------------------------------
Dim gen_rs As New ADODB.Recordset
Dim Sale_SLedger1 As String


If gen_rs.State = 1 Then gen_rs.Close
gen_rs.Open "select GENLEDGER,SUBLEDGER,TEXT FROM INVOICEEND where " & stringyear & " and  " & stringyear, CON

'--------------------------------------------------------------------------------


i = 1

If edit = False Then

           If copy_inv = True Then
           Else
            txtino = MaxSNo_Tax("invoicea", "INVOICENO")
           End If
           
            
            CON.BeginTrans
            
            CON.Execute "exec insertData_Invoicea " & txtino & ",'" & dateInv & "','" & dateIssue & "'," & _
            "'" & dateDispatch & "','" & txtCenteral.Text & "','SUNDRY DEBTORS','" & txtParty1 & "','" & txtDest & "'," & _
            "'" & txtModePay & "','" & txtTrans & "','" & txtBoxes & "','" & txtWeight & "','" & txtRR & "','" & dateRR.Text & "'," & _
            "'" & txtFrieght & "','" & txtWagon & "','" & Val(txtTotal) & "','" & Val(txtNet) & "','" & with_without & "','" & netrate & "','tax','" & main.username & "','" & main.username & "','" & main.session & "'," & _
            "" & main.setupid & ""
            
            
            
              CON.Execute "update invoicea  set NonTaxable = '" & nontax & "',orderby='" & IIf(txtOrdered = "", "Direct Office", txtOrdered) & "'," & _
            " aexp5='" & tax(0) & "',aexp5rate=" & txtRate(0) & ",aexp5am='" & txtamount(0) & "'," & _
            " aexp6='" & tax(1) & "',aexp6rate=" & txtRate(1) & ",aexp6am='" & txtamount(1) & "'," & _
            " aexp7='" & tax(2) & "',aexp7rate=" & txtRate(2) & ",aexp7am='" & txtamount(2) & "'," & _
            " aexp2='" & tax(3) & "',aexp2rate=" & txtRate(3) & ",aexp2am='" & txtamount(3) & "'," & _
            " aexp3='" & tax(4) & "',aexp3rate=" & txtRate(4) & ",aexp3am='" & txtamount(4) & "'," & _
            " aexp4='" & tax(5) & "',aexp4rate=" & txtRate(5) & ",aexp4am='" & txtamount(5) & "',packingno='" & Val(txtDays) & "'," & _
            " txt2a='" & cboWarehouse.Text & "',TermsPayment='" & txtPaymentRem.Text & "',permitno='" & txtPermitNo.Text & "',formNo='" & txtFormNo.Text & "', AdviceRemark='" & txtparty & "',lexp1='" & non_taxable & "'" & _
            " where invoiceno=" & txtino & " and " & stringyear & ""
          
            
            
            
            
            
             
            For i = 1 To vs.Rows - 1
            
            If vs.TextMatrix(i, 1) <> "" Then
            
            CON.Execute "exec insertData_Invoiceb " & txtino & ",'" & dateInv & "','" & txtParty1 & "'," & _
            "" & Val(vs.TextMatrix(i, 0)) & ",'" & vs.TextMatrix(i, 1) & "'," & Val(vs.TextMatrix(i, 3)) & "," & _
            "" & Val(vs.TextMatrix(i, 4)) & "," & Val(vs.TextMatrix(i, 5)) & "," & Val(vs.TextMatrix(i, 6)) & "," & _
            "'tax','" & main.username & "','" & main.username & "','" & main.session & "','" & vs.TextMatrix(i, 2) & "'," & _
            "" & main.setupid & ""
            
            End If
            
            Next
            
            For J = 0 To txtamount.Count - 1
            
                
            '----------------------------------------
            
            gen_rs.MoveFirst
            gen_rs.Find "TEXT='" & tax(J) & "'"
            If gen_rs.EOF = False Then
               Sale_SLedger1 = gen_rs!subledger
               gledger(J) = gen_rs!Genledger
               If Len(Sale_SLedger1) < 2 Then
                  Sale_SLedger1 = Sale_SLedger
               End If
            End If
            
            '----------------------------------------
                
                CON.Execute "insert into invoicec" & _
                "(INVOICENO,INVOICEDate,GENLEDGER,Subledger,GAmount,Rate,Amount,typeofinvoice,DebitOrCredit," & _
                "text,fyear,createdby,createdon,updatedby,updatedon,setupid,Subledger1) values(" & _
                "" & txtino & ",'" & Format(dateInv, "MM/dd/yyyy") & "','" & gledger(J) & "','" & Sale_SLedger1 & "'," & Val(txtNet) & "," & Val(txtRate(J)) & "," & _
                "" & Val(txtamount(J)) & ",'tax','" & Debit_Credit(J) & "','" & tax(J) & "','" & main.session & "','" & main.username & "'," & _
                "'" & Format(Date, "MM/DD/yyyy") & "','" & main.username & "','" & Format(Date, "MM/dd/yyyy") & "'," & main.setupid & ",'" & Sale_SLedger & "')"
            
            Next
            
            
            CON.CommitTrans




Else

            '-------------------------------------------
            
            
            
            CON.BeginTrans
            
            
            CON.Execute "delete from invoicea where " & stringyear & " and  invoiceNo =" & txtino & ""
            CON.Execute "delete from invoiceb where " & stringyear & " and  invoiceNo =" & txtino & ""
            CON.Execute "delete from invoicec where " & stringyear & " and  invoiceNo =" & txtino & ""
            
            
            
            CON.Execute "exec insertData_Invoicea " & txtino & ",'" & dateInv & "','" & dateIssue & "'," & _
            "'" & dateDispatch & "','" & txtCenteral.Text & "','SUNDRY DEBTORS','" & txtParty1 & "','" & txtDest & "'," & _
            "'" & txtModePay & "','" & txtTrans & "','" & txtBoxes & "','" & txtWeight & "','" & txtRR & "','" & dateRR.Text & "'," & _
            "'" & txtFrieght & "','" & txtWagon & "','" & Val(txtTotal) & "','" & Val(txtNet) & "','" & with_without & "','" & netrate & "','tax','" & main.username & "','" & main.username & "','" & main.session & "'," & _
            "" & main.setupid & ""
            

            CON.Execute "update invoicea  set NonTaxable = '" & nontax & "',orderby='" & IIf(txtOrdered = "", "Direct Office", txtOrdered) & "'," & _
            " aexp5='" & tax(0) & "',aexp5rate=" & txtRate(0) & ",aexp5am='" & txtamount(0) & "'," & _
            " aexp6='" & tax(1) & "',aexp6rate=" & txtRate(1) & ",aexp6am='" & txtamount(1) & "'," & _
            " aexp7='" & tax(2) & "',aexp7rate=" & txtRate(2) & ",aexp7am='" & txtamount(2) & "'," & _
            " aexp2='" & tax(3) & "',aexp2rate=" & txtRate(3) & ",aexp2am='" & txtamount(3) & "'," & _
            " aexp3='" & tax(4) & "',aexp3rate=" & txtRate(4) & ",aexp3am='" & txtamount(4) & "'," & _
            " aexp4='" & tax(5) & "',aexp4rate=" & txtRate(5) & ",aexp4am='" & txtamount(5) & "',packingno='" & Val(txtDays) & "'," & _
            " txt2a='" & cboWarehouse.Text & "',TermsPayment='" & txtPaymentRem.Text & "',permitno='" & txtPermitNo.Text & "',formNo='" & txtFormNo.Text & "', AdviceRemark='" & txtparty & "',lexp1='" & non_taxable & "'" & _
            " where invoiceno=" & txtino & " and " & stringyear & ""


            
            
            For i = 1 To vs.Rows - 1
            
            If vs.TextMatrix(i, 1) <> "" Then
            
            CON.Execute "exec insertData_Invoiceb " & txtino & ",'" & dateInv & "','" & txtParty1 & "'," & _
            "" & Val(vs.TextMatrix(i, 0)) & ",'" & vs.TextMatrix(i, 1) & "'," & Val(vs.TextMatrix(i, 3)) & "," & _
            "" & Val(vs.TextMatrix(i, 4)) & "," & Val(vs.TextMatrix(i, 5)) & "," & Val(vs.TextMatrix(i, 6)) & "," & _
            "'tax','" & main.username & "','" & main.username & "','" & main.session & "','" & vs.TextMatrix(i, 2) & "'," & _
            "" & main.setupid & ""
            
            End If
            
            Next
            
            
            
            For J = 0 To txtamount.Count - 1
            
            If tax(J) <> "-" Then
            
            
            
            gen_rs.MoveFirst
            gen_rs.Find "TEXT='" & tax(J) & "'"
            If gen_rs.EOF = False Then
               Sale_SLedger1 = gen_rs!subledger
               gledger(J) = gen_rs!Genledger
               If Len(Sale_SLedger1) < 2 Then
                  Sale_SLedger1 = Sale_SLedger
               End If
            End If

                
              
                
                CON.Execute "insert into invoicec" & _
                "(INVOICENO,INVOICEDate,GENLEDGER,Subledger,GAmount,Rate,Amount,typeofinvoice,DebitOrCredit," & _
                "text,fyear,createdby,createdon,updatedby,updatedon,setupid,Subledger1) values(" & _
                "" & txtino & ",'" & Format(dateInv, "MM/dd/yyyy") & "','" & gledger(J) & "','" & Sale_SLedger1 & "'," & Val(txtNet) & "," & Val(txtRate(J)) & "," & _
                "" & Val(txtamount(J)) & ",'tax','" & Debit_Credit(J) & "','" & tax(J) & "','" & main.session & "','" & main.username & "'," & _
                "'" & Format(Date, "MM/DD/yyyy") & "','" & main.username & "','" & Format(Date, "MM/dd/yyyy") & "'," & main.setupid & ",'" & Sale_SLedger & "')"
    
                
            
            End If
            
            
            Next
            
            
            
            
            CON.CommitTrans
            
            
            edit = False

'---------------------------------------------

End If


cmdEdit_4.Enabled = True
cmdDelete_3.Enabled = True
cmdSave_2.Enabled = False
'Call cmdAdd_1_Click



Exit Sub


save:

CON.RollbackTrans


If err.Number = "-2147217900" Then
   MsgBox "" & err.DESCRIPTION, vbCritical
   'txtCode.SetFocus
End If




End Sub

Private Sub cmdSearch_Click()
popuplist10 "select [INVOICENO], [INVOICEDATE], [SUBLEDGER] from INVOICEA where " & stringyear & " order by cast(INVOICENO as int)", CON
End Sub

Private Sub cmdSearch_GotFocus()
txtino = PopUpValue1
searchData
PopUpValue1 = ""
PopUpValue2 = ""
PopUpValue3 = ""
End Sub

Private Sub dateInv_LostFocus()
dateIssue.Text = dateInv.Text
dateDispatch.Text = dateInv.Text
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
setwidth


If rs.State = 1 Then rs.Close

st1 = "select INVOICENO,INVOICEDATE,IssueDate,DisPatchDate,CentralExise,SUBLEDGER,STATION," & _
"ModeOfPayment,THROUGH,BUNDLES,WEIGHT,BILTYNO,BILTYDate,FREIGHT,TXT1,[with_withoutFormc],[NetRate]," & _
"NonTaxable,Orderby,gamount,netAmount,netrate,txt2a,TermsPayment,permitno,formNo,AdviceRemark,lexp1,packingno from invoicea where invoiceno=" & txtino & " and  " & stringyear & ""
     
     
     
On Error Resume Next



     
rs.Open st1, CON

If rs.EOF = True Then
   Exit Sub
End If

If rs.EOF = False Then
  
  txtDays = rs!packingno & ""
  
  If rs!lexp1 = "1" Then
     Check1_nontaxable.Value = 1
  Else
    Check1_nontaxable.Value = 0
  End If
  
  txtPaymentRem.Text = rs!TermsPayment & ""
  cboWarehouse.Text = rs!txt2a & ""
  txtparty.Text = rs!ADVICEREMARK & ""
  
  txtParty1.Enabled = False
  
  dateInv = rs!InvoiceDate
  dateIssue = Format(rs!IssueDate, "dd/MM/yyyy")
  dateDispatch = Format(rs!DisPatchDate, "dd/MM/yyyy")
  txtCenteral = rs!CentralExise
  txtParty1 = rs!subledger
  txtDest = rs!station
  txtModePay = rs!ModeOfPayment
  txtTrans = rs!through
  txtBoxes = rs!bundles
  txtWeight = rs!weight
  dateRR = rs!BILTYDATE
  txtRR = rs!biltyno
  txtWagon = rs!txt1
  txtFrieght = rs!freight
  txtOrdered = rs!ORDERBY & ""
  
  
  txtPermitNo.Text = rs!permitno & ""
  txtFormNo.Text = rs!formNo & ""
  
  If rs![with_withoutFormc] = 1 Then
     Option_with.Value = True
  ElseIf rs![with_withoutFormc] = 2 Then
     Option_without.Value = True
  ElseIf rs![with_withoutFormc] = 3 Then
     Option_NonTaxable.Value = True
  ElseIf rs![with_withoutFormc] = 4 Then
     Option_Check_SaleAgainstFormH.Value = True
  End If
  
  
  
  
  If rs![netrate] = "y" Then
     Check_Net.Value = 1
  Else
     Check_Net.Value = 0
  End If
  
  
  
  txtTotal = Format(rs!GAmount, "0.00")
  txtNet = Format(rs!netamount, "0.00")
  
  
  
  If rs.State = 1 Then rs.Close
   rs.Open "SELECT [SUBLEDGER],address1,address2,[City],[District],[State],[bcityid],typeofcust,tinno " & _
   "from [ExportData].[dbo].[SubledgerQry] WHERE SUBLEDGER ='" & txtParty1 & "' and  " & stringyear & "", CON
   If rs.EOF = False Then
     txtAdd = rs![subledger] & " " & vbCrLf & rs!address1 & " " & rs!address1 & vbCrLf & rs![CITY] + vbCrLf + rs![District] & "," & rs![State]
     Me.ptype.Caption = rs!TypeOfCust
     lbltin.Caption = "Buyer's Tin No. : " & rs!tinno
   End If
  
End If

lblQty = 0

If rs.State = 1 Then rs.Close
rs.Open "select * from invoiceb where INVOICENO=" & txtino.Text & " and  " & stringyear & " order by printorder", CON, adOpenDynamic, adLockOptimistic
For i = 1 To rs.RecordCount
If rs.EOF = False Then
   
    vs.TextMatrix(i, 0) = rs.Fields("printorder").Value
    vs.TextMatrix(i, 1) = rs.Fields("BOOKCODE").Value
    
    vs.TextMatrix(i, 3) = rs.Fields("Quantity").Value
    vs.TextMatrix(i, 4) = rs.Fields("Rate").Value
    vs.TextMatrix(i, 5) = Format(rs.Fields("NetRate").Value, ".00")
    vs.TextMatrix(i, 6) = Format(rs.Fields("Amount").Value, ".00")
    
    lblQty = (Val(lblQty) + rs.Fields("Quantity").Value)
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select ProductQuality,TypeofProduct,rulling,rate,NoofPages from copymaster " & _
    "where " & stringyear & " and  bookno='" & vs.TextMatrix(i, 1) & "'", CON                     ' and  " & stringyear & "", CON
    If rs1.EOF = False Then
       vs.TextMatrix(i, 2) = rs1!TypeofProduct + " (" + rs1!rulling + ")" + str(rs1!NoOfPages) + " " + rs1!ProductQuality
    End If
    
rs.MoveNext
End If
Next


'Total

            
  

If rs.State = 1 Then rs.Close
rs.Open "select [aexp1],[aexp1rate],[aexp1am],[aexp2],[aexp2rate],[aexp2am]," & _
"[aexp3],[aexp3rate],[aexp3am],[aexp4],[aexp4rate],[aexp4am],[aexp5],[aexp5rate],[aexp5am],[aexp6],[aexp6rate],[aexp6am],[aexp7],[aexp7rate],[aexp7am] from invoicea " & _
" where INVOICENO=" & txtino.Text & " and  " & stringyear & "", CON, adOpenDynamic, adLockOptimistic
If rs.EOF = False Then
   
   
   tax(0) = rs![aexp5]
   txtRate(0) = rs![aexp5rate]
   txtamount(0) = Format(rs![aexp5am], "0.00")
   
   tax(1) = rs![aexp6]
   txtRate(1) = rs![aexp6rate]
   txtamount(1) = Format(rs![aexp6am], "0.00")
   
   tax(2) = rs![aexp7]
   txtRate(2) = rs![aexp7rate]
   txtamount(2) = Format(rs![aexp7am], "0.00")
   
   
   '---------------------------------------------------------------
   
   tax(3) = rs![aexp2]
   txtRate(3) = rs![aexp2rate]
   txtamount(3) = Format(rs![aexp2am], "0.00")
   
   tax(4) = rs![aexp3]
   txtRate(4) = rs![aexp3rate]
   txtamount(4) = Format(rs![aexp3am], "0.00")
   
   
   tax(5) = rs![aexp4]
   txtRate(5) = rs![aexp4rate]
   txtamount(5) = Format(rs![aexp4am], "0.00")
   
   'tax(3) = rs![aexp4]
   'txtRate(3) = rs![aexp4rate]
   'txtamount(3) = Format(rs![aexp4am], "0.00")
     
     
   cmdSave_2.Enabled = False

End If






thead = ""

If Option_NonTaxable.Value = True Then
   thead = Option_NonTaxable.Caption
ElseIf Option_Check_SaleAgainstFormH.Value = True Then
   thead = Option_Check_SaleAgainstFormH.Caption
ElseIf Option_with.Value = True Then
   thead = Option_with.Caption
ElseIf Option_without.Value = True Then
   thead = Option_without.Caption
   
End If


If rs1.State = 1 Then rs1.Close
rs1.Open "select subledger from InvoiceSubLedger where " & stringyear & " and  taxhead='" & thead & "'", CON
If rs1.EOF = False Then
   Sale_SLedger = rs1!subledger
End If



'-----------------------

If rs.State = 1 Then rs.Close
rs.Open "select State,tinno from SubledgerQry where " & stringyear & " and  subledger='" & txtParty1 & "'", CON
If rs.EOF = False Then
    ST = rs(0)
    If LCase(ST) = "u.p." Then
     If Len(rs!tinno) > 0 Then
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




K1 = 0
ReDim gledger(txtamount.Count)
ReDim Debit_Credit(txtamount.Count)

If rs.State = 1 Then rs.Close
rs.Open "select text,Rate,GENLEDGER,DEBITORCREDIT from invoicec " & _
" where  where " & stringyear & " and invoiceNo=" & txtino & " order by auto", CON
For kk1 = 1 To txtamount.Count + 1
    If rs.EOF = False Then
       tax(K1) = rs(0)
       txtRate(K1) = rs(1)
       gledger(K1) = rs!Genledger
       Debit_Credit(K1) = rs!DEBITORCREDIT
       K1 = K1 + 1
       rs.MoveNext
    End If
Next


'==============================================================
'==============================================================
'--------------------------------------------------------------

If rs.State = 1 Then rs.Close
rs.Open "select State,tinno from SubledgerQry where " & stringyear & " and  subledger='" & txtParty1 & "'", CON
If rs.EOF = False Then


If Len(rs!tinno) > 0 Then
   lblinv.Caption = "TAX INVOICE"
   lbltin.Caption = "Tin No : " & rs!tinno
Else
   lblinv.Caption = "SALE INVOICE"
   lbltin.Caption = ""
End If


If ST = "U.P." Then

    If Len(rs!tinno) = 0 Then
       Sale_SLedger = "sales Invoice"
     Else
       Sale_SLedger = "Tax Invoice Sales"
    End If
    
   If Option_NonTaxable.Value = True Then
      Sale_SLedger = "sales Non - Taxable"
   End If

Else


If Option_with.Value = True Then
   Sale_SLedger = "Sales Agst. Form-C"
ElseIf Option_without.Value = True Then
   Sale_SLedger = "Sales Without Form -C"
ElseIf Option_NonTaxable.Value = True Then
   Sale_SLedger = "sales Non - Taxable"
ElseIf Option_Check_SaleAgainstFormH.Value = True Then
   Sale_SLedger = "sales Agst.Form - H"
End If


End If
End If
  



subtot = (Val(txtTotal.Text) + Val(txtamount(0)) + Val(txtamount(1)) + Val(txtamount(2)))
lblGTotal.Caption = Format(subtot, ".00")



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
 
 copy_inv = False
 
 setwidth
 
 dateInv.Text = Format(Date, "dd/MM/yyyy")
 
 txtino = MaxSNo_Tax("invoicea", "INVOICENO")
 
 'txtRate(0) = VAT
 
 dateRR.Text = Format(Date, "dd/MM/yyyy")
 
 
 
 withForm = VAT1
 withoutForm = VAT
 
 Frame1.Width = ((5775 / 2) - 430)
 
 
 
If rs.State = 1 Then rs.Close
rs.Open "select subledger from sledger where " & stringyear & " and  gledger='SALES'", CON
While rs.EOF = False
  List_SaleSubledger.AddItem rs(0)
  rs.MoveNext
Wend
 
cboPrint.ListIndex = 0

If rs.State = 1 Then rs.Close
rs.Open "select distinct agentname from agentmaster", CON
While rs.EOF = False
  txtOrdered.AddItem rs(0)
  rs.MoveNext
Wend
 
If main.username = "admin" Then
   check_AccSet.Enabled = True
Else
  check_AccSet.Enabled = False
End If
 
 
If rs.State = 1 Then rs.Close
rs.Open "select * from Warehouse order by Warehouse", CON, adOpenKeyset
While rs.EOF = False
  cboWarehouse.AddItem rs(0)
  rs.MoveNext
Wend
 
cboWarehouse.ListIndex = 0
 
ButtonPermission cmdSave_2, cmdDelete_3, cmdEdit_4
 
End Sub
Sub setwidth()
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
''   If KeyCode = 13 Then txtParty1.SetFocus
''End Sub

Private Sub Option_Check_SaleAgainstFormH_Click()
 
If Option_Check_SaleAgainstFormH.Value = True Then
  ' Check1_nontaxable.Value = 0
End If
 

 
 taxhead = Option_Check_SaleAgainstFormH.Caption
' SearchDataForSaleSubledger

End Sub

Private Sub Option_NonTaxable_Click()


If Option_NonTaxable.Value = True Then
'   Check1_nontaxable.Value = 0
End If



taxhead = Option_NonTaxable.Caption
'SearchDataForSaleSubledger

End Sub

Private Sub Option_NonTaxable_DblClick()
  Option_NonTaxable.Value = False
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
'''popuplist2 "select INVOICENO,INVOICEDATE,SUBLEDGER from invoicea order by INVOICENO", CON
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

If Option_with.Value = True Then
 '  Check1_nontaxable.Value = 0
End If

taxhead = Option_with.Caption


End Sub

Private Sub Option_without_Click()
'If Option_with.Value = True Then
'  txtRate(0) = withForm
'Else
'  txtRate(0) = withoutForm
'End If

If Option_without.Value = True Then
'   Check1_nontaxable.Value = 0
End If



taxhead = Option_without.Caption


'FatchTaxFromSate

'SearchDataForSaleSubledger
End Sub

Private Sub txtamount_GotFocus(Index As Integer)
HIT
End Sub

Private Sub txtamount_LostFocus(Index As Integer)

Dim sum1 As Double

txtNet = 0
'lblGTotal = 0
sum1 = 0


sum1 = (Val(txtTotal.Text) + Val(txtamount(0)) + Val(txtamount(1)) + Val(txtamount(2)) + Val(txtamount(3)))


txtNet = Val(txtNet) + sum1

'-------------------------------------------------------------
Add_Amt = Val(txtamount(4))
Less_Amt = Val(txtamount(5))
'-------------------------------------------------------------

txtNet = (Val(txtNet) + Add_Amt)
txtNet = (Val(txtNet) + Less_Amt)



End Sub

Private Sub txtino_GotFocus()
HIT
End Sub

Private Sub txtino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  searchData
  'showCat
End If
End Sub


Private Sub txtparty_GotFocus()
HIT
If PopUpValue1 <> "" Then
   txtParty1 = PopUpValue2
   
   
  If rs.State = 1 Then rs.Close
   rs.Open "SELECT [SUBLEDGER],address1,address2,[City],[District],[State],[bcityid],typeofcust,DESCFORINVOICE " & _
   "from [ExportData].[dbo].[SubledgerQry] where " & stringyear & " and  SUBLEDGER ='" & txtParty1 & "'", CON
   If rs.EOF = False Then
     txtAdd = rs![subledger] & " " & vbCrLf & rs!address1 & " " & rs!address1 & vbCrLf & rs![CITY] + vbCrLf + rs![District] & "," & rs![State]
     Me.ptype.Caption = rs!TypeOfCust
     Me.txtparty.Text = rs!DESCFORINVOICE & ""
   End If
   
   
   
   FatchTaxFromSate
   showCat
   
   txtDest.SetFocus
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   
End If
End Sub

Private Sub txtparty_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
  
  If rs.State = 1 Then rs.Close
   rs.Open "SELECT [SUBLEDGER],address1,address2,[City],[District],[State],[bcityid],typeofcust " & _
   "from [ExportData].[dbo].[SubledgerQry] where " & stringyear & " and  SUBLEDGER ='" & txtParty1 & "'", CON
   If rs.EOF = False Then
     txtAdd = rs![subledger] & " " & vbCrLf & rs!address1 & " " & rs!address1 & vbCrLf & rs![CITY] + vbCrLf + rs![District] & "," & rs![State]
     Me.ptype.Caption = rs!TypeOfCust
   End If
   
   
   
   FatchTaxFromSate
   showCat
   
   txtDest.SetFocus
   
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
     txtqty.SelLength = 10
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



Private Sub VS_KeyDown(KeyCode As Integer, Shift As Integer)

  
  
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
    
    
    If rs.State = 1 Then rs.Close
    rs.Open "select ProductQuality,TypeofProduct,rulling,rate,NoofPages from copymaster " & _
    "where " & stringyear & " and  bookno='" & vs.TextMatrix(vs.RowSel, 1) & "'", CON
    If rs.EOF = False Then
          
          Item = rs!TypeofProduct + " (" + rs!rulling + ")" + str(rs!NoOfPages) + " " + rs!ProductQuality
       
          vs.TextMatrix(vs.RowSel, 2) = Item
          
          vs.TextMatrix(vs.RowSel, 4) = rs.Fields("Rate").Value
    Else
       Exit Sub
    End If
    
    SendKeys "{right}"
    SendKeys "{right}"
 
 End If
 
 
 If Check_Net.Value = 1 Then
 
 
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
    If Option_NonTaxable.Value = False Then
    'If Check_Nontaxable.Value = 0 Then
        If Option_with.Value = True Then
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


       


showCat


End If



End Sub

Sub FatchTaxFromSate()

On Error Resume Next


Dim with_without As String
Dim B_CH  As Boolean

B_CH = False

If Option_with.Value = True Then
   with_without = Option_with.Caption
Else
   with_without = Option_without.Caption
End If




If rs1.State = 1 Then rs1.Close
rs1.Open "select [State],tinno from SubledgerQry where subledger='" & txtParty1 & "' and  " & stringyear & "", CON
If rs1.EOF = False Then
   ST = rs1(0)
   B_CH = True
End If





'===============================================
'===============================================
'===============================================
'===============================================
'===============================================
'===============================================
'===============================================
'===============================================






If rs1.State = 1 Then rs1.Close
rs1.Open "select add_val,less_val from [state_tax_list] where " & stringyear & " and  statename='" & ST & "'" & _
" and with_without='" & with_without & "'", CON
If rs1.EOF = False Then
   VAT_Add = rs1(0)
   VAT_less = rs1(1)
   

CON.Execute "update invoiceend set ryn='y'  where  " & stringyear
   


   
If (Option_NonTaxable.Value = True Or Option_Check_SaleAgainstFormH.Value = True) Then
 
 CON.Execute "update invoiceend set ryn='n'  where [Rate] >0 and  " & stringyear & ""

Else
 
 CON.Execute "update invoiceend set ryn='n'  where  ([Rate] <> " & VAT_Add & " and [Rate] >0) and  " & stringyear & ""

End If

If ST = "U.P." Then
   CON.Execute "update invoiceend set ryn='y'  where  " & stringyear & " and text like '%VAT%' and [Rate] = " & VAT_Add & ""
   CON.Execute "update invoiceend set ryn='n'  where  " & stringyear & " and text like '%CST%' and [Rate] = " & VAT_Add & ""
Else
   CON.Execute "update invoiceend set ryn='n'  where  " & stringyear & " and text like '%VAT%' and [Rate] = " & VAT_Add & ""
   CON.Execute "update invoiceend set ryn='y'  where  " & stringyear & " and text like '%CST%' and [Rate] = " & VAT_Add & ""
End If

End If






If Option_NonTaxable.Value = True Then
   txtRate(0) = 0
End If


K1 = 0

If (Option_NonTaxable.Value = True Or Option_Check_SaleAgainstFormH.Value = True) Then
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

If rs1.State = 1 Then rs1.Close
rs1.Open "select printorder from invoiceend where  " & stringyear & " and RYN='y' order by printorder", CON
While rs1.EOF = False
  CON.Execute "update invoiceend set id=" & kk1 & "  where  printorder=" & rs1(0) & " and " & stringyear
  kk1 = kk1 + 1
  rs1.MoveNext
Wend

'============================================================



K1 = 3

ReDim gledger(txtamount.Count)
ReDim Debit_Credit(txtamount.Count)


For kk1 = 3 To txtamount.Count - 1

If rs1.State = 1 Then rs1.Close
rs1.Open "select text,Rate,GENLEDGER,DEBITORCREDIT from invoiceend where  " & stringyear & " and id=" & kk1 & " and RYN='y' order by printorder", CON
If rs1.EOF = False Then
   tax(K1) = rs1(0)
   txtRate(K1) = rs1(1)
   gledger(K1) = rs1!Genledger
   Debit_Credit(K1) = rs1!DEBITORCREDIT
   K1 = K1 + 1


If B_CH = True Then
   tax(K1) = nontax
   txtRate(K1) = 0
   gledger(K1) = "-"
   Debit_Credit(K1) = "-"
   B_CH = False
   K1 = K1 + 1
End If


End If

Next




'====================================================


tax(4) = "Post./Courier Charges"
tax(5) = "Round Off"



If (Option_NonTaxable.Value = False And Option_Check_SaleAgainstFormH.Value = False) Then
'================================

If UCase(ST) = "U.P." Then


    If (Option_NonTaxable.Value = False And Option_Check_SaleAgainstFormH.Value = False) Then
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select text,Rate,GENLEDGER,DEBITORCREDIT from invoiceend where  " & stringyear & "  and (SUBLEDGER LIKE '%OUTPUT%') order by RATE", CON
    If rs1.EOF = False Then
       tax(3) = rs1!Text
       txtRate(3) = rs1!rate
    End If
    
    End If

'For Exup
'===========================================
'===========================================

ElseIf UCase(ST) <> "U.P." Then


If Option_with.Value = True Then


    If rs1.State = 1 Then rs1.Close
    rs1.Open "select text,Rate,GENLEDGER,DEBITORCREDIT from invoiceend where  " & stringyear & "  and text LIKE '%CST%' AND RATE<4 order by RATE", CON
    If rs.EOF = False Then
       tax(3) = rs1!Text
       txtRate(3) = rs1!rate
    End If

ElseIf Option_without.Value = True Then

    If rs1.State = 1 Then rs1.Close
    rs1.Open "select text,Rate,GENLEDGER,DEBITORCREDIT from invoiceend where  " & stringyear & "  and text LIKE '%CST%' AND RATE>4 order by RATE", CON
    If rs.EOF = False Then
       tax(3) = rs1!Text
       txtRate(3) = rs1!rate
    End If


End If
End If
End If

'================Non Taxable========================================

If (Option_NonTaxable.Value = False Or Option_Check_SaleAgainstFormH.Value = False) Then

    If (Option_NonTaxable.Value = True) Then
        tax(3) = "Non Taxable"
        txtRate(3) = 0
    ElseIf (Option_Check_SaleAgainstFormH.Value = True) Then
        tax(3) = "Sale Against Form-H"
        txtRate(3) = 0
    End If

End If


'================Non Taxable========================================
'================Non Taxable========================================
'===================================================================
'================ SearchDataForSaleSubledger =======================



End Sub
Private Sub vs_SelChange()
  
  If vs.Col = 2 Then
     vs.Editable = flexEDNone
  Else
     vs.Editable = flexEDKbdMouse
  End If

End Sub
