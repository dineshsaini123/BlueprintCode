VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPackingList 
   Caption         =   "Packing List"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   14835
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdUpDateGrid 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Update Grid"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   7440
      Width           =   1875
   End
   Begin VB.TextBox txtShipToAdd 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   10425
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   69
      Top             =   2700
      Width           =   3915
   End
   Begin VB.TextBox txtForCustom 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   10425
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   68
      Top             =   1650
      Visible         =   0   'False
      Width           =   3900
   End
   Begin VB.CheckBox Check1 
      Caption         =   "For Custom"
      Height          =   195
      Left            =   10425
      TabIndex        =   67
      Top             =   1425
      Width           =   2115
   End
   Begin VB.CheckBox Check_manullay 
      Caption         =   "Enter Packing No Manullay"
      Height          =   255
      Left            =   1380
      TabIndex        =   66
      Top             =   540
      Width           =   2895
   End
   Begin VB.CommandButton cmdAdd1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Add "
      Height          =   330
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Add Teacher"
      Top             =   3180
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtGwt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   13080
      TabIndex        =   25
      Top             =   8805
      Width           =   1155
   End
   Begin VB.TextBox txtPortDischarge 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11850
      TabIndex        =   20
      Top             =   1050
      Width           =   2505
   End
   Begin VB.TextBox txtPortLoading 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11850
      TabIndex        =   19
      Top             =   690
      Width           =   2505
   End
   Begin VB.TextBox txtTermsOfPayment 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   14
      Top             =   2325
      Width           =   3585
   End
   Begin VB.TextBox txtBrand 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   12
      Top             =   1725
      Width           =   3585
   End
   Begin VB.TextBox txtTermsOfDelivery 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   13
      Top             =   2025
      Width           =   3585
   End
   Begin VB.TextBox txtBuyerBatchNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   11
      Top             =   1425
      Width           =   3585
   End
   Begin VB.TextBox txtBuyerOrderNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   10
      Top             =   1125
      Width           =   3585
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   240
      TabIndex        =   26
      Top             =   7905
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
         Style           =   1  'Graphical
         TabIndex        =   33
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
         Style           =   1  'Graphical
         TabIndex        =   32
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
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Style           =   1  'Graphical
         TabIndex        =   30
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
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Left            =   1275
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   135
         Width           =   1230
      End
   End
   Begin VB.TextBox txtAdd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   735
      Left            =   6600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   390
      Width           =   3690
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   8
      Top             =   90
      Width           =   3690
   End
   Begin VB.TextBox txtino 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1350
      TabIndex        =   0
      Top             =   165
      Width           =   1155
   End
   Begin VB.TextBox txtFDestination 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   16
      Top             =   2925
      Width           =   3585
   End
   Begin VB.TextBox txtExporterPan 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1740
      TabIndex        =   5
      Text            =   "AAEFC-7614G"
      Top             =   2205
      Width           =   2415
   End
   Begin VB.TextBox txtIECode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1740
      TabIndex        =   4
      Text            =   "0505055317"
      Top             =   1845
      Width           =   2415
   End
   Begin VB.TextBox txtPlaceofPrecarriage 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11850
      TabIndex        =   18
      Top             =   210
      Width           =   2475
   End
   Begin VB.TextBox txtPreCarriageBy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   17
      Top             =   3225
      Width           =   3585
   End
   Begin VB.TextBox txtEPSG 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1740
      TabIndex        =   7
      Text            =   "0530142551/5/11/00 Dt. 04/12/06"
      Top             =   2985
      Width           =   2715
   End
   Begin VB.TextBox txtNetW 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   13080
      TabIndex        =   24
      Top             =   8445
      Width           =   1155
   End
   Begin VB.TextBox txtCBM 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   13080
      TabIndex        =   23
      Top             =   8085
      Width           =   1155
   End
   Begin VB.TextBox txtTotalCartoons 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   13080
      TabIndex        =   22
      Top             =   7785
      Width           =   1155
   End
   Begin VB.TextBox txtTotalPSC 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   13080
      TabIndex        =   21
      Top             =   7470
      Width           =   1155
   End
   Begin VB.ComboBox txtExportCurrency 
      Height          =   315
      ItemData        =   "frmPackingList.frx":0000
      Left            =   1740
      List            =   "frmPackingList.frx":0010
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2565
      Width           =   1395
   End
   Begin VB.TextBox txtCountryOrigin 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   15
      Text            =   "India"
      Top             =   2625
      Width           =   1245
   End
   Begin Crystal.CrystalReport CR 
      Left            =   13515
      Top             =   6585
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3885
      Left            =   120
      TabIndex        =   34
      Top             =   3525
      Width           =   16455
      _cx             =   29025
      _cy             =   6853
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   500
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
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   180
         Begin MSDataListLib.DataCombo cboItem 
            Height          =   2310
            Left            =   0
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   0
            Width           =   225
            _ExtentX        =   397
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
   Begin MSMask.MaskEdBox dateInv 
      Height          =   315
      Left            =   3120
      TabIndex        =   1
      Top             =   165
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
      Left            =   3120
      TabIndex        =   2
      Top             =   945
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
      Left            =   3120
      TabIndex        =   3
      Top             =   1305
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
   Begin VB.Label Label1 
      Caption         =   "Ship To Address  :"
      Height          =   300
      Index           =   5
      Left            =   10425
      TabIndex        =   71
      Top             =   2475
      Width           =   1785
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Dealer/Buyer :"
      Height          =   300
      Index           =   3
      Left            =   8400
      TabIndex        =   70
      Top             =   2925
      Width           =   1785
   End
   Begin VB.Label Label10 
      Caption         =   "Press F5 For Insert Raw in Grid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6180
      TabIndex        =   65
      Top             =   7440
      Width           =   3375
   End
   Begin VB.Label Label9 
      Caption         =   "Press F4 For Delete Raw from Grid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   63
      Top             =   7440
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Gross Weight (Kg.)"
      Height          =   255
      Left            =   11700
      TabIndex        =   62
      Top             =   8805
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Port of Discharge"
      Height          =   270
      Index           =   15
      Left            =   10410
      TabIndex        =   61
      Top             =   990
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "Terms of Payment"
      Height          =   210
      Index           =   14
      Left            =   4560
      TabIndex        =   60
      Top             =   2325
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "Port of Loading"
      Height          =   270
      Index           =   13
      Left            =   10410
      TabIndex        =   59
      Top             =   690
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "Terms of Delivery"
      Height          =   210
      Index           =   11
      Left            =   4560
      TabIndex        =   58
      Top             =   2025
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "Date &&  Time of Dispatch"
      Height          =   270
      Index           =   8
      Left            =   120
      TabIndex        =   57
      Top             =   1305
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Date &&  Time of issue of invoice"
      Height          =   270
      Index           =   7
      Left            =   120
      TabIndex        =   56
      Top             =   945
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Exporter's I.E. code"
      Height          =   300
      Index           =   4
      Left            =   120
      TabIndex        =   55
      Top             =   1845
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Date "
      Height          =   270
      Index           =   1
      Left            =   2580
      TabIndex        =   54
      Top             =   165
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Packing List No"
      Height          =   270
      Index           =   0
      Left            =   150
      TabIndex        =   53
      Top             =   165
      Width           =   1260
   End
   Begin VB.Label PlaceofCarriage 
      AutoSize        =   -1  'True
      Caption         =   "Place of Recipt by PreCarrige"
      Height          =   390
      Left            =   10410
      TabIndex        =   52
      Top             =   210
      Width           =   1515
      WordWrap        =   -1  'True
   End
   Begin VB.Label PreCarriageBy 
      AutoSize        =   -1  'True
      Caption         =   "Pre Carriage By"
      Height          =   195
      Left            =   4560
      TabIndex        =   51
      Top             =   3285
      Width           =   1875
   End
   Begin VB.Label ExportCurrency 
      AutoSize        =   -1  'True
      Caption         =   "Brand"
      Height          =   195
      Index           =   0
      Left            =   4560
      TabIndex        =   50
      Top             =   1725
      Width           =   1920
   End
   Begin VB.Label ExporterPan 
      AutoSize        =   -1  'True
      Caption         =   "Exporter's Pan"
      Height          =   195
      Left            =   120
      TabIndex        =   49
      Top             =   2205
      Width           =   1020
   End
   Begin VB.Label OrderDate 
      AutoSize        =   -1  'True
      Caption         =   "Order Date(s)"
      Height          =   195
      Left            =   14340
      TabIndex        =   48
      Top             =   -495
      Width           =   945
   End
   Begin VB.Line Line5 
      X1              =   11040
      X2              =   19860
      Y1              =   -675
      Y2              =   -675
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Dealer/Buyer :"
      Height          =   300
      Index           =   2
      Left            =   4560
      TabIndex        =   47
      Top             =   90
      Width           =   1785
   End
   Begin VB.Label BuyersOrderNum 
      AutoSize        =   -1  'True
      Caption         =   "Buyer's Order No(s)."
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   46
      Top             =   1125
      Width           =   1890
   End
   Begin VB.Label BuyerBatchNum 
      AutoSize        =   -1  'True
      Caption         =   "Buyer's Batch No."
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   45
      Top             =   1425
      Width           =   1875
   End
   Begin VB.Label CountryofDestination 
      AutoSize        =   -1  'True
      Caption         =   "Country of Final Destination"
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   44
      Top             =   2925
      Width           =   1935
   End
   Begin VB.Label ExportCurrency 
      AutoSize        =   -1  'True
      Caption         =   "Export Currency"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   43
      Top             =   2625
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "EPCG License No."
      Height          =   195
      Left            =   120
      TabIndex        =   42
      Top             =   2985
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Net Weight (Kg.)"
      Height          =   255
      Left            =   11700
      TabIndex        =   41
      Top             =   8445
      Width           =   1275
   End
   Begin VB.Label Label5 
      Caption         =   "Total CBM"
      Height          =   195
      Left            =   11700
      TabIndex        =   40
      Top             =   8145
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Total Cartons"
      Height          =   195
      Left            =   11700
      TabIndex        =   39
      Top             =   7845
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Total Pcs"
      Height          =   255
      Left            =   11700
      TabIndex        =   38
      Top             =   7485
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Country of Origin"
      Height          =   255
      Left            =   4560
      TabIndex        =   37
      Top             =   2625
      Width           =   1755
   End
End
Attribute VB_Name = "frmPackingList"
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

 Dim d1, D2 As Long
 

Dim header As String

Const dis1 As Double = 25
Const dis2 As Double = 9.09
Const dis3 As Double = 7.5
Const dis4 As Double = 5





Private Sub cmdMain_Click()
Unload Me
End Sub
Sub cellposi()
 'VsFrame.Width = 3165
 VsFrame.TOP = vs.TOP + ((vs.CellTop)) - 1400
 VsFrame.Left = (vs.Left) - 200
End Sub
Sub Total()

On Error Resume Next


'txtTotalPSC.Text = 0
'
'
'total1 = 0
'
'
'If Val(vs.TextMatrix(J, 12)) > 0 Then
'txtTotalPSC.Text = (Val(txtTotalPSC.Text) + Val(vs.TextMatrix(J, 12)))
'End If
''
'
'
'
'
'txtTotal = Format(txtTotal, ".00")
'
'
'


End Sub

Sub cellposiVs()
 Vs1Frame.Width = 2500
 Vs1Frame.TOP = vs1.TOP + ((vs1.CellTop))
 Vs1Frame.Left = (vs1.Left) + 550
End Sub
Sub AddItemInGrid1()
'
'    Dim rs_4 As New ADODB.Recordset
'
'    rs_4.Open "select * from bm order by BKDESC", con, adOpenDynamic, adLockOptimistic
'
'    Set cboitemvs1.RowSource = rs_4
'    cboitemvs1.ListField = "BKDESC"
'    cboitemvs1.BoundColumn = "BKCODE"
'    cboitemvs1.ReFill
    
End Sub
Sub AddItemInGrid3()
    'Adodc1.ConnectionString = "filedsn=Saru"
    'Adodc1.CommandType = adCmdText
'
'    Dim rs_3 As New ADODB.Recordset
'
'    'rs_3.Open "select * from ItemMaster " & stringyear & " and  (ItemGp='Finish Item' or ItemGp='Scrap' or ItemGp='Losses' or (ItemGp='Chunnus (Pointing)' or ItemGp='Chunnus (Sonicut)') or  Itemgp='Rejuction') order by ItemName", con, adOpenDynamic, adLockOptimistic
'    rs_3.Open "select * from ItemMaster order by ItemName", con, adOpenDynamic, adLockOptimistic
'
'    'Adodc1.Refresh
'    Set cboItemVs3.RowSource = rs_3
'    cboItemVs3.ListField = "ItemName"
'    cboItemVs3.BoundColumn = "ItemName"
'    cboItemVs3.ReFill
    
End Sub
Sub AddItemInGrid2()
'    'Adodc1.ConnectionString = "filedsn=Saru"
'    'Adodc1.CommandType = adCmdText
'    Dim rs_2 As New ADODB.Recordset
'
'    'rs_2.Open "select * from ItemMaster " & stringyear & " and  (ItemGp='Semi Finish (R/D)' or ItemGp= 'Semi Finish (Store)' or Itemgp='Rejuction') order by ItemName", con, adOpenDynamic, adLockOptimistic
'    rs_2.Open "select * from ItemMaster order by ItemName", con, adOpenDynamic, adLockOptimistic
'
'    Set cboItemVs2.RowSource = rs_2
'    cboItemVs2.ListField = "ItemName"
'    cboItemVs2.BoundColumn = "ItemName"
'    cboItemVs2.ReFill
    
End Sub
Sub AddItemInGrid()
'    'Adodc1.ConnectionString = "filedsn=Saru"
'    'Adodc1.CommandType = adCmdText
'    Dim rs_1 As New ADODB.Recordset
'
'    'rs_1.Open "select * from ItemMaster " & stringyear & " and  ItemGp='Raw Item' or ItemGp='Scrap' or (ItemGp='Chunnus (Pointing)' or ItemGp='Chunnus (Sonicut)') order by ItemName", con, adOpenDynamic, adLockOptimistic
'    rs_1.Open "select * from bm order by BKDESC", CON, adOpenDynamic, adLockOptimistic
'
'    Set cboItem.RowSource = rs_1
'    cboItem.ListField = "BKDESC"
'    cboItem.BoundColumn = "BKCODE"
'    cboItem.ReFill
    
End Sub
Private Sub cboFinishItem_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtqty.SetFocus
End Sub


Private Sub cbogodown_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then txtRemarks.SetFocus
End Sub

Private Sub cbogodown_LostFocus()
If cboGodown = "" Then
   MsgBox "Select Godown Name ..", vbCritical
   cboGodown.SetFocus
   Exit Sub
End If
End Sub








Private Sub cboItem_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        cellposi
        If cboItem.Text = "" Then
        
        VsFrame.Visible = False
        cmdSave_2.SetFocus
        Exit Sub
        End If
'        vs.TextMatrix(vs.RowSel, 0) = cboItem.Text
'        vs.TextMatrix(vs.RowSel, 6) = cboItem.BoundText
        
        
        vs.SetFocus
        
     ElseIf KeyCode = 27 Then
       
          VsFrame.Visible = False
        
     End If
End Sub
Sub saveInMaster()
         On Error Resume Next
      
         If rs.State = 1 Then rs.Close
         rs.Open "select * from ItemMaster where " & stringyear & " and  ItemName='" & iitem1 & "'", CON, adOpenDynamic, adLockOptimistic
         If rs.EOF = True Then
            rs.addNew
            rs.Fields("ItemGp").Value = frmAddMaster.cbogp.Text
            rs.Fields("ItemName").Value = iitem1
            rs.Fields("Unit").Value = "Kg"
            rs.Update
         Else
            MsgBox "This Item Already Exist !!", vbCritical
            Exit Sub
         End If
         frmAddMaster.Visible = False
  
End Sub
Private Sub cboitemvs1_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        cellposiVs
        vs1.TextMatrix(vs1.RowSel, 0) = cboitemvs1.Text
        
        If rs.State = 1 Then rs.Close
        rs.Open "select * from ItemMaster where " & stringyear & " and  ItemName='" & cboitemvs1.Text & "'", CON
        If rs.EOF = True Then
            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
                frmAddMaster.Show 1
                iitem1 = cboitemvs1.Text
                saveInMaster
                cboitemvs1.Text = ""
                Vs1Frame.Visible = False
                vs1.SetFocus
             End If
        End If
        vs1.SetFocus
     ElseIf KeyCode = 27 Then
        Vs1Frame.Visible = False
     End If
End Sub


Private Sub cboItemVs2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        
        
        'cellposiVs2
        vs3.TextMatrix(vs3.RowSel, 0) = cboItemVs2.Text
        Set rs = New ADODB.Recordset
        If rs.State = 1 Then rs.Close
        rs.Open "select * from ItemMaster where " & stringyear & " and  ItemName='" & cboItemVs2.Text & "'", CON
        If rs.EOF = True Then
            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
                frmAddMaster.Show 1
                iitem1 = cboItemVs2.Text
                saveInMaster
                
                cboItemVs2.Text = ""
             End If
        End If
        vs3.SetFocus
        
ElseIf KeyCode = 27 Then
         FrameVs2.Visible = False
End If
End Sub

Private Sub cboItemVs3_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 13 Then
        'cellposiVs3
        
        vs2.TextMatrix(vs2.RowSel, 0) = cboItemVs3.Text
        Set rs = New ADODB.Recordset
        If rs.State = 1 Then rs.Close
        rs.Open "select * from ItemMaster where " & stringyear & " and  ItemName='" & cboItemVs3.Text & "'", CON
        If rs.EOF = True Then
            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
                Vs3Frame.Visible = False
                iitem1 = cboItemVs3.Text
                frmAddMaster.Show 1
                saveInMaster
                cboItemVs3.Text = ""
                vs2.SetFocus
             End If
        End If
        Vs3Frame.Visible = False
        'cboItemVs3.Visible = False
        vs2.SetFocus
     ElseIf KeyCode = 27 Then
        Vs3Frame.Visible = False
     End If

End Sub

Private Sub cmdAdd_Click()
 If rs.State = 1 Then rs.Close
 rs.Open "select HeatingNo from IssueMaster where " & stringyear & " and  HeatingDate >=datevalue('" & fromdate.Value & "') and HeatingDate <=datevalue('" & todate.Value & "') order by HeatingNo", CON
 ListHeatingNo.Clear
 If rs.EOF = False Then
    While rs.EOF = False
       ListHeatingNo.AddItem rs(0)
       rs.MoveNext
    Wend
 End If
End Sub
Private Sub cmdDelete_Click()
   
''   If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
''
''      DeleteRecord txtHeating.Text, "HeatingNo", "IssueMaster"
''      DeleteRecord txtHeating.Text, "HeatingNo", "IssueRawMetrial"
''      Call cmdref_Click
''
''   End If
End Sub
Sub DeleteStock()
    
'''Dim rr As New ADODB.Recordset
'''Dim rs_u As New ADODB.Recordset
'''Dim openning As Double
'''
'''
'''
'''
'''
''''================ Issue For Casting
'''
'''
''' If StockFlag = "1" Then
'''
'''    If rs_u.State = 1 Then rs_u.Close
'''    rs_u.Open "select * from Stock where " & stringyear & " and  ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
'''    If rs_u.EOF = False Then
'''        rs_u!qty = rs_u!qty + qty
'''        rs_u.Update
'''    End If
'''
''' End If
'''
'''
'''
''' '================ Receive For Casting
'''
''' If StockFlag = "2" Then
'''
'''   If itemgp <> "Losses" Then
'''
'''        If rs_u.State = 1 Then rs_u.Close
'''        rs_u.Open "select * from Stock where " & stringyear & " and  ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
'''        If rs_u.EOF = False Then
'''            rs_u!qty = rs_u!qty - qty
'''            rs_u.Update
'''        End If
'''
'''    End If
'''
''' End If
'''
'''
'''
''' '================ Receive For Finish
'''
''' If StockFlag = "3" Then
'''
'''   If itemgp <> "Losses" Then
'''
'''        If rs_u.State = 1 Then rs_u.Close
'''        rs_u.Open "select * from Stock where " & stringyear & " and  ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
'''        If rs_u.EOF = False Then
'''             rs_u!qty = rs_u!qty - qty
'''            rs_u.Update
'''        End If
'''
'''    End If
'''
''' End If
'''
'''
'''
''' '================ Issue For Finish
'''
''' If StockFlag = "4" Then
'''
'''   If itemgp <> "Losses" Then
'''
'''        If rs_u.State = 1 Then rs_u.Close
'''        rs_u.Open "select * from Stock where " & stringyear & " and  ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
'''        If rs_u.EOF = False Then
'''             rs_u!qty = rs_u!qty + qty
'''            rs_u.Update
'''        End If
'''
'''    End If
'''
''' End If
'''
''' '====================================
'''
'''
'''
'''
'''
    
   
   
    
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdFatch_Click()
'AddSemifinish
'Total4

End Sub

Private Sub cmdFind_Click()
 Frame1.Visible = True
 fromdate.SetFocus
End Sub

Private Sub cmdModify_Click()
'   If MsgBox("Want To Modify ?", vbQuestion + vbYesNo) = vbYes Then
'
'
'      DeleteRecord txtHeating.Text, "HeatingNo", "IssueMaster"
'      DeleteRecord txtHeating.Text, "HeatingNo", "IssueRawMetrial"
'
'
'
'      'SaveData
'
'      UpdateIssue
'
'      Call cmdref_Click
'   End If
End Sub
Sub UpdateIssue()

Dim rss As New ADODB.Recordset
Dim search As New ADODB.Recordset
    
If search.State = 1 Then search.Close
search.Open "select ItemName,qty from Invoice where " & stringyear & " and  HeatNo='" & txtHeating.Text & "'", CON
If search.EOF = False Then
While search.EOF = False

    If rss.State = 1 Then rss.Close
    rss.Open "select * from IssueRawMetrial where " & stringyear & " and  HeatingNo=" & txtHeating.Text & " and ItemName='" & search.Fields(0).Value & "'", CON, adOpenDynamic, adLockOptimistic
    If rss.EOF = False Then
       rss.Fields("Issue").Value = (CDbl(rss.Fields("Issue").Value) + CDbl(search.Fields("qty").Value))
       rss.Update
    End If
    
    search.MoveNext
    
Wend
  
End If
  
End Sub

Private Sub cmdRef_Click()
      txtHeating.Text = ""
      txtparty.Text = ""
      
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
Private Sub cmdSave_Click()
    
    
    
    
    If rs.State = 1 Then rs.Close
    rs.Open "select * from IssueMaster where " & stringyear & " and  HeatingNo=" & txtHeating.Text & "", CON
    If rs.EOF = False Then
       MsgBox "Heating No. Already Exist !!", vbInformation
       Exit Sub
    End If
    
    If txtHeating.Text = "" Then
       MsgBox "Please Enter Heating No !!", vbCritical
       txtHeating.SetFocus
       Exit Sub
    End If
    
    
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "select * from IssueMaster where " & stringyear & " and  HeatingNo=" & txtHeating.Text & "", CON
    If rs.EOF = True Then
       If MsgBox("Want To Save ?", vbQuestion + vbYesNo) = vbYes Then
        '  SaveData
       End If
    Else
          MsgBox "Dublicate Heating No !!", vbCritical
    End If
End Sub
Sub ItemGpSearch(str As String)
    
    If rs1.State = 1 Then rs1.Close
    rs1.Open "select ItemGp,Rate from ItemMaster where " & stringyear & " and  ItemName='" & str & "'", CON
    If rs1.EOF = False Then
       itemgp = rs1.Fields(0).Value
       rates = rs1.Fields(1).Value
    End If
    
End Sub
Sub UpdateStock()
    Dim rr As New ADODB.Recordset
    Dim rs_u As New ADODB.Recordset
    Dim openning As Double
    
 
    
    
 '================ Issue For Casting
 
 
 If StockFlag = "1" Then
    
    If rs_u.State = 1 Then rs_u.Close
    rs_u.Open "select * from Stock where " & stringyear & " and  ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
    If rs_u.EOF = True Then
        rs_u.addNew
        rs_u!itemname = Item_Name
        ItemGpSearch Item_Name
        rs_u!itemgp = itemgp
        rs_u!unit = unit
        rs_u!rate = rates
        rs_u!qty = (-1 * qty)
        rs_u.Update
     Else
        rs_u!qty = rs_u!qty - qty
        rs_u.Update
    End If
    
 End If
 
 '====================================
 
 
 '================ Receive For Casting
 
 If StockFlag = "2" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.Close
        rs_u.Open "select * from Stock where " & stringyear & " and  ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = True Then
            rs_u.addNew
            rs_u!itemname = Item_Name
            ItemGpSearch Item_Name
            rs_u!itemgp = itemgp
            rs_u!unit = unit
            rs_u!rate = rates
            rs_u!qty = qty
            rs_u.Update
         Else
            rs_u!qty = rs_u!qty + qty
            rs_u.Update
        End If
    
    End If
    
 End If
 
 '====================================
 
 
 '================ Receive For Finish
 
 If StockFlag = "3" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.Close
        rs_u.Open "select * from Stock where " & stringyear & " and  ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = True Then
            rs_u.addNew
            rs_u!itemname = Item_Name
            ItemGpSearch Item_Name
            rs_u!itemgp = itemgp
            rs_u!unit = unit
            rs_u!rate = rates
            rs_u!qty = qty
            rs_u.Update
         Else
            rs_u!qty = rs_u!qty + qty
            rs_u.Update
        End If
    
    End If
    
 End If
 
 '====================================
 
 
  '================ Issue For Finish
 
 If StockFlag = "4" Then
    
   If itemgp <> "Losses" Then
    
        If rs_u.State = 1 Then rs_u.Close
        rs_u.Open "select * from Stock where " & stringyear & " and  ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
        If rs_u.EOF = True Then
            rs_u.addNew
            rs_u!itemname = Item_Name
            ItemGpSearch Item_Name
            rs_u!itemgp = itemgp
            rs_u!unit = unit
            rs_u!rate = rates
            rs_u!qty = (-1 * qty)
            rs_u.Update
         Else
            rs_u!qty = rs_u!qty - qty
            rs_u.Update
        End If
    
    End If
    
 End If
 
 '====================================
    
    
    
    
End Sub
 
 
     
  


Private Sub Check_manullay_Click()
  txtino.Enabled = True
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   txtForCustom.Visible = True
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

'Check_Nontaxable.Value = False

'ptype.Caption = ""

a = "AAEFC-7614G"

dateInv.Text = Format(Date, "dd/MM/yyyy")

dateInv.SetFocus
txtCenteral = "4/2006 CE Dt. 01/03/06 S.No.97"
txtModePay = "BY Road"
txtIECode = "0505055317"
txtExporterPan = "AAEFC-7614G"
txtEPSG.Text = "0530142551/5/11/00 Dt. 04/12/06"

txtCountryOrigin = "India"
   
   vs.Clear
   setwidth
   cmdDelete_3.Enabled = False
   cmdEdit_4.Enabled = False
   cmdprint.Enabled = False
   
'   cmdPrint_7.Enabled = False
   cmdSave_2.Enabled = True
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
   
  If Check_manullay.Value = 0 Then
   txtino.Text = MaxSNo_Export("Casha", "PACKINGNO")
  End If
   
   
   txtino.Enabled = False
   edit = False
   
   
End Sub
Function MaxSNo_Export(tbl As String, fld As String) As Double
    Dim rs As New Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "Select max(cast(" & fld & " as int)) from " & tbl & " where " & stringyear & " and  typeofinvoice = 'Packing'", CON
    If IsNull(rs(0)) Then
        MaxSNo_Export = 1
    Else
        MaxSNo_Export = Val(rs(0)) + 1
    End If
    rs.Close
End Function


Private Sub cmdAdd1_Click()
K1 = 0
For i = 1 To vs.Rows - 1
If vs.TextMatrix(i, 2) <> "" Then
K1 = K1 + 1
End If

Next




K = 1



vs.Rows = vs.Rows + 1

For i = 1 To vs.Rows - 1
     
If vs.RowSel <= K1 Then
     


If vs.TextMatrix(i, 2) <> "" Then
     
    
    vs.TextMatrix(K1 + 1, 1) = vs.TextMatrix(K1, 1)
    vs.TextMatrix(K1 + 1, 2) = vs.TextMatrix(K1, 2)
    vs.TextMatrix(K1 + 1, 3) = vs.TextMatrix(K1, 3)
    vs.TextMatrix(K1 + 1, 4) = vs.TextMatrix(K1, 4)
    vs.TextMatrix(K1 + 1, 5) = vs.TextMatrix(K1, 5)
    vs.TextMatrix(K1 + 1, 6) = vs.TextMatrix(K1, 6)
    vs.TextMatrix(K1 + 1, 7) = vs.TextMatrix(K1, 7)
    vs.TextMatrix(K1 + 1, 8) = vs.TextMatrix(K1, 8)
    vs.TextMatrix(K1 + 1, 9) = vs.TextMatrix(K1, 9)
    vs.TextMatrix(K1 + 1, 10) = vs.TextMatrix(K1, 10)
    

    K1 = K1 - 1

End If



End If

Next


vs.TextMatrix(vs.RowSel, 1) = ""
vs.TextMatrix(vs.RowSel, 2) = ""
vs.TextMatrix(vs.RowSel, 3) = ""
vs.TextMatrix(vs.RowSel, 4) = ""
vs.TextMatrix(vs.RowSel, 5) = ""
vs.TextMatrix(vs.RowSel, 6) = ""
vs.TextMatrix(vs.RowSel, 7) = ""
vs.TextMatrix(vs.RowSel, 8) = ""
vs.TextMatrix(vs.RowSel, 9) = ""
vs.TextMatrix(vs.RowSel, 10) = ""

'vs.TextMatrix(vs.RowSel, 6) = "add"


vs.Editable = flexEDKbdMouse



End Sub

Private Sub cmdDelete_3_Click()
If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
   
   CON.BeginTrans
   CON.Execute "delete from Casha where " & stringyear & " and  PackingNo =" & txtino & ""
   CON.Execute "delete from CashB where " & stringyear & " and  PackingNo =" & txtino & ""
   'CON.Execute "delete from CashC where " & stringyear & " and  PackingNo =" & txtino & ""
   CON.CommitTrans
   
   Call cmdAdd_1_Click
End If
cmdEdit_4.Enabled = False
cmdDelete_3.Enabled = False
End Sub

Private Sub cmdEdit_4_Click()
   cmdDelete_3.Enabled = True
   cmdEdit_4.Enabled = False
   cmdSave_2.Enabled = True
   cmdAdd_1.Enabled = True
   cmdExit_12.Enabled = True
   txtino.Enabled = False
   edit = True
End Sub

Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub cmdPrint_7_Click()

CR.Reset
CR.ReportFileName = App.Path & "/Reports/CHALLAN.rpt"
CR.ReplaceSelectionFormula "{Casha.invoiceno}=" & txtHeating.Text & ""
CR.WindowShowPrintSetupBtn = True
CR.WindowState = crptMaximized
CR.Action = 1

End Sub


''Sub searchData()
''
''If rs.State = 1 Then rs.Close
''rs.Open "select * from Casha where " & stringyear & " and  INVOICENO=" & txtHeating.Text & "", CON, adOpenDynamic, adLockOptimistic
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
''rs.Open "select * from CashB where " & stringyear & " and  INVOICENO=" & txtHeating.Text & "", CON, adOpenDynamic, adLockOptimistic
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
If KeyCode = 13 Then txtparty.SetFocus
End Sub

Private Sub cmdPrint_Click()
    
If Check1.Value = 0 Then
    CR.Reset
    CR.Connect = constr
    CR.ReportFileName = App.Path & "\REPORTS\PackingList.rpt"
    CR.ReplaceSelectionFormula "{ExportINVQry.packingno} = '" & txtino & "'"
    CR.WindowState = crptMaximized
    CR.WindowShowPrintSetupBtn = True
    CR.Action = 1
Else
    CR.Reset
    CR.Connect = constr
    CR.ReportFileName = App.Path & "\REPORTS\PackingList_forcustom.rpt"
    CR.ReplaceSelectionFormula "{ExportINVQry.packingno} = '" & txtino & "'"
    CR.WindowState = crptMaximized
    CR.WindowShowPrintSetupBtn = True
    CR.Action = 1

End If


End Sub

Private Sub cmdSave_2_Click()


On Error GoTo save:

Dim n As Date
Dim i As Integer
Dim netrate As String
Dim with_without As String




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

If txtExportCurrency.Text = "" Then
  MsgBox "Select Exporter's Currency ...", vbCritical
  txtExportCurrency.SetFocus
  Exit Sub
End If



If txtparty.Text = "" Then
  MsgBox "Enter Party Name ...", vbCritical
  txtparty.SetFocus
  Exit Sub
End If



Dim nontax As String




i = 1
'
If edit = False Then

            If Check_manullay.Value = 0 Then
               'txtino = MaxSNo("Casha", "PACKINGNO")
               txtino.Text = MaxSNo_Export("Casha", "PACKINGNO")
            End If
            
            
            CON.BeginTrans

            CON.Execute "exec Packing_Casha " & txtino & ",'" & dateInv & "','" & dateIssue & "'," & _
            "'" & dateDispatch & "','" & txtparty.Text & "','" & txtIECode.Text & "','" & txtExporterPan.Text & "','" & txtExportCurrency.Text & "'," & _
            "'" & txtEPSG.Text & "','" & txtBuyerOrderNo.Text & "','" & txtBuyerBatchNo.Text & "','" & txtBrand.Text & "','" & txtTermsOfDelivery.Text & "','" & txtTermsOfPayment.Text & "'," & _
            "'" & txtFDestination.Text & "','" & txtPreCarriageBy.Text & "','" & txtPlaceofPrecarriage.Text & "','" & txtPortLoading.Text & "','" & txtPortDischarge.Text & "'," & _
            "'" & txtTotalPSC.Text & "','" & txtTotalCartoons.Text & "','" & txtCBM.Text & "','" & txtNetW.Text & "','" & txtGwt.Text & "'," & _
            "" & "'Packing','" & main.username & "','" & main.username & "','" & main.session & "'," & _
            "'" & main.setupid & "'"
            
            CON.Execute "update casha set station='" & txtForCustom & "',invoicedate='" & Format(dateInv, "MM/dd/yyyy") & "',remark='" & txtShipToadd & "' where " & stringyear & " and  PACKINGNO=" & txtino.Text & ""
            
            
            For i = 1 To vs.Rows - 1
            If (vs.TextMatrix(i, 1) <> "" And vs.TextMatrix(i, 2) <> "") Then

            CON.Execute "insert into Cashb(PACKINGNO,SNO,Carton,itemname,BOOKCODE,btno,NetWt,GWt,InnerPack,OuterPack,QUANTITY,Size,TotalCBM2,typeofinvoice,NoOfBox,Per_Box_GW,Per_Box_NW)" & _
            " values(" & txtino.Text & ",'" & (vs.TextMatrix(i, 0)) & "','" & vs.TextMatrix(i, 4) & "','" & vs.TextMatrix(i, 5) & "','" & vs.TextMatrix(i, 6) & "'," & _
            "'" & vs.TextMatrix(i, 7) & "','" & vs.TextMatrix(i, 8) & "','" & vs.TextMatrix(i, 9) & "'," & _
            "'" & Val(vs.TextMatrix(i, 10)) & "','" & Val(vs.TextMatrix(i, 11)) & "'," & Val(vs.TextMatrix(i, 12)) & ",'" & (vs.TextMatrix(i, 13)) & _
            "','" & Val(vs.TextMatrix(i, 14)) & "','Packing','" & Val(vs.TextMatrix(i, 1)) & "','" & Val(vs.TextMatrix(i, 2)) & "','" & Val(vs.TextMatrix(i, 3)) & "')"
                      
            Else
            
                Exit For
            
            End If
            Next
            
            
            
            
            CON.CommitTrans
            
    Else
            
            
            
            CON.BeginTrans

            
            CON.Execute "delete from CashB where " & stringyear & " and  PACKINGNO=" & txtino.Text & ""


             CON.Execute "update Casha set PACKINGNO='" & txtino & "',IssueDate='" & dateIssue & "',DisPatchDate=" & _
            "'" & dateDispatch & "',SUBLEDGER='" & txtparty.Text & "',TXT1A='" & txtIECode.Text & "',TXT2='" & txtExporterPan.Text & "',TXT2A='" & txtExportCurrency.Text & "',T2=" & _
            "'" & txtEPSG.Text & "',ORDERNO='" & txtBuyerOrderNo.Text & "',AdviceRemark='" & txtBuyerBatchNo.Text & "',Brand='" & txtBrand.Text & "',TermsDelevery='" & txtTermsOfDelivery.Text & "',TermsPayment='" & txtTermsOfPayment.Text & "',COUNTRYDEST=" & _
            "'" & txtFDestination.Text & "',PRECARRIAGE='" & txtPreCarriageBy.Text & "',PLACERECEIPT='" & txtPlaceofPrecarriage.Text & "',PORTLOADING='" & txtPortLoading.Text & "',PORTDISCHARGE='" & txtPortDischarge.Text & "',TOTALPSC=" & _
            "'" & txtTotalPSC.Text & "',TotalCartons='" & txtTotalCartoons.Text & "',TotalCBM='" & txtCBM.Text & "',NetWeight='" & txtNetW.Text & "',GWeight='" & txtGwt.Text & _
            "',updatedby='" & main.username & "',updatedon=" & Date & " where " & stringyear & " and  PACKINGNO=" & txtino.Text & ""
            CON.Execute "update casha set station='" & txtForCustom & "',invoicedate='" & Format(dateInv, "MM/dd/yyyy") & "',remark='" & txtShipToadd & "' where " & stringyear & " and  PACKINGNO=" & txtino.Text & ""
            
            
                    
            For i = 1 To vs.Rows - 1
            If (vs.TextMatrix(i, 1) <> "" And vs.TextMatrix(i, 2) <> "") Then

            CON.Execute "insert into Cashb(PACKINGNO,SNO,Carton,itemname,BOOKCODE,btno,NetWt,GWt,InnerPack,OuterPack,QUANTITY,Size,TotalCBM2,typeofinvoice,NoOfBox,Per_Box_GW,Per_Box_NW)" & _
            " values(" & txtino.Text & ",'" & (vs.TextMatrix(i, 0)) & "','" & vs.TextMatrix(i, 4) & "','" & vs.TextMatrix(i, 5) & "','" & vs.TextMatrix(i, 6) & "'," & _
            "'" & vs.TextMatrix(i, 7) & "','" & vs.TextMatrix(i, 8) & "','" & vs.TextMatrix(i, 9) & "'," & _
            "'" & Val(vs.TextMatrix(i, 10)) & "','" & Val(vs.TextMatrix(i, 11)) & "'," & Val(vs.TextMatrix(i, 12)) & ",'" & (vs.TextMatrix(i, 13)) & _
            "','" & Val(vs.TextMatrix(i, 14)) & "','Packing','" & Val(vs.TextMatrix(i, 1)) & "','" & Val(vs.TextMatrix(i, 2)) & "','" & Val(vs.TextMatrix(i, 3)) & "')"
                                  
            End If
                                  
            Next
            
            
            
            
            CON.CommitTrans
                    
            
            
            

End If
MsgBox "Data Saved..", vbInformation

cmdprint.Enabled = True





Check_manullay.Value = 0


cmdEdit_4.Enabled = True
cmdDelete_3.Enabled = True
'Call cmdAdd_1_Click

Exit Sub
save:
CON.RollbackTrans
MsgBox err.DESCRIPTION
'txtCode.SetFocus




End Sub
Private Sub totalvalue()

txtTotalPSC = "0.00"
txtTotalCartoons = "0.00"
txtCBM = "0.00"
txtNetW = "0.00"
txtGwt = "0.00"

For i = 1 To vs.Rows - 1
If vs.TextMatrix(i, 12) <> "" Then

txtTotalCartoons.Text = (Val(txtTotalCartoons.Text) + Val(vs.TextMatrix(i, 1)))
txtTotalPSC = (txtTotalPSC + Val(vs.TextMatrix(i, 12)))
txtCBM = Val(txtCBM) + Val(vs.TextMatrix(i, 14))
txtNetW = Val(txtNetW) + Val(vs.TextMatrix(i, 8))
txtGwt = Val(txtGwt) + Val(vs.TextMatrix(i, 9))
End If

Next

txtCBM = Format((txtCBM), "0.00")
txtNetW = Format(Round(txtNetW, 3), "0.000")
txtGwt = Format(Round(txtGwt, 3), "0.000")


End Sub
Private Sub cmdSearch_Click()

 popuplist10 "select PackingNo,[SUBLEDGER] as [Party Name],[INVOICEDATE] as [Date] from Casha where " & stringyear & " order by cast(PackingNo as int)", CON

''If txtino.Enabled = False Then
''txtino.Enabled = True
''txtino.Text = ""
''txtino.SetFocus
''End If
''
''If txtino.Text = "" Then
'''MsgBox "Enter Invoice No...", vbInformation
''txtino.SetFocus
''Exit Sub
''End If
''SearchData

End Sub
Private Sub cmdSearch_GotFocus()
   If PopUpValue1 <> "" Then
      txtino = PopUpValue1
      searchData
      
      
      PopUpValue1 = ""
      PopUpValue2 = ""
      PopUpValue3 = ""
   End If
End Sub

Private Sub cmdUpDateGrid_Click()
UpDateGrid
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
 
 
 
 
 
 If (UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("vs")) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("txtForCustom"))) And UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("txtShipToAdd")) Then
    SendKeys "{tab}"
    HIT
 End If
 
 End If
 


 
 
End Sub
Sub searchData()

On Error Resume Next

vs.Clear
setwidth


If rs.State = 1 Then rs.Close

st1 = "select PACKINGNO,INVOICEDATE,IssueDate,DisPatchDate," & _
"TXT1A,TXT2,TXT2A,T2,SUBLEDGER,ORDERNO,ADVICEREMARK,Brand,TermsDelevery,TermsPayment,COUNTRYDEST,PRECARRIAGE," & _
"PLACERECEIPT,PORTLOADING,PORTDISCHARGE,TOTALPSC,TOTALCARTONS,TOTALCBM,NETWEIGHT,GWEIGHT,Station,remark" & _
" from Casha where " & stringyear & " and  PACKINGNO=" & txtino & " and typeofinvoice='packing'"
     
rs.Open st1, CON
If rs.EOF = False Then
  
  
cmdSave_2.Enabled = False
cmdDelete_3.Enabled = False
cmdAdd_1.Enabled = True
cmdprint.Enabled = True
cmdEdit_4.Enabled = True
txtino.Enabled = False
  
  
  
  txtShipToadd = rs!remark & ""
  
  dateInv = rs!InvoiceDate
  dateIssue = Format(rs!IssueDate, "dd/MM/yyyy")
  If IsDate(rs!DisPatchDate) Then
     dateDispatch = Format(rs!DisPatchDate, "dd/MM/yyyy")
  End If
  
  Check1.Value = 0
  If Not IsNull(rs!station) Then
    If Len(rs!station) > 0 Then
       txtForCustom = rs!station
       Check1.Value = 1
    End If
  Else
       
       
       txtForCustom = ""
       txtForCustom.Visible = False
  End If
  
  
  txtIECode = rs!txt1a & ""
  txtExporterPan = rs!txt2 & ""
  txtExportCurrency = rs!txt2a
  txtEPSG = rs!T2
  txtparty = rs!subledger
  txtBuyerOrderNo = rs!ORDERNO
  txtBuyerBatchNo = rs!ADVICEREMARK
  txtBrand = rs!Brand
  txtTermsOfDelivery = rs!TermsDelevery
  txtTermsOfPayment = rs!TermsPayment
  txtFDestination = rs!COUNTRYDEST
  txtPreCarriageBy = rs!PRECARRIAGE
  txtPlaceofPrecarriage = rs!PLACERECEIPT
  txtPortLoading = rs!PORTLOADING
  txtPortDischarge = rs!PORTDISCHARGE
  txtTotalPSC = rs!TOTALPSC
  txtTotalCartoons = rs!TOTALCARTONS
  txtCBM = rs!TOTALCBM
  txtNetW = rs!NETWEIGHT
  txtGwt = rs!GWEIGHT
    
If rs.State = 1 Then rs.Close
rs.Open "select ADDRESS1,ADDRESS2 from SLEDGER where " & stringyear & " and  SUBLEDGER = '" & txtparty.Text & "'", CON
Dim addr As String
addr = rs!address1 & "" & vbCrLf & "" & rs!ADDRESS2
txtAdd.Text = addr



    



If rs.State = 1 Then rs.Close
rs.Open "select INVOICENO,SNO,Carton,itemname,BOOKCODE,btno,NetWt,GWt,InnerPack,OuterPack,QUANTITY,SIZE,TotalCBM2,NoOfBox,Per_Box_GW,Per_Box_NW" & _
" from CashB where " & stringyear & " and  PACKINGNO=" & txtino.Text & " and typeofinvoice='packing' order by sno", CON
For i = 1 To rs.RecordCount
vs.TextMatrix(i, 0) = rs!sno


vs.TextMatrix(i, 1) = rs!NoOfBox
vs.TextMatrix(i, 2) = rs!Per_Box_GW
vs.TextMatrix(i, 3) = rs!Per_Box_NW


vs.TextMatrix(i, 4) = rs!Carton
vs.TextMatrix(i, 5) = rs!itemname
vs.TextMatrix(i, 6) = rs!Bookcode
vs.TextMatrix(i, 7) = rs!btno
vs.TextMatrix(i, 8) = rs!NetWt
vs.TextMatrix(i, 9) = rs!Gwt
vs.TextMatrix(i, 10) = rs!InnerPack
vs.TextMatrix(i, 11) = rs!OuterPack
vs.TextMatrix(i, 12) = rs!quantity
vs.TextMatrix(i, 13) = rs!Size
vs.TextMatrix(i, 14) = rs!TotalCBM2


rs.MoveNext
Next

End If



totalvalue

End Sub
'Sub TotalFinal()
'   If txtTotal3.Text = "" Then
'      txtTotal3.Text = 0
'   End If
'
'   If txtTotal2.Text = "" Then
'      txtTotal2.Text = 0
'   End If
'
'
'    txtRawAndCasting.Text = (CDbl(txtTotal2.Text) + CDbl(txtTotal3.Text))
'    txtRawAndCasting.Text = Format(txtRawAndCasting.Text, "#,###.000")
'End Sub
Private Sub Form_Load()
 
 setwidth
 
 
 
 'txtino = MaxSNo("Casha", "PACKINGNO")
  txtino.Text = MaxSNo_Export("Casha", "PACKINGNO")
  
 'txtRate(0) = VAT
 
 
cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = False
cmdprint.Enabled = False

ButtonPermission cmdSave_2, cmdDelete_3, cmdEdit_4
 
End Sub
Sub setwidth()
vs.Cols = 10
vs.FormatString = "<S. No.|NoOfBox|GW_Per_Box|NW_Per_Box|Carton|<Description of Goods|^Product Code|^Batch No.|^Net Wt.|^Gross Wt.|^Inner Pack|^Outer|>Quantity|>Size(cm)|>CBM"
vs.ColWidth(0) = 300
vs.ColWidth(1) = 900
vs.ColWidth(2) = 1200
vs.ColWidth(3) = 1200

vs.ColWidth(4) = 1100
vs.ColWidth(5) = 4000
vs.ColWidth(6) = 1400



vs.ColWidth(7) = 1000
vs.ColWidth(8) = 1000
vs.ColWidth(9) = 1000

vs.ColWidth(10) = 1000
vs.ColWidth(11) = 1000

vs.ColWidth(12) = 1000
vs.ColWidth(13) = 1800
vs.ColWidth(14) = 1500



End Sub
Private Sub fromdate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then todate.SetFocus
End Sub
Private Sub HeatingDate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then txtparty.SetFocus
End Sub

Private Sub ListHeatingNo_Click()
  Call cmdRef_Click
  searchData
'  TotalFinal
  'Frame1.Visible = False
End Sub
Private Sub todate_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then Call cmdAdd_Click
End Sub
Private Sub txtGrade_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then txtRemarks.SetFocus
End Sub

Private Sub txtHeating_GotFocus()
If PopUpValue1 <> "" Then
txtHeating.Text = PopUpValue1
Dates.Value = PopUpValue2
vs.Clear
setwidth
searchData
PopUpValue1 = ""
PopUpValue2 = ""
End If
End Sub

Private Sub txtHeating_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 113 Then
'popuplist2 "select INVOICENO,INVOICEDATE,SUBLEDGER from Casha order by INVOICENO", CON
'End If
'
'If KeyCode = 13 Then
'SearchData
'End If

End Sub

Private Sub txtHeating_KeyPress(KeyAscii As Integer)
   
   If KeyAscii = 13 Then
        
   Dates.SetFocus

        
  End If
  

End Sub


Private Sub Option_with_Click()
'If Option_with.Value = True Then
'  txtRate(0) = withForm
'Else
'  txtRate(0) = withoutForm
'End If
FatchTaxFromSate
End Sub

Private Sub Option_without_Click()
'If Option_with.Value = True Then
'  txtRate(0) = withForm
'Else
'  txtRate(0) = withoutForm
'End If

FatchTaxFromSate
End Sub


Private Sub txtamount_Change(Index As Integer)
''On Error Resume Next
''
''total1 = (Val(txtTotal) - Val(txtamount(1)))
''
''txtamount(0) = Round((total1 * Val(txtRate(0)) / 100), 3)
''txtamount(0) = Format(txtamount(0), ".00")
''
''txtNet = Format((Val(total1) + Val(txtamount(0))), ".00")
''
End Sub

Private Sub txtamount_LostFocus(Index As Integer)
''txtamount(1) = Format(txtamount(1), ".00")
End Sub









Private Sub txtBrand_LostFocus()
  txtBrand = UCase(txtBrand)
End Sub

Private Sub txtBuyerBatchNo_LostFocus()
txtBuyerBatchNo = UCase(txtBuyerBatchNo)
End Sub

Private Sub txtExportCurrency_LostFocus()
If txtExportCurrency.Text = "" Then
MsgBox "Please select export currency...", vbInformation
txtExportCurrency.SetFocus
End If
End Sub

Private Sub txtFDestination_LostFocus()
  txtFDestination = UCase(txtFDestination)
End Sub


Private Sub txtForCustom_LostFocus()
txtForCustom = UCase(txtForCustom)
End Sub

Private Sub txtino_GotFocus()
HIT
End Sub

Private Sub txtino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  searchData
End If
End Sub

Private Sub txtparty_GotFocus()
HIT

If PopUpValue1 <> "" Then
      
   txtparty = PopUpValue1
   txtAdd.Text = PopUpValue2 & vbCrLf & PopUpValue3 & vbCrLf & popupvalue4
   txtBuyerOrderNo.SetFocus
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   
End If

End Sub
Private Sub txtParty_KeyUp(KeyCode As Integer, Shift As Integer)
''If (KeyCode = 13 Or KeyCode = 144) Then Exit Sub
''tblNo = 1
''frmSearchItem.Show
popuplist10 "select Subledger,Address1,Address2,Country from Sledger where " & stringyear & " order by Subledger", CON

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



Private Sub txtPlaceofPrecarriage_LostFocus()
txtPlaceofPrecarriage = UCase(txtPlaceofPrecarriage)
End Sub



Private Sub txtPortDischarge_LostFocus()
txtPortDischarge = UCase(txtPortDischarge)
End Sub

Private Sub txtPortLoading_Change()
txtPortLoading = UCase(txtPortLoading)
End Sub

Private Sub txtPreCarriageBy_LostFocus()
  txtPreCarriageBy = UCase(txtPreCarriageBy)
End Sub

Private Sub txtShipToAdd_GotFocus()
HIT

If PopUpValue1 <> "" Then
      
   txtShipToadd = PopUpValue1 & vbCrLf & PopUpValue2 & vbCrLf & PopUpValue3
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   vs.SetFocus
End If

End Sub

Private Sub txtShipToAdd_KeyUp(KeyCode As Integer, Shift As Integer)

popuplist10 "select Subledger,Address1,Address2 from Sledger where " & stringyear & " order by Subledger", CON

End Sub


Private Sub txtShipToAdd_LostFocus()
txtShipToadd = UCase(txtShipToadd)
End Sub

Private Sub txtTermsOfDelivery_LostFocus()
txtTermsOfDelivery = UCase(txtTermsOfDelivery)
End Sub


Private Sub txtTermsOfPayment_LostFocus()
txtTermsOfPayment = UCase(txtTermsOfPayment)
End Sub

Private Sub vs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
''     If vs.col = 0 Then
''        cellposi
''        vs.TextMatrix(vs.RowSel, 0) = cboItem.Text
''     End If
End Sub
Sub addNew()
K1 = 0
For i = 1 To vs.Rows - 1
If vs.TextMatrix(i, 2) <> "" Then
K1 = K1 + 1
End If

Next




K = 1



vs.Rows = vs.Rows + 1

For i = 1 To vs.Rows - 1
     
If vs.RowSel <= K1 Then
     


If vs.TextMatrix(i, 2) <> "" Then
     
    
    vs.TextMatrix(K1 + 1, 1) = vs.TextMatrix(K1, 1)
    vs.TextMatrix(K1 + 1, 2) = vs.TextMatrix(K1, 2)
    vs.TextMatrix(K1 + 1, 3) = vs.TextMatrix(K1, 3)
    vs.TextMatrix(K1 + 1, 4) = vs.TextMatrix(K1, 4)
    vs.TextMatrix(K1 + 1, 5) = vs.TextMatrix(K1, 5)
    vs.TextMatrix(K1 + 1, 6) = vs.TextMatrix(K1, 6)
    vs.TextMatrix(K1 + 1, 7) = vs.TextMatrix(K1, 7)
    vs.TextMatrix(K1 + 1, 8) = vs.TextMatrix(K1, 8)
    vs.TextMatrix(K1 + 1, 9) = vs.TextMatrix(K1, 9)
    vs.TextMatrix(K1 + 1, 10) = vs.TextMatrix(K1, 10)
    
    vs.TextMatrix(K1 + 1, 11) = vs.TextMatrix(K1, 11)
    vs.TextMatrix(K1 + 1, 12) = vs.TextMatrix(K1, 12)
    vs.TextMatrix(K1 + 1, 13) = vs.TextMatrix(K1, 13)
    
    

    K1 = K1 - 1

End If



End If

Next


vs.TextMatrix(vs.RowSel, 1) = ""
vs.TextMatrix(vs.RowSel, 2) = ""
vs.TextMatrix(vs.RowSel, 3) = ""
vs.TextMatrix(vs.RowSel, 4) = ""
vs.TextMatrix(vs.RowSel, 5) = ""
vs.TextMatrix(vs.RowSel, 6) = ""
vs.TextMatrix(vs.RowSel, 7) = ""
vs.TextMatrix(vs.RowSel, 8) = ""
vs.TextMatrix(vs.RowSel, 9) = ""
vs.TextMatrix(vs.RowSel, 10) = ""

'vs.TextMatrix(vs.RowSel, 6) = "add"


vs.Editable = flexEDKbdMouse

End Sub
Private Sub VS_KeyDown(KeyCode As Integer, Shift As Integer)

  
  
  If KeyCode = 115 Then
  If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
    vs.RemoveItem (vs.RowSel)
    Total
  End If
  End If
  
  If KeyCode = 116 Then
  If MsgBox("Want to Add New Raw ?", vbQuestion + vbYesNo) = vbYes Then
     addNew
  End If
  End If
  
  
  
  
  
  If KeyCode = 13 Then
     
     If vs.Col = 0 Then
        vs.Editable = flexEDNone
        VsFrame.Visible = True
        cboItem.SetFocus
     Else
        vs.Editable = flexEDKbdMouse
        cellposi
     End If

  End If
  
  
  
  
  
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)

On Error Resume Next

Dim Item As String
Dim totalcarton, D1_D2




If KeyCode = 13 Then
        
 
 If vs.Col = 1 Then
    
        If vs.TextMatrix(vs.RowSel, 1) <> "" Then
         SendKeys "{right}"
        End If
    
    
 ElseIf vs.Col = 2 Then
 
 If vs.TextMatrix(vs.RowSel, 2) <> "" Then
   SendKeys "{right}"
 End If

 
 ElseIf vs.Col = 3 Then
 
 If vs.TextMatrix(vs.RowSel, 3) <> "" Then
   SendKeys "{right}"
   SendKeys "{right}"
   
   
   If vs.Row = 1 Then
      vs.TextMatrix(vs.RowSel, 4) = "1-" & vs.TextMatrix(vs.RowSel, 1)
   Else
      totalcarton = Split(vs.TextMatrix((vs.RowSel - 1), 4), "-")
      vs.TextMatrix(vs.RowSel, 4) = totalcarton(1) & "-" & vs.TextMatrix(vs.RowSel, 1)
   End If
   
   
    vs.TextMatrix(vs.RowSel, 8) = Val(vs.TextMatrix(vs.RowSel, 1)) * Val(vs.TextMatrix(vs.RowSel, 3))
    vs.TextMatrix(vs.RowSel, 9) = Val(vs.TextMatrix(vs.RowSel, 1)) * Val(vs.TextMatrix(vs.RowSel, 2))
   
   
 End If

 
 ElseIf vs.Col = 4 Then
    
 
    If vs.TextMatrix(vs.RowSel, 4) <> "" Then
       K = InStr(1, vs.TextMatrix(vs.RowSel, 4), "-")
       If K <> 0 Then
        totalcarton = Split(vs.TextMatrix((vs.RowSel), 4), "-")
        txtTotalCartoons.Text = totalcarton(1)
         txtTotalCartoons.SetFocus
       End If
       
       
    End If


If vs.TextMatrix(vs.RowSel, 4) <> "" Then
 SendKeys "{right}"
End If
 
 
 
 ElseIf vs.Col = 5 Then
    vs.TextMatrix(vs.RowSel, vs.Col) = UCase(vs.TextMatrix(vs.RowSel, vs.Col))
    SendKeys "{right}"
 ElseIf vs.Col = 6 Then
    vs.TextMatrix(vs.RowSel, vs.Col) = UCase(vs.TextMatrix(vs.RowSel, vs.Col))
    SendKeys "{right}"
 ElseIf vs.Col = 7 Then
    SendKeys "{right}"
 ElseIf vs.Col = 8 Then
    SendKeys "{right}"
 ElseIf vs.Col = 9 Then
    SendKeys "{right}"
 ElseIf vs.Col = 10 Then
    SendKeys "{right}"
 ElseIf vs.Col = 11 Then
    K1 = 0
    
    K1 = InStr(1, vs.TextMatrix(vs.RowSel, 4), "-")
    If K1 <> 0 Then
    D1_D2 = Split(vs.TextMatrix((vs.RowSel), 4), "-")

    d1 = D1_D2(0)
    D2 = D1_D2(1)
    
    If vs.Row <= 1 Then
    d1 = (D2)
    Else
    d1 = ((D2 - d1) + 1)
    End If
    
    
    vs.TextMatrix(vs.RowSel, 12) = (Val(vs.TextMatrix(vs.RowSel, 11)) * d1)
    
    
    
    
 End If
    
    SendKeys "{right}"
    SendKeys "{right}"
 
 ElseIf vs.Col = 12 Then
    SendKeys "{right}"
 ElseIf vs.Col = 13 Then
 
 
 
 
    K1 = 0
    
    K1 = InStr(1, vs.TextMatrix(vs.RowSel, 4), "-")
    If K1 <> 0 Then
    D1_D2 = Split(vs.TextMatrix((vs.RowSel), 4), "-")

    d1 = D1_D2(0)
    D2 = D1_D2(1)
    
    If vs.Row <= 1 Then
    d1 = ((D2 - d1) + 1)
    Else
    d1 = (D2 - d1)
    End If
    End If
 
    
    sizeVal = ""
    v1 = 0
    
    
    s3 = UCase(vs.TextMatrix(vs.RowSel, 13))
    
    K2 = InStr(1, s3, "X")
    If K2 <> 0 Then
    sizeVal = Split(s3, "X")
    If vs.Row = 4 Then
    v1 = ((d1 * (sizeVal(0) * sizeVal(1) * sizeVal(2))) / 1000000)
    Else
    v1 = (((d1 + 1) * (sizeVal(0) * sizeVal(1) * sizeVal(2))) / 1000000)
    End If
    
    End If
    vs.TextMatrix(vs.RowSel, 14) = Round(v1, 2)
    
    vs.TextMatrix(vs.RowSel, 0) = vs.RowSel
    
    SendKeys "{down}"
    SendKeys "{home}"
 Else
    vs.TextMatrix(vs.RowSel, 0) = vs.RowSel
    SendKeys "{down}"
    SendKeys "{home}"
 End If
 
 

Total

totalvalue


End If

End Sub
Sub UpDateGrid()

On Error Resume Next

Dim Item As String
Dim totalcarton, D1_D2
Dim sum1 As Double

sum1 = 0


For i = 1 To vs.Rows - 1


'==========================================================================
If vs.TextMatrix(i, 1) = "" Then
   Exit For
End If


If i = 1 Then
   vs.TextMatrix(i, 4) = "1-" & vs.TextMatrix(i, 1)
Else
   totalcarton = Split(vs.TextMatrix((i - 1), 4), "-")
   'vs.TextMatrix(i, 4) = totalcarton(1) & "-" & (Val(vs.TextMatrix(i, 1)) + Val(totalcarton(1)))
   vs.TextMatrix(i, 4) = totalcarton(1) & "-" & (Val(vs.TextMatrix(i, 1)) + Val(totalcarton(1)))
End If


vs.TextMatrix(i, 8) = Val(vs.TextMatrix(i, 1)) * Val(vs.TextMatrix(i, 3))
vs.TextMatrix(i, 9) = Val(vs.TextMatrix(i, 1)) * Val(vs.TextMatrix(i, 2))


If vs.TextMatrix(i, 4) <> "" Then
   K = InStr(1, vs.TextMatrix(i, 4), "-")
   If K <> 0 Then
     totalcarton = Split(vs.TextMatrix(i, 4), "-")
     txtTotalCartoons.Text = totalcarton(1)
   End If
End If

K1 = 0


'''==========================

K1 = InStr(1, vs.TextMatrix(i, 4), "-")
If K1 <> 0 Then
D1_D2 = Split(vs.TextMatrix(i, 4), "-")

d1 = D1_D2(0)
D2 = D1_D2(1)

    If i <= 1 Then
        d1 = (D2)
      Else
        d1 = ((D2 - d1) + 1)
     End If


     vs.TextMatrix(i, 12) = (Val(vs.TextMatrix(i, 11)) * d1)

End If

''''===================================


K1 = 0

K1 = InStr(1, vs.TextMatrix(i, 4), "-")

If K1 <> 0 Then
   D1_D2 = Split(vs.TextMatrix(i, 4), "-")
   d1 = D1_D2(0)
   D2 = D1_D2(1)

        If i <= 1 Then
                d1 = ((D2 - d1) + 1)
            Else
                d1 = (D2 - d1)
        End If


        sizeVal = ""
        v1 = 0

       s3 = UCase(vs.TextMatrix(i, 13))
       K2 = InStr(1, s3, "X")
        
       If K2 <> 0 Then
            sizeVal = Split(s3, "X")
                If vs.Row = 4 Then
                    v1 = ((d1 * (sizeVal(0) * sizeVal(1) * sizeVal(2))) / 1000000)
                Else
                    v1 = (((d1 + 1) * (sizeVal(0) * sizeVal(1) * sizeVal(2))) / 1000000)
                End If
        End If

End If

vs.TextMatrix(i, 14) = Round(v1, 2)
vs.TextMatrix(i, 0) = i







'======================================
Next

    
    
 Total
totalvalue

 




End Sub






Sub FatchTaxFromSate()

Dim ST As String
Dim with_without As String





If rs.State = 1 Then rs.Close
rs.Open "select State,tinno from SubledgerQry where " & stringyear & " and  subledger='" & txtparty & "'", CON
If rs.EOF = False Then
   ST = rs(0)
If LCase(ST) = "u.p." Then

 If Len(rs!tinno) > 0 Then
   'header = "TAX-INVOICE"
 Else
   'header = "SALE-INVOICE"
 End If
   
Else
   'header = "SALE-INVOICE"
  
End If

End If


If rs.State = 1 Then rs.Close
rs.Open "select add_val,less_val from [state_tax_list] where " & stringyear & " and  statename='" & ST & "' and with_without='" & with_without & "'", CON
If rs.EOF = False Then
   VAT_Add = rs(0)
   VAT_less = rs(1)
   
End If




End Sub


Private Sub vs_LeaveCell()
  Total
End Sub

'''Private Sub vs1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'''     If vs1.Col = 0 Then
'''        cellposiVs
'''        vs1.TextMatrix(vs1.RowSel, 0) = cboitemvs1.Text
'''     End If
'''
'''End Sub
'''Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)
'''  If KeyCode = 46 Then
'''    vs1.RemoveItem (vs1.RowSel)
'''    'Total1
'''    'TotalFinal
'''  End If
'''
'''  If KeyCode = 13 Then
'''     If vs1.Col = 0 Then
'''        vs1.Editable = flexEDNone
'''        Vs1Frame.Visible = True
'''        cboitemvs1.Visible = True
'''        cboitemvs1.SetFocus
'''     Else
'''        vs1.Editable = flexEDKbdMouse
'''        cellposiVs
'''     End If
'''  End If
'''End Sub

''''Private Sub vs1_KeyUp(KeyCode As Integer, Shift As Integer)
''''
''''If KeyCode = 13 Then
''''
'''' If vs1.Col = 0 Then
''''    vs1.Editable = flexEDNone
''''    Vs1Frame.Visible = True
''''    cboitemvs1.SetFocus
''''
''''    Set rs = New ADODB.Recordset
''''    If rs.State = 1 Then rs.Close
''''    rs.Open "select * from ItemMaster where " & stringyear & " and  ItemName='" & vs1.TextMatrix(vs1.RowSel, 3) & "'", CON
''''    If rs.EOF = False Then
''''       vs1.TextMatrix(vs1.RowSel, 1) = rs.Fields("Unit").Value
''''       SendKeys "{right}"
''''       SendKeys "{right}"
''''       Vs1Frame.Visible = False
''''       vs1.Editable = flexEDKbdMouse
''''       vs1.SetFocus
''''    Else
''''       vs1.TextMatrix(vs1.RowSel, 1) = "Kg"
''''       SendKeys "{right}"
''''       SendKeys "{right}"
''''       Vs1Frame.Visible = False
''''       vs1.Editable = flexEDKbdMouse
''''       vs1.SetFocus
''''
''''    End If
''''
'''' End If
''''
'''' If vs1.Col = 2 Then
''''
''''    SendKeys "{home}"
''''    SendKeys "{down}"
''''
''''    AddItemInGrid1
'''' End If
''''
''''
''''
'''' 'Total1
'''' 'TotalFinal
''''
''''End If
''''
''''
''''End Sub
''Sub AddSemifinish()
''   Dim J As Integer
''
''   J = 1
''
''   vs3.Clear
''   For i = 1 To vs1.Rows - 1
''
''   If vs1.TextMatrix(i, 0) <> "" Then
''      If rs.State = 1 Then rs.Close
''      rs.Open "select * from ItemMaster where " & stringyear & " and  ItemName='" & vs1.TextMatrix(i, 0) & "'", CON
''      If rs.Fields("itemgp").Value = "Semi Finish (R/D)" Or rs.Fields("itemgp").Value = "Semi Finish (Store)" Then
''         vs3.TextMatrix(J, 0) = vs1.TextMatrix(i, 0)
''         vs3.TextMatrix(J, 1) = vs1.TextMatrix(i, 1)
''         vs3.TextMatrix(J, 2) = vs1.TextMatrix(i, 2)
''         J = J + 1
''      End If
''   End If
''
''   Next
''
''
''End Sub
''Private Sub vs1_LeaveCell()
''   'Total1
''End Sub
''
''Private Sub vs2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
''     If vs2.Col = 0 Then
''        'cellposiVs3
''        vs2.TextMatrix(vs2.RowSel, 0) = cboItemVs3.Text
''     End If
''
''End Sub
''
''Private Sub vs2_KeyDown(KeyCode As Integer, Shift As Integer)
'' If KeyCode = 46 Then
''    vs2.RemoveItem (vs2.RowSel)
''    'Total2
''    'TotalFinal
'' End If
''
''
''  If KeyCode = 13 Then
''
''     If vs2.Col = 0 Then
''        vs2.Editable = flexEDNone
''        Vs3Frame.Visible = True
''        cboItemVs3.Visible = True
''        cboItemVs3.SetFocus
''     Else
''        vs2.Editable = flexEDKbdMouse
''        'cellposiVs3
''     End If
''
''  End If
''
''End Sub

''Private Sub vs2_KeyUp(KeyCode As Integer, Shift As Integer)
''If KeyCode = 13 Then
''
''
'' If vs2.Col = 0 Then
''
''      vs2.Editable = flexEDNone
''      Vs3Frame.Visible = True
''
''
''
''    Set rs = New ADODB.Recordset
''    If rs.State = 1 Then rs.Close
''    rs.Open "select * from ItemMaster where " & stringyear & " and  ItemName='" & vs2.TextMatrix(vs2.RowSel, 0) & "'", CON
''    If rs.EOF = False Then
''       vs2.TextMatrix(vs2.RowSel, 1) = rs.Fields("Unit").Value
''       SendKeys "{right}"
''       SendKeys "{right}"
''       Vs3Frame.Visible = False
''       vs2.Editable = flexEDKbdMouse
''       vs2.SetFocus
''    Else
''       vs2.TextMatrix(vs2.RowSel, 1) = "Kg"
''       SendKeys "{right}"
''       SendKeys "{right}"
''       Vs3Frame.Visible = False
''       vs2.Editable = flexEDKbdMouse
''       vs2.SetFocus
''
''    End If
''
'' End If
''
''
''    If vs2.Col = 2 Then
''
''           SendKeys "{home}"
''           SendKeys "{down}"
''           Vs3Frame.TOP = Vs3Frame.TOP + 170
''    End If
''
''
''   'Total2
''
''End If
''
''End Sub
''Private Sub vs2_LeaveCell()
''   'Total2
''End Sub
''Private Sub vs3_AfterEdit(ByVal Row As Long, ByVal Col As Long)
''     If vs3.Col = 0 Then
''        'cellposiVs2
''        'vs3.TextMatrix(vs3.RowSel, 0) = cboitemvscboItemVs2.Text
''     End If
''
''End Sub
''
''Private Sub vs3_KeyDown(KeyCode As Integer, Shift As Integer)
''
''  If KeyCode = 46 Then
''    vs3.RemoveItem (vs3.RowSel)
''    'Total4
''  End If
''
''  If KeyCode = 13 Then
''     If vs3.Col = 0 Then
''
''        vs3.Editable = flexEDNone
''        FrameVs2.Visible = True
''        cboItemVs2.Visible = True
''        cboItemVs2.SetFocus
''     Else
''
''        vs3.Editable = flexEDKbdMouse
''
''     End If
''  End If
''
''End Sub
''
''Private Sub vs3_KeyUp(KeyCode As Integer, Shift As Integer)
''
''If KeyCode = 13 Then
''
'' If vs3.Col = 0 Then
''    vs3.Editable = flexEDNone
''    FrameVs2.Visible = True
''    cboItemVs2.SetFocus
''
''
''
''
''    Set rs = New ADODB.Recordset
''    If rs.State = 1 Then rs.Close
''    rs.Open "select * from ItemMaster where " & stringyear & " and  ItemName='" & vs3.TextMatrix(vs3.RowSel, 0) & "'", CON
''    If rs.EOF = False Then
''       vs3.TextMatrix(vs3.RowSel, 1) = rs.Fields("Unit").Value
''       SendKeys "{right}"
''       SendKeys "{right}"
''       FrameVs2.Visible = False
''       vs3.Editable = flexEDKbdMouse
''       vs3.SetFocus
''    Else
''       vs3.TextMatrix(vs3.RowSel, 1) = "Kg"
''       SendKeys "{right}"
''       SendKeys "{right}"
''       FrameVs2.Visible = False
''       vs3.Editable = flexEDKbdMouse
''       vs3.SetFocus
''
''    End If
''
'' End If
''
'' If vs3.Col = 2 Then
''
''   If rs.State = 1 Then rs.Close
''   rs.Open "select  OpeningStock from ItemMaster where " & stringyear & " and  ItemName='" & vs3.TextMatrix(vs.RowSel, 0) & "'", CON
''   If rs.EOF = False Then
''      If Val(rs.Fields(0).Value) < Val(vs3.TextMatrix(vs3.RowSel, 2)) Then
''         MsgBox "Stock Less !!", vbInformation
''
''      End If
''   End If
''
''
''    SendKeys "{home}"
''    SendKeys "{down}"
''
''    FrameVs2.TOP = FrameVs2.TOP + 170
''    'AddItemInGrid2
'' End If
''
''
''
'' 'Total4
''
''End If
''
''End Sub


