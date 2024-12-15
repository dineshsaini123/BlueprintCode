VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmExportInvoice 
   Caption         =   "Export Invoice"
   ClientHeight    =   10200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   15840
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtins 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      TabIndex        =   83
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox txtShipToadd 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10380
      TabIndex        =   87
      Top             =   3540
      Width           =   3660
   End
   Begin VB.TextBox txtFr 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      TabIndex        =   85
      Top             =   6960
      Width           =   1275
   End
   Begin VB.CheckBox Check1 
      Caption         =   "For Custom"
      Height          =   195
      Left            =   10500
      TabIndex        =   84
      Top             =   1500
      Width           =   2115
   End
   Begin VB.TextBox txtForCustom 
      Appearance      =   0  'Flat
      Height          =   825
      Left            =   10500
      MultiLine       =   -1  'True
      TabIndex        =   82
      Top             =   1740
      Width           =   3525
   End
   Begin VB.CheckBox Check_manullay 
      Caption         =   "Enter Packing No Manullay"
      Height          =   255
      Left            =   2520
      TabIndex        =   81
      Top             =   360
      Width           =   2235
   End
   Begin VB.TextBox txtOrderDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   11
      Top             =   1110
      Width           =   3585
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Height          =   795
      Left            =   240
      TabIndex        =   77
      Top             =   2820
      Width           =   3975
      Begin VB.TextBox txtShipNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1620
         TabIndex        =   9
         Top             =   480
         Width           =   2235
      End
      Begin VB.ComboBox cboref 
         Height          =   315
         ItemData        =   "frmExportInvoice.frx":0000
         Left            =   1620
         List            =   "frmExportInvoice.frx":0002
         TabIndex        =   8
         Top             =   120
         Width           =   2235
      End
      Begin VB.Label ExportCurrency 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order Ref."
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   79
         Top             =   180
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Shiping Bill No."
         Height          =   195
         Left            =   60
         TabIndex        =   78
         Top             =   480
         Width           =   1185
      End
   End
   Begin VB.TextBox txtwords 
      Height          =   315
      Left            =   300
      TabIndex        =   76
      Top             =   7320
      Width           =   9255
   End
   Begin VB.TextBox txtTotalNew 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11820
      TabIndex        =   70
      Top             =   6960
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "To be export in A/c"
      Height          =   1425
      Left            =   10380
      TabIndex        =   71
      Top             =   60
      Width           =   3615
      Begin VB.ListBox List_Balance 
         Height          =   1035
         Left            =   120
         TabIndex        =   72
         Top             =   300
         Width           =   3315
      End
   End
   Begin VB.TextBox txtCValue 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10860
      TabIndex        =   69
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox txtPackingNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1260
      TabIndex        =   67
      Top             =   360
      Width           =   1245
   End
   Begin VB.TextBox txtCountryOrigin 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   18
      Text            =   "India"
      Top             =   2910
      Width           =   1245
   End
   Begin VB.ComboBox txtExportCurrency 
      Height          =   315
      ItemData        =   "frmExportInvoice.frx":0004
      Left            =   1800
      List            =   "frmExportInvoice.frx":0014
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2160
      Width           =   1395
   End
   Begin VB.TextBox txtTotalPSC 
      Alignment       =   1  'Right Justify
      Height          =   240
      Left            =   12060
      TabIndex        =   26
      Top             =   7365
      Width           =   1155
   End
   Begin VB.TextBox txtTotalCartoons 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12060
      TabIndex        =   27
      Top             =   7620
      Width           =   1155
   End
   Begin VB.TextBox txtCBM 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12060
      TabIndex        =   28
      Top             =   7920
      Width           =   1155
   End
   Begin VB.TextBox txtNetW 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12060
      TabIndex        =   29
      Top             =   8220
      Width           =   1155
   End
   Begin VB.TextBox txtEPSG 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtPreCarriageBy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   20
      Top             =   3510
      Width           =   3585
   End
   Begin VB.TextBox txtPlaceofPrecarriage 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11760
      TabIndex        =   21
      Top             =   2580
      Width           =   2250
   End
   Begin VB.TextBox txtIECode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Text            =   "0505055317"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtExporterPan 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "AAEFC-7614G"
      Top             =   1860
      Width           =   2415
   End
   Begin VB.TextBox txtFDestination 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   19
      Top             =   3210
      Width           =   3585
   End
   Begin VB.TextBox txtino 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      TabIndex        =   0
      Top             =   60
      Width           =   1245
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   10
      Top             =   60
      Width           =   3690
   End
   Begin VB.TextBox txtAdd 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   735
      Left            =   6600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   375
      Width           =   3690
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6900
      TabIndex        =   25
      Top             =   6960
      Width           =   1395
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   300
      TabIndex        =   31
      Top             =   7740
      Width           =   8835
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
         Picture         =   "frmExportInvoice.frx":002F
         Style           =   1  'Graphical
         TabIndex        =   38
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
         Picture         =   "frmExportInvoice.frx":0C13
         Style           =   1  'Graphical
         TabIndex        =   37
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
         Picture         =   "frmExportInvoice.frx":17F7
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   135
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
         Picture         =   "frmExportInvoice.frx":23DB
         Style           =   1  'Graphical
         TabIndex        =   35
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
         Picture         =   "frmExportInvoice.frx":27E8
         Style           =   1  'Graphical
         TabIndex        =   34
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
         Picture         =   "frmExportInvoice.frx":33CC
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   120
         Width           =   1230
      End
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
         Picture         =   "frmExportInvoice.frx":3FB0
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   1230
      End
   End
   Begin VB.TextBox txtBuyerOrderNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   13
      Top             =   1410
      Width           =   3585
   End
   Begin VB.TextBox txtBuyerBatchNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   14
      Top             =   1710
      Width           =   3585
   End
   Begin VB.TextBox txtTermsOfDelivery 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   16
      Top             =   2310
      Width           =   3585
   End
   Begin VB.TextBox txtBrand 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   15
      Top             =   2010
      Width           =   3585
   End
   Begin VB.TextBox txtTermsOfPayment 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6600
      TabIndex        =   17
      Top             =   2610
      Width           =   3585
   End
   Begin VB.TextBox txtPortLoading 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11760
      TabIndex        =   22
      Top             =   2940
      Width           =   2280
   End
   Begin VB.TextBox txtPortDischarge 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11760
      TabIndex        =   23
      Top             =   3240
      Width           =   2280
   End
   Begin VB.TextBox txtGwt 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12060
      TabIndex        =   30
      Top             =   8520
      Width           =   1155
   End
   Begin Crystal.CrystalReport CR 
      Left            =   13575
      Top             =   6660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3090
      Left            =   180
      TabIndex        =   24
      Top             =   3795
      Width           =   13410
      _cx             =   23654
      _cy             =   5450
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   800
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmExportInvoice.frx":4B94
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
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   180
         Begin MSDataListLib.DataCombo cboItem 
            Height          =   2310
            Left            =   0
            TabIndex        =   40
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
      Left            =   3240
      TabIndex        =   1
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
      Left            =   3240
      TabIndex        =   2
      Top             =   780
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
      Left            =   3240
      TabIndex        =   3
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
   Begin VB.Label Label13 
      Caption         =   "Insurance :"
      Height          =   255
      Left            =   1200
      TabIndex        =   88
      Top             =   6960
      Width           =   795
   End
   Begin VB.Label Label12 
      Caption         =   "Frieght :"
      Height          =   255
      Left            =   3480
      TabIndex        =   86
      Top             =   6960
      Width           =   675
   End
   Begin VB.Label BuyersOrderNum 
      AutoSize        =   -1  'True
      Caption         =   "Order Date/Dates"
      Height          =   195
      Index           =   0
      Left            =   4620
      TabIndex        =   80
      Top             =   1110
      Width           =   1275
   End
   Begin VB.Label Label10 
      Caption         =   "INR  Value :"
      Height          =   255
      Left            =   5760
      TabIndex        =   75
      Top             =   6960
      Width           =   1035
   End
   Begin VB.Label ExportCurrency 
      AutoSize        =   -1  'True
      Caption         =   "Export Currency Values :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   8520
      TabIndex        =   74
      Top             =   7020
      Width           =   2115
   End
   Begin VB.Label Label9 
      Caption         =   "X"
      Height          =   195
      Left            =   11580
      TabIndex        =   73
      Top             =   7020
      Width           =   255
   End
   Begin VB.Shape Shape1 
      Height          =   435
      Left            =   10740
      Top             =   6900
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Packing No:"
      Height          =   270
      Index           =   3
      Left            =   180
      TabIndex        =   68
      Top             =   360
      Width           =   1110
   End
   Begin VB.Label Label8 
      Caption         =   "Country of Origin"
      Height          =   255
      Left            =   4620
      TabIndex        =   66
      Top             =   2910
      Width           =   1755
   End
   Begin VB.Label Label7 
      Caption         =   "Total Pcs"
      Height          =   255
      Left            =   10620
      TabIndex        =   65
      Top             =   7380
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Total Cartons"
      Height          =   195
      Left            =   10620
      TabIndex        =   64
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Total CBM"
      Height          =   195
      Left            =   10620
      TabIndex        =   63
      Top             =   7980
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Net Weight (Kg.)"
      Height          =   255
      Left            =   10620
      TabIndex        =   62
      Top             =   8220
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "EPCG License No."
      Height          =   195
      Left            =   180
      TabIndex        =   61
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label ExportCurrency 
      AutoSize        =   -1  'True
      Caption         =   "Export Currency"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   60
      Top             =   2220
      Width           =   1125
   End
   Begin VB.Label CountryofDestination 
      AutoSize        =   -1  'True
      Caption         =   "Country of Final Destination"
      Height          =   195
      Index           =   1
      Left            =   4620
      TabIndex        =   59
      Top             =   3210
      Width           =   1935
   End
   Begin VB.Label BuyerBatchNum 
      AutoSize        =   -1  'True
      Caption         =   "Buyer's Batch No."
      Height          =   195
      Index           =   1
      Left            =   4620
      TabIndex        =   58
      Top             =   1710
      Width           =   1875
   End
   Begin VB.Label BuyersOrderNum 
      AutoSize        =   -1  'True
      Caption         =   "Buyer's Order No(s)."
      Height          =   195
      Index           =   1
      Left            =   4620
      TabIndex        =   57
      Top             =   1410
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Dealer/Buyer :"
      Height          =   300
      Index           =   2
      Left            =   4620
      TabIndex        =   56
      Top             =   60
      Width           =   1785
   End
   Begin VB.Line Line5 
      X1              =   11100
      X2              =   19920
      Y1              =   -600
      Y2              =   -600
   End
   Begin VB.Label OrderDate 
      AutoSize        =   -1  'True
      Caption         =   "Order Date(s)"
      Height          =   195
      Left            =   14400
      TabIndex        =   55
      Top             =   -420
      Width           =   945
   End
   Begin VB.Label ExporterPan 
      AutoSize        =   -1  'True
      Caption         =   "Exporter's Pan"
      Height          =   195
      Left            =   180
      TabIndex        =   54
      Top             =   1860
      Width           =   1020
   End
   Begin VB.Label ExportCurrency 
      AutoSize        =   -1  'True
      Caption         =   "Brand"
      Height          =   195
      Index           =   0
      Left            =   4620
      TabIndex        =   53
      Top             =   2010
      Width           =   1920
   End
   Begin VB.Label PreCarriageBy 
      AutoSize        =   -1  'True
      Caption         =   "Pre Carriage By"
      Height          =   195
      Left            =   4620
      TabIndex        =   52
      Top             =   3510
      Width           =   1875
   End
   Begin VB.Label PlaceofCarriage 
      AutoSize        =   -1  'True
      Caption         =   "Place of Receipt by PreCarrige"
      Height          =   390
      Left            =   10320
      TabIndex        =   51
      Top             =   2580
      Width           =   1515
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Invoice No:"
      Height          =   270
      Index           =   0
      Left            =   210
      TabIndex        =   50
      Top             =   60
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "Date "
      Height          =   270
      Index           =   1
      Left            =   2760
      TabIndex        =   49
      Top             =   60
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Exporter's I.E. code"
      Height          =   300
      Index           =   4
      Left            =   180
      TabIndex        =   48
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Date &&  Time of issue of invoice"
      Height          =   270
      Index           =   7
      Left            =   180
      TabIndex        =   47
      Top             =   780
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Date &&  Time of Dispatch"
      Height          =   270
      Index           =   8
      Left            =   180
      TabIndex        =   46
      Top             =   1140
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Terms of Delivery"
      Height          =   210
      Index           =   11
      Left            =   4620
      TabIndex        =   45
      Top             =   2310
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "Port of Loading"
      Height          =   270
      Index           =   13
      Left            =   10320
      TabIndex        =   44
      Top             =   3000
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "Terms of Payment"
      Height          =   210
      Index           =   14
      Left            =   4620
      TabIndex        =   43
      Top             =   2610
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "Port of Discharge"
      Height          =   270
      Index           =   15
      Left            =   10320
      TabIndex        =   42
      Top             =   3240
      Width           =   1530
   End
   Begin VB.Label Label2 
      Caption         =   "Gross Weight (Kg.)"
      Height          =   255
      Left            =   10620
      TabIndex        =   41
      Top             =   8580
      Width           =   1335
   End
End
Attribute VB_Name = "frmExportInvoice"
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

txtTotal.Text = 0
txtTotalPSC.Text = 0


total1 = 0

txtTotalNew = 0

For J = 1 To vs.Rows - 1
If vs.TextMatrix(J, 1) <> "" Then
txtTotalNew.Text = (Val(txtTotalNew.Text) + Val(vs.TextMatrix(J, 9)))
End If

If Val(vs.TextMatrix(J, 7)) > 0 Then
txtTotalPSC.Text = (Val(txtTotalPSC.Text) + Val(vs.TextMatrix(J, 7)))
End If
'

Next

If Val(txtCvalue) > 0 Then
txtTotal = (Val(txtTotalNew.Text) * Val(txtCvalue))
Else
txtTotal = (Val(txtTotalNew.Text))
End If


txtTotalNew = Format(txtTotalNew, ".00")


txtTotal = Format(txtTotal, ".00")





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
'    'rs_3.Open "select * from ItemMaster where " & stringyear & " and  (ItemGp='Finish Item' or ItemGp='Scrap' or ItemGp='Losses' or (ItemGp='Chunnus (Pointing)' or ItemGp='Chunnus (Sonicut)') or  Itemgp='Rejuction') order by ItemName", con, adOpenDynamic, adLockOptimistic
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
'    'rs_2.Open "select * from ItemMaster where " & stringyear & " and  (ItemGp='Semi Finish (R/D)' or ItemGp= 'Semi Finish (Store)' or Itemgp='Rejuction') order by ItemName", con, adOpenDynamic, adLockOptimistic
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
'    'rs_1.Open "select * from ItemMaster where " & stringyear & " and  ItemGp='Raw Item' or ItemGp='Scrap' or (ItemGp='Chunnus (Pointing)' or ItemGp='Chunnus (Sonicut)') order by ItemName", con, adOpenDynamic, adLockOptimistic
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
        vs.TextMatrix(vs.RowSel, 0) = cboItem.Text
        vs.TextMatrix(vs.RowSel, 6) = cboItem.BoundText
        
        
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
'''    rs_u.Open "select * from Stock " & stringyear & " and  ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
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
'''        rs_u.Open "select * from Stock " & stringyear & " and  ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
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
'''        rs_u.Open "select * from Stock " & stringyear & " and  ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
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
'''        rs_u.Open "select * from Stock " & stringyear & " and  ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
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
AddSemifinish
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
Else
   txtForCustom.Visible = False
End If
End Sub

Private Sub cmdAdd_1_Click()
   
   
Dim o As Object

Check1.Value = 0

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
  
''  If Check_manullay.Value = 0 Then
''   txtino.Text = MaxSNo_Export("Casha", "INVOICENO")
''  End If
   
  txtino.Enabled = False
  edit = False
   
   addPendingBills
End Sub
Function MaxSNo_Export(tbl As String, fld As String) As Double
    Dim rs As New Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "Select max(" & fld & ") from " & tbl & " where " & stringyear & " and  typeofinvoice = 'export'", CON
    If IsNull(rs(0)) Then
        MaxSNo_Export = 1
    Else
        MaxSNo_Export = Val(rs(0)) + 1
    End If
    rs.Close
End Function


Private Sub cmdDelete_3_Click()
If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
   
   CON.BeginTrans
   CON.Execute "delete from Casha where " & stringyear & " and  invoiceNo =" & txtino & ""
   CON.Execute "delete from Cashb where " & stringyear & " and  invoiceNo =" & txtino & ""
   CON.Execute "delete from Cashc where " & stringyear & " and  invoiceNo =" & txtino & ""
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
''rs.Open "select * from Casha " & stringyear & " and  INVOICENO=" & txtHeating.Text & "", CON, adOpenDynamic, adLockOptimistic
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
''rs.Open "select * from Cashb " & stringyear & " and  INVOICENO=" & txtHeating.Text & "", CON, adOpenDynamic, adLockOptimistic
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
    
Dim net As Double
net = 0
    
net = (Val(txtTotalNew.Text) + Val(txtFr.Text) + Val(txtins.Text))
    
If (Val(txtFr) > 0 Or Val(txtins) > 0) Then

If Check1.Value = 0 Then

    
    
    CR.Reset
    CR.Connect = constr
    CR.ReportFileName = App.Path & "\REPORTS\ExportInvoice_fr.rpt"
    CR.ReplaceSelectionFormula "{ExportInvQry.invoiceno} = " & txtino & " "
    CR.Formulas(0) = "toword='" & txtwords.Text & "'"
    CR.Formulas(1) = "frieght='" & Format(txtFr.Text, ".00") & "'"
    CR.Formulas(2) = "netpay=" & Format(net, ".00") & ""
    
    CR.WindowState = crptMaximized
    CR.WindowShowPrintSetupBtn = True
    CR.Action = 1

Else

    CR.Reset
    CR.Connect = constr
    CR.ReportFileName = App.Path & "\REPORTS\ExportInvoice_forcustom_fr.rpt"
    
    CR.ReplaceSelectionFormula "{ExportInvQry.invoiceno} = " & txtino & " "
    CR.Formulas(0) = "toword='" & txtwords.Text & "'"
    CR.Formulas(1) = "frieght='" & Format(txtFr.Text, ".00") & "'"
    CR.Formulas(2) = "netpay=" & Format(net, ".00") & ""
    CR.WindowState = crptMaximized
    CR.WindowShowPrintSetupBtn = True
    CR.Action = 1
End If

Else

If Check1.Value = 0 Then

    CR.Reset
    CR.Connect = constr
    CR.ReportFileName = App.Path & "\REPORTS\ExportInvoice.rpt"
    CR.ReplaceSelectionFormula "{ExportInvQry.invoiceno} = " & txtino & " "
    CR.Formulas(0) = "toword='" & txtwords.Text & "'"
    CR.WindowState = crptMaximized
    CR.WindowShowPrintSetupBtn = True
    CR.Action = 1

Else

    CR.Reset
    CR.Connect = constr
    CR.ReportFileName = App.Path & "\REPORTS\ExportInvoice_forcustom.rpt"
    CR.ReplaceSelectionFormula "{ExportInvQry.invoiceno} = " & txtino & " "
    CR.Formulas(0) = "toword='" & txtwords.Text & "'"
    CR.WindowState = crptMaximized
    CR.WindowShowPrintSetupBtn = True
    CR.Action = 1
End If


End If



End Sub

Private Sub cmdSave_2_Click()

'
'On Error GoTo save:

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
Dim amt As Double
amt = 0

export_currency = txtExportCurrency.Text
export_flage = True

amt = (Val(txtTotalNew.Text) + Val(txtFr.Text))

txtwords = toword(amt)
export_flage = False

i = 1

If edit = False Then


            
''         If Check_manullay.Value = 0 Then
''            txtino.Text = MaxSNo_Export("Casha", "INVOICENO")
''          End If

            
            CON.BeginTrans
            
            If Val(txtCvalue) <> 0 Then

            CON.Execute "exec Export_Casha " & txtino & ",'" & dateInv & "','" & dateIssue & "'," & _
            "'" & dateDispatch & "','SUNDRY DEBTORS','" & txtparty.Text & "','" & txtIECode.Text & "','" & txtExporterPan.Text & "','" & txtExportCurrency.Text & "'," & _
            "'" & txtEPSG.Text & "','" & txtBuyerOrderNo.Text & "','" & txtBuyerBatchNo.Text & "','" & txtBrand.Text & "','" & txtTermsOfDelivery.Text & "','" & txtTermsOfPayment.Text & "'," & _
            "'" & txtFDestination.Text & "','" & txtPreCarriageBy.Text & "','" & txtPlaceofPrecarriage.Text & "','" & txtPortLoading.Text & "','" & txtPortDischarge.Text & "'," & _
            "'" & txtTotalPSC.Text & "','" & txtTotalCartoons.Text & "','" & txtCBM.Text & "','" & txtNetW.Text & "','" & txtGwt.Text & "'," & _
            "" & txtTotal.Text & "," & txtTotal.Text & ",0,'" & Val(txtCvalue.Text) & "'," & Val(txtTotalNew) & ",'" & main.username & "','" & main.username & "','" & main.session & "'," & _
            "" & main.setupid & ""
            
            'CON.Execute "update casha set orderby='" & cboref.Text & "',marka='" & Val(txtShipNo.Text) & "',orderdate='" & txtOrderDate.Text & "' " & stringyear & " and  INVOICENO=" & txtino.Text & ""
            
            CON.Execute "update casha set orderby='" & cboref.Text & "',marka='" & Val(txtShipNo.Text) & "',orderdate='" & txtOrderDate.Text & "'," & _
            "txt1='" & txtwords.Text & "',station='" & txtForCustom & "',freight=" & Val(txtFr.Text) & ",remark='" & txtShipToadd.Text & "',lexp3rate=" & Val(txtins.Text) & " where " & stringyear & " and  INVOICENO=" & txtino.Text & ""
   
            Else
            
            CON.Execute "exec Export_Casha " & txtino & ",'" & dateInv & "','" & dateIssue & "'," & _
            "'" & dateDispatch & "','SUNDRY DEBTORS','" & txtparty.Text & "','" & txtIECode.Text & "','" & txtExporterPan.Text & "','" & txtExportCurrency.Text & "'," & _
            "'" & txtEPSG.Text & "','" & txtBuyerOrderNo.Text & "','" & txtBuyerBatchNo.Text & "','" & txtBrand.Text & "','" & txtTermsOfDelivery.Text & "','" & txtTermsOfPayment.Text & "'," & _
            "'" & txtFDestination.Text & "','" & txtPreCarriageBy.Text & "','" & txtPlaceofPrecarriage.Text & "','" & txtPortLoading.Text & "','" & txtPortDischarge.Text & "'," & _
            "'" & txtTotalPSC.Text & "','" & txtTotalCartoons.Text & "','" & txtCBM.Text & "','" & txtNetW.Text & "','" & txtGwt.Text & "'," & _
            "" & "0.00" & "," & "0.00" & ",0,'" & Val(txtCvalue.Text) & "'," & Val(txtTotalNew) & ",'" & main.username & "','" & main.username & "','" & main.session & "'," & _
            "" & main.setupid & ""
            
            'CON.Execute "update casha set orderby='" & cboref.Text & "',marka='" & Val(txtShipNo.Text) & "',orderdate='" & txtOrderDate.Text & "' where " & stringyear & " and  INVOICENO=" & txtino.Text & ""
            CON.Execute "update casha set orderby='" & cboref.Text & "',marka='" & Val(txtShipNo.Text) & "',orderdate='" & txtOrderDate.Text & "'," & _
            "txt1='" & txtwords.Text & "',station='" & txtForCustom & "',freight=" & Val(txtFr.Text) & ",remark='" & txtShipToadd.Text & "',lexp3rate=" & Val(txtins.Text) & " where " & stringyear & " and  INVOICENO=" & txtino.Text & ""
   
            
            End If
            
            
            For i = 1 To vs.Rows - 1
            If (vs.TextMatrix(i, 1) <> "" And vs.TextMatrix(i, 2) <> "") Then

            CON.Execute "insert into Cashb(INVOICENO,SNO,Carton,itemname,BOOKCODE,btno,InnerPack,OuterPack,QUANTITY,RATE,AMOUNT,typeofinvoice)" & _
            " values(" & txtino.Text & ",'" & (vs.TextMatrix(i, 0)) & "','" & vs.TextMatrix(i, 1) & "','" & vs.TextMatrix(i, 2) & "','" & vs.TextMatrix(i, 3) & "'," & _
            "'" & vs.TextMatrix(i, 4) & "','" & vs.TextMatrix(i, 5) & "','" & vs.TextMatrix(i, 6) & "'," & _
            "'" & Val(vs.TextMatrix(i, 7)) & "','" & Val(vs.TextMatrix(i, 8)) & "','" & Val(vs.TextMatrix(i, 9)) & "','export')"
                      
            Else
            
                Exit For
            
            
            End If
            Next
            
                 
           '----------------------------
                 
            CON.Execute "insert into cashc" & _
            "(INVOICENO,INVOICEDate,GENLEDGER,Subledger,GAmount,Rate,Amount,typeofinvoice,DebitOrCredit," & _
            "text,fyear,createdby,createdon,updatedby,updatedon,setupid,currencyValue) values(" & _
            "" & txtino & ",'" & Format(dateInv, "MM/dd/yyyy") & "','SALES','" & exportsale & "'," & Val(txtTotal) & ",0," & _
            "" & Val(txtTotalNew.Text) & ",'tax','Debit','" & txtExportCurrency.Text & " " & Val(txtTotalNew) & " @" & Val(txtCvalue) & "/-" & "','" & main.session & "','" & main.username & "'," & _
            "'" & Format(Date, "MM/DD/yyyy") & "','" & main.username & "','" & Format(Date, "MM/dd/yyyy") & "'," & main.setupid & ",'" & Val(txtCvalue) & "')"
      
            '-----------------------------
            
            
            
            CON.CommitTrans
            
    Else
            
'         txtino = MaxSNo("Casha", "INVOICENO")
            CON.BeginTrans

            CON.Execute "delete from cashb where " & stringyear & " and  INVOICENO=" & txtino.Text & ""
            CON.Execute "delete from cashC where " & stringyear & " and  INVOICENO=" & txtino.Text & ""
            
            If Val(txtCvalue) <> 0 Then
            
            
             
             CON.Execute "update Casha set INVOICEDATE='" & Format(dateInv, "yyyy/mm/dd") & "',IssueDate='" & dateIssue & "',DisPatchDate=" & _
            "'" & dateDispatch & "',GENLEDGER='SUNDRY DEBTORS',SUBLEDGER='" & txtparty.Text & "',TXT1A='" & txtIECode.Text & "',TXT2='" & txtExporterPan.Text & "',TXT2A='" & txtExportCurrency.Text & "',T2=" & _
            "'" & txtEPSG.Text & "',ORDERNO='" & txtBuyerOrderNo.Text & "',AdviceRemark='" & txtBuyerBatchNo.Text & "',Brand='" & txtBrand.Text & "',TermsDelevery='" & txtTermsOfDelivery.Text & "',TermsPayment='" & txtTermsOfPayment.Text & "',COUNTRYDEST=" & _
            "'" & txtFDestination.Text & "',PRECARRIAGE='" & txtPreCarriageBy.Text & "',PLACERECEIPT='" & txtPlaceofPrecarriage.Text & "',PORTLOADING='" & txtPortLoading.Text & "',PORTDISCHARGE='" & txtPortDischarge.Text & "',TOTALPSC=" & _
            "'" & txtTotalPSC.Text & "',TotalCartons='" & txtTotalCartoons.Text & "',TotalCBM='" & txtCBM.Text & "',NetWeight='" & txtNetW.Text & "',GWeight='" & txtGwt.Text & "',NetAmount=" & _
            "" & txtTotal.Text & ",GAMOUNT=" & txtTotal.Text & ",NetRate=" & Val(txtTotalNew) & ",CurrencyValue='" & Val(txtCvalue.Text) & "',updatedby='" & main.username & "',updatedon=" & Date & ",typeofinvoice='export',orderby='" & cboref.Text & "',marka='" & Val(txtShipNo.Text) & "' where " & stringyear & " and  INVOICENO=" & txtino.Text & ""
         
            Else
            
            CON.Execute "update Casha set INVOICEDATE='" & Format(dateInv, "yyyy/mm/dd") & "',IssueDate='" & dateIssue & "',DisPatchDate=" & _
            "'" & dateDispatch & "',GENLEDGER='SUNDRY DEBTORS',SUBLEDGER='" & txtparty.Text & "',TXT1A='" & txtIECode.Text & "',TXT2='" & txtExporterPan.Text & "',TXT2A='" & txtExportCurrency.Text & "',T2=" & _
            "'" & txtEPSG.Text & "',ORDERNO='" & txtBuyerOrderNo.Text & "',AdviceRemark='" & txtBuyerBatchNo.Text & "',Brand='" & txtBrand.Text & "',TermsDelevery='" & txtTermsOfDelivery.Text & "',TermsPayment='" & txtTermsOfPayment.Text & "',COUNTRYDEST=" & _
            "'" & txtFDestination.Text & "',PRECARRIAGE='" & txtPreCarriageBy.Text & "',PLACERECEIPT='" & txtPlaceofPrecarriage.Text & "',PORTLOADING='" & txtPortLoading.Text & "',PORTDISCHARGE='" & txtPortDischarge.Text & "',TOTALPSC=" & _
            "'" & txtTotalPSC.Text & "',TotalCartons='" & txtTotalCartoons.Text & "',TotalCBM='" & txtCBM.Text & "',NetWeight='" & txtNetW.Text & "',GWeight='" & txtGwt.Text & "',NetAmount=" & _
            "" & "0.00" & ",GAMOUNT=" & "0.00" & ",NetRate=" & Val(txtTotalNew) & ",CurrencyValue='" & Val(txtCvalue.Text) & "',updatedby='" & main.username & "',updatedon=" & Date & ",typeofinvoice='export',orderby='" & cboref.Text & "',marka='" & Val(txtShipNo.Text) & "' where " & stringyear & " and  INVOICENO=" & txtino.Text & ""
            
            End If
        
            
            For i = 1 To vs.Rows - 1
            If (vs.TextMatrix(i, 1) <> "" And vs.TextMatrix(i, 2) <> "") Then

            CON.Execute "insert into Cashb(INVOICENO,SNO,Carton,itemname,BOOKCODE,btno,InnerPack,OuterPack,QUANTITY,RATE,AMOUNT,typeofinvoice)" & _
            " values(" & txtino.Text & ",'" & (vs.TextMatrix(i, 0)) & "','" & vs.TextMatrix(i, 1) & "','" & vs.TextMatrix(i, 2) & "','" & vs.TextMatrix(i, 3) & "'," & _
            "'" & vs.TextMatrix(i, 4) & "','" & vs.TextMatrix(i, 5) & "','" & vs.TextMatrix(i, 6) & "'," & _
            "'" & Val(vs.TextMatrix(i, 7)) & "','" & Val(vs.TextMatrix(i, 8)) & "','" & Val(vs.TextMatrix(i, 9)) & "','export')"
                      
            Else
            
                Exit For
            
            
            End If
            Next
            
            
            
            
           '----------------------------
                 
            CON.Execute "insert into cashc" & _
            "(INVOICENO,INVOICEDate,GENLEDGER,Subledger,GAmount,Rate,Amount,typeofinvoice,DebitOrCredit," & _
            "text,fyear,createdby,createdon,updatedby,updatedon,setupid,currencyValue) values(" & _
            "" & txtino & ",'" & Format(dateInv, "MM/dd/yyyy") & "','SALES','" & exportsale & "'," & Val(txtTotal) & ",0," & _
            "" & Val(txtTotalNew.Text) & ",'tax','Debit','" & txtExportCurrency.Text & " " & Val(txtTotalNew) & " @" & Val(txtCvalue) & "/-" & "','" & main.session & "','" & main.username & "'," & _
            "'" & Format(Date, "MM/DD/yyyy") & "','" & main.username & "','" & Format(Date, "MM/dd/yyyy") & "'," & main.setupid & ",'" & (txtCvalue) & "')"
      
            '-----------------------------
   
            CON.Execute "update casha set orderby='" & cboref.Text & "',marka='" & Val(txtShipNo.Text) & "',orderdate='" & txtOrderDate.Text & "'," & _
            "txt1='" & txtwords.Text & "',station='" & txtForCustom & "',FREIGHT=" & Val(txtFr.Text) & ",remark='" & txtShipToadd.Text & "',lexp3rate=" & Val(txtins.Text) & " where " & stringyear & " and  INVOICENO=" & txtino.Text & ""
            
            
            CON.CommitTrans
                    
            
            
            

End If
MsgBox "Data Saved..", vbInformation

cmdprint.Enabled = True

Check_manullay.Value = 0



cmdEdit_4.Enabled = True
cmdDelete_3.Enabled = True
'Call cmdAdd_1_Click

'Exit Sub
'
'
'save:
'
'CON.RollbackTrans
'MsgBox err.DESCRIPTION, vbCritical




End Sub

Private Sub cmdSearch_Click()
popuplist10 "select [INVOICENO], [INVOICEDATE], [SUBLEDGER] from [Casha] where " & stringyear & " order by cast(INVOICENO as int)", CON
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
Sub addPendingBills()

List_Balance.Clear

If rs.State = 1 Then rs.Close
rs.Open "select invoiceno from casha where " & stringyear & " and  (typeofinvoice<>'Packing' and currencyValue='0') order by invoiceno", CON
While rs.EOF = False
List_Balance.AddItem rs(0)
rs.MoveNext
Wend


End Sub
Sub searchData()

'On Error Resume Next

vs.Clear
setwidth


If rs.State = 1 Then rs.Close
If txtino = "" Then Exit Sub
st1 = "select INVOICENO,INVOICEDATE,IssueDate,DisPatchDate," & _
"TXT1A,TXT2,TXT2A,T2,SUBLEDGER,ORDERNO,ADVICEREMARK,Brand,TermsDelevery,TermsPayment,COUNTRYDEST,PRECARRIAGE," & _
"PLACERECEIPT,PORTLOADING,PORTDISCHARGE,TOTALPSC,TOTALCARTONS,TOTALCBM,NETWEIGHT,GWEIGHT," & _
"NetAmount,CurrencyValue,orderby,marka,txt1,Station,freight,lexp3rate from Casha where " & stringyear & " and  invoiceno=" & txtino & ""
     
rs.Open st1, CON

'If rs.EOF = True Then
'txtwords = toword(txtTotalNew.Text)
'End If

If rs.EOF = False Then
  
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

  
  
  txtwords.Text = rs!txt1 & ""
  cboref.Text = rs!ORDERBY & ""
  txtShipNo.Text = rs!marka & ""
  
  txtCvalue = rs!CurrencyValue & ""
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
  txtTotal = rs!netamount
  
  txtins = rs!lexp3rate & ""
  
  'txtCBM = rs!TOTALCBM
  'txtNetW = rs!NETWEIGHT
  'txtGwt = rs!GWEIGHT
  
  
  
  txtCBM = Format(rs!TOTALCBM, "0.00")
  txtNetW = Format(rs!NETWEIGHT, "0.000")
  txtGwt = Format(rs!GWEIGHT, "0.000")
  txtFr.Text = Format(rs!freight, "0.00")
  
  
    
If rs.State = 1 Then rs.Close
rs.Open "select ADDRESS1,ADDRESS2 from SLEDGER where " & stringyear & " and  SUBLEDGER = '" & txtparty.Text & "'", CON
Dim addr As String
If rs.EOF = False Then
addr = rs!address1 & "" & vbCrLf & "" & rs!ADDRESS2
txtAdd.Text = addr
End If



    



If rs.State = 1 Then rs.Close
rs.Open "select INVOICENO,SNO,Carton,itemname,BOOKCODE,btno,InnerPack,OuterPack,QUANTITY,RATE,AMOUNT" & _
" from Cashb where " & stringyear & " and  INVOICENO=" & txtino.Text & " and typeofinvoice='export' order by sno", CON
For i = 1 To rs.RecordCount
vs.TextMatrix(i, 0) = rs!sno
vs.TextMatrix(i, 1) = rs!Carton
vs.TextMatrix(i, 2) = rs!itemname
vs.TextMatrix(i, 3) = rs!btno
vs.TextMatrix(i, 4) = rs!Bookcode
vs.TextMatrix(i, 5) = rs!InnerPack
vs.TextMatrix(i, 6) = rs!OuterPack
vs.TextMatrix(i, 7) = rs!quantity
vs.TextMatrix(i, 8) = rs!rate
vs.TextMatrix(i, 9) = Format(rs!amount, ".00")
rs.MoveNext
Next



End If

cmdSave_2.Enabled = False
cmdDelete_3.Enabled = False
cmdAdd_1.Enabled = True
cmdprint.Enabled = True
cmdEdit_4.Enabled = True
txtino.Enabled = False

Total


export_currency = txtExportCurrency.Text
export_flage = True
txtwords = toword(txtTotalNew.Text)
export_flage = False


End Sub
Sub SearchPacking()

On Error Resume Next

vs.Clear
setwidth


 

If rs.State = 1 Then rs.Close

st1 = "select INVOICEDATE,IssueDate,DisPatchDate," & _
"TXT1A,TXT2,TXT2A,T2,SUBLEDGER,ORDERNO,ADVICEREMARK,Brand,TermsDelevery,TermsPayment,COUNTRYDEST,PRECARRIAGE," & _
"PLACERECEIPT,PORTLOADING,PORTDISCHARGE,TOTALPSC,TOTALCARTONS,TOTALCBM,NETWEIGHT,GWEIGHT," & _
"NetAmount,station,remark from Casha where " & stringyear & " and  PackingNo=" & txtPackingNo & " and typeofinvoice='Packing'"
     
rs.Open st1, CON
If rs.EOF = False Then



    Check1.Value = 0
    txtForCustom = ""
    txtForCustom.Visible = False


  
  If Not IsNull(rs!station) Then
    If Len(rs!station) > 0 Then
       txtForCustom = rs!station
       Check1.Value = 1
    End If
  End If
       
       

  txtShipToadd.Text = rs!remark & ""
    
  dateInv = rs!InvoiceDate
  dateIssue = Format(rs!IssueDate, "dd/MM/yyyy")
  If IsDate(rs!DisPatchDate) Then
     dateDispatch = Format(rs!DisPatchDate, "dd/MM/yyyy")
  End If
  
  txtIECode = rs!txt1a & ""
  txtExporterPan = rs!txt2 & ""
  txtExportCurrency.Text = rs!txt2a
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
 ' txtTotal = rs!netamount
    
If rs.State = 1 Then rs.Close
rs.Open "select ADDRESS1,ADDRESS2,Country from SLEDGER where " & stringyear & " and  SUBLEDGER = '" & txtparty.Text & "'", CON
Dim addr As String
addr = rs!address1 & "" & vbCrLf & "" & rs!ADDRESS2 & vbCrLf & rs!country
txtAdd.Text = addr



    



If rs.State = 1 Then rs.Close
rs.Open "select SNO,Carton,itemname,BOOKCODE,btno,InnerPack,OuterPack,QUANTITY,RATE,AMOUNT" & _
" from Cashb where " & stringyear & " and  PackingNo=" & txtPackingNo.Text & " and typeofinvoice='Packing' order by sno", CON
For i = 1 To rs.RecordCount
vs.TextMatrix(i, 0) = rs!sno
vs.TextMatrix(i, 1) = rs!Carton
vs.TextMatrix(i, 2) = rs!itemname
vs.TextMatrix(i, 3) = rs!btno
vs.TextMatrix(i, 4) = rs!Bookcode
vs.TextMatrix(i, 5) = rs!InnerPack
vs.TextMatrix(i, 6) = rs!OuterPack
vs.TextMatrix(i, 7) = rs!quantity
'vs.TextMatrix(i, 8) = rs!rate
'vs.TextMatrix(i, 9) = rs!amount
rs.MoveNext
Next

End If

''cmdSave_2.Enabled = False
''cmdDelete_3.Enabled = False
''cmdAdd_1.Enabled = True
''cmdPrint.Enabled = True
''cmdEdit_4.Enabled = True
'''txtino.Enabled = False




End Sub



Sub TotalFinal()
   If txtTotal3.Text = "" Then
      txtTotal3.Text = 0
   End If
   
   If txtTotal2.Text = "" Then
      txtTotal2.Text = 0
   End If
   
   
    txtRawAndCasting.Text = (CDbl(txtTotal2.Text) + CDbl(txtTotal3.Text))
    txtRawAndCasting.Text = Format(txtRawAndCasting.Text, "#,###.000")
End Sub
Private Sub Form_Load()
 
 setwidth
 
 'dateInv.Text = Format(Date, "dd/MM/yyyy")
 
 txtino = MaxSNo("Casha", "INVOICENO")
 
 'txtRate(0) = VAT
 
 
 
 withForm = VAT1
 withoutForm = VAT
 
 cmdDelete_3.Enabled = False
cmdEdit_4.Enabled = False
cmdprint.Enabled = False


If rs.State = 1 Then rs.Close
rs.Open "select agentname from agentmaster group by agentname", CON
While rs.EOF = False
  cboref.AddItem rs(0)
  rs.MoveNext
Wend


ButtonPermission cmdSave_2, cmdDelete_3, cmdEdit_4

addPendingBills
End Sub
Sub setwidth()
vs.Cols = 10
vs.FormatString = "<S. No.|Carton|<Description of Goods|^Product Code|^Batch No.|^Inner Pack|^Outer Pack|>Quantity|>Rate|>Net Amount"
vs.ColWidth(0) = 300
vs.ColWidth(1) = 1200
vs.ColWidth(2) = 4000
vs.ColWidth(3) = 1200
vs.ColWidth(4) = 1000
vs.ColWidth(5) = 1000
vs.ColWidth(6) = 1000
vs.ColWidth(7) = 900
vs.ColWidth(8) = 800
vs.ColWidth(9) = 1200

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
  TotalFinal
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





Private Sub List_Balance_Click()
   txtino = List_Balance.Text
   searchData
End Sub



Private Sub txtBrand_LostFocus()
txtBrand = UCase(txtBrand)
End Sub

Private Sub txtBuyerBatchNo_LostFocus()
txtBuyerBatchNo = UCase(txtBuyerBatchNo)
End Sub

Private Sub txtBuyerOrderNo_LostFocus()
   txtBuyerOrderNo = UCase(txtBuyerOrderNo)
End Sub

Private Sub txtCValue_GotFocus()
HIT
End Sub

Private Sub txtCValue_LostFocus()
Total
End Sub

Private Sub txtExportCurrency_LostFocus()
If txtExportCurrency.Text = "" Then
MsgBox "Please select export currency...", vbInformation
txtExportCurrency.SetFocus

End If

If txtExportCurrency.Text <> "" Then
If rs.State = 1 Then rs.Close
rs.Open "Select * from CurValues where " & stringyear & " and  CName='" & txtExportCurrency.Text & "'", CON
If Not rs.EOF Then
txtCvalue.Text = rs!CValue
End If
End If
End Sub

Private Sub txtExportCurrency_Click()
''If txtExportCurrency.Text <> "" Then
''If rs.State = 1 Then rs.Close
''rs.Open "Select * from CurValues where " & stringyear & " and  CName='" & txtExportCurrency.Text & "'", CON
''If Not rs.EOF Then
''txtCvalue.Text = rs!CValue
''End If
''End If
End Sub


Private Sub txtFDestination_LostFocus()
txtFDestination = UCase(txtFDestination)
End Sub



Private Sub txtForCustom_LostFocus()
txtForCustom = UCase(txtForCustom)
End Sub

Private Sub txtFr_GotFocus()
HIT
End Sub

Private Sub txtino_GotFocus()
HIT
End Sub

Private Sub txtino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  searchData
  
End If
End Sub





Private Sub txtPackingNo_KeyPress(KeyAscii As Integer)
On Error GoTo aa1:
If KeyAscii = 13 Then
  If rs.State = 1 Then rs.Close
  rs.Open "SELECT invoiceno FROM Casha where " & stringyear & " and  INVOICENO=" & txtPackingNo & "", CON
  If rs.EOF = False Then
     MsgBox "Invoice Already Creadted, Related Packing No ", vbCritical
     Exit Sub
  End If
  txtino = txtPackingNo
  SearchPacking
End If

Exit Sub
aa1:

MsgBox err.DESCRIPTION

End Sub

Private Sub txtPackingNo_LostFocus()
txtino = txtPackingNo
End Sub

Private Sub txtparty_GotFocus()
HIT
If PopUpValue1 <> "" Then
   txtparty = PopUpValue1
   txtAdd.Text = PopUpValue2 & vbCrLf & PopUpValue3
   
   
'  If RS.State = 1 Then RS.Close
'   RS.Open "SELECT [SUBLEDGER],address1,address2,[City],[District],[State],[bcityid],typeofcust " & _
'   "from [ExportData].[dbo].[SubledgerQry] where " & stringyear & " and  SUBLEDGER ='" & txtParty & "'", CON
'   If RS.EOF = False Then
'     txtAdd = RS![SUBLEDGER] & " " & vbCrLf & RS!address1 & " " & RS!address1 & vbCrLf & RS![CITY] + vbCrLf + RS![District] & "," & RS![State]
'   End If
   
   
   
   FatchTaxFromSate
   
   txtBuyerOrderNo.SetFocus
   
   PopUpValue1 = ""
   PopUpValue2 = ""
   PopUpValue3 = ""
   popupvalue4 = ""
   
End If

End Sub
Private Sub txtParty_KeyUp(KeyCode As Integer, Shift As Integer)
'If (KeyCode = 13 Or KeyCode = 144) Then Exit Sub
'tblNo = 1
'frmSearchItem.Show

popuplist10 "select Subledger,Address1,Address2 from Sledger where " & stringyear & "  order by Subledger", CON


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

Private Sub txtPortLoading_LostFocus()
txtPortLoading = UCase(txtPortLoading)
End Sub

Private Sub txtPreCarriageBy_LostFocus()
txtPreCarriageBy = UCase(txtPreCarriageBy)
End Sub

Private Sub txtTermsOfDelivery_LostFocus()
txtTermsOfDelivery = UCase(txtTermsOfDelivery)
End Sub

Private Sub txtTermsOfPayment_LostFocus()
txtTermsOfPayment = UCase(txtTermsOfPayment)
End Sub


Private Sub txtwords_LostFocus()
txtwords = UCase(txtwords)
End Sub

Private Sub vs_AfterEdit(ByVal Row As Long, ByVal Col As Long)
''     If vs.col = 0 Then
''        cellposi
''        vs.TextMatrix(vs.RowSel, 0) = cboItem.Text
''     End If
End Sub

Private Sub VS_KeyDown(KeyCode As Integer, Shift As Integer)

  
  
  If KeyCode = 115 Then
  If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
    vs.RemoveItem (vs.RowSel)
    Total
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

Dim Item As String
Dim totalcarton


If KeyCode = 13 Then
        
 If vs.Col = 1 Then
    
    
    
 
    If vs.TextMatrix(vs.RowSel, 1) = "" Then
       K = InStr(1, vs.TextMatrix(vs.RowSel - 1, 1), "-")
       If K <> 0 Then
        totalcarton = Split(vs.TextMatrix((vs.RowSel - 1), 1), "-")
        txtTotalCartoons.Text = totalcarton(1)
        txtTotalCartoons.SetFocus
       End If
       
       
    End If


If vs.TextMatrix(vs.RowSel, 1) <> "" Then
 SendKeys "{right}"
End If
 
 
 
 ElseIf vs.Col = 2 Then
    vs.TextMatrix(vs.RowSel, vs.Col) = UCase(vs.TextMatrix(vs.RowSel, vs.Col))
    SendKeys "{right}"
    
 ElseIf vs.Col = 3 Then
    vs.TextMatrix(vs.RowSel, vs.Col) = UCase(vs.TextMatrix(vs.RowSel, vs.Col))
    
    SendKeys "{right}"
 ElseIf vs.Col = 4 Then
    SendKeys "{right}"
 ElseIf vs.Col = 5 Then
    SendKeys "{right}"
 ElseIf vs.Col = 6 Then
    SendKeys "{right}"
 ElseIf vs.Col = 7 Then
    SendKeys "{right}"
    
 ElseIf vs.Col = 8 Then
    vs.TextMatrix(vs.RowSel, 9) = Format(Val(vs.TextMatrix(vs.RowSel, 7)) * Val(vs.TextMatrix(vs.RowSel, 8)), ".00")
    vs.TextMatrix(vs.RowSel, 0) = vs.Row
    
    SendKeys "{down}"
    'SendKeys "{home}"
    
 Else
    ''
    ''
 End If
 
 

Total




End If

End Sub
Sub FatchTaxFromSate()

'Dim ST As String
'Dim with_without As String
'
'
'
'
'
'If RS.State = 1 Then RS.Close
'RS.Open "select State,tinno from SubledgerQry " & stringyear & " and  subledger='" & txtParty & "'", CON
'If RS.EOF = False Then
'   ST = RS(0)
'If LCase(ST) = "u.p." Then
' If Len(RS!tinno) > 0 Then
'   'header = "TAX-INVOICE"
' Else
'   'header = "SALE-INVOICE"
' End If
'
'Else
'   'header = "SALE-INVOICE"
'
'End If
'
'End If
'
'
'If RS.State = 1 Then RS.Close
'RS.Open "select add_val,less_val from [state_tax_list] " & stringyear & " and  statename='" & ST & "' and with_without='" & with_without & "'", CON
'If RS.EOF = False Then
'   VAT_Add = RS(0)
'   VAT_less = RS(1)
'
'End If
'
'


End Sub

Private Sub vs_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)

'If KeyCode = 13 Then
'
'If vs.Col = 1 Then
'If Not (KeyCode >= 48 And KeyCode <= 57) Then
'MsgBox "Enter Only Numeric Value !!", vbInformation
'vs.TextMatrix(vs.RowSel, 1) = ""
'vs.SetFocus
'Exit Sub
'End If
'End If
'
'
'If vs.Col = 2 Then
'If Not (KeyCode >= 48 And KeyCode <= 57) Then
'MsgBox "Enter Only Numeric Value !!", vbInformation
'vs.TextMatrix(vs.RowSel, 2) = ""
'vs.SetFocus
'Exit Sub
'End If
'End If
'
'
'If vs.Col = 3 Then
'If Not (KeyCode >= 48 And KeyCode <= 57) Then
'MsgBox "Enter Only Numeric Value !!", vbInformation
'vs.TextMatrix(vs.RowSel, 3) = ""
'vs.SetFocus
'Exit Sub
'End If
'End If
'
'
'If vs.Col = 4 Then
'If Not (KeyCode >= 48 And KeyCode <= 57) Then
'MsgBox "Enter Only Numeric Value !!", vbInformation
'vs.TextMatrix(vs.RowSel, 4) = ""
'vs.SetFocus
'Exit Sub
'End If
'End If
'
'
'
'
'End If

End Sub

Private Sub vs_LeaveCell()
  Total
End Sub

Private Sub vs1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs1.Col = 0 Then
        cellposiVs
        vs1.TextMatrix(vs1.RowSel, 0) = cboitemvs1.Text
     End If

End Sub
Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 46 Then
    vs1.RemoveItem (vs1.RowSel)
    'Total1
    TotalFinal
  End If
  
  If KeyCode = 13 Then
     If vs1.Col = 0 Then
        vs1.Editable = flexEDNone
        Vs1Frame.Visible = True
        cboitemvs1.Visible = True
        cboitemvs1.SetFocus
     Else
        vs1.Editable = flexEDKbdMouse
        cellposiVs
     End If
  End If
End Sub

Private Sub vs1_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
        
 If vs1.Col = 0 Then
    vs1.Editable = flexEDNone
    Vs1Frame.Visible = True
    cboitemvs1.SetFocus
          
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "select * from ItemMaster where " & stringyear & " and  ItemName='" & vs1.TextMatrix(vs1.RowSel, 0) & "'", CON
    If rs.EOF = False Then
       vs1.TextMatrix(vs1.RowSel, 1) = rs.Fields("Unit").Value
       SendKeys "{right}"
       SendKeys "{right}"
       Vs1Frame.Visible = False
       vs1.Editable = flexEDKbdMouse
       vs1.SetFocus
    Else
       vs1.TextMatrix(vs1.RowSel, 1) = "Kg"
       SendKeys "{right}"
       SendKeys "{right}"
       Vs1Frame.Visible = False
       vs1.Editable = flexEDKbdMouse
       vs1.SetFocus
    
    End If
    
 End If
    
 If vs1.Col = 2 Then
           
    SendKeys "{home}"
    SendKeys "{down}"
    
    AddItemInGrid1
 End If
    
    

 'Total1
 TotalFinal

End If


End Sub
Sub AddSemifinish()
   Dim J As Integer
   
   J = 1
    
   vs3.Clear
   For i = 1 To vs1.Rows - 1
    
   If vs1.TextMatrix(i, 0) <> "" Then
      If rs.State = 1 Then rs.Close
      rs.Open "select * from ItemMaster where " & stringyear & " and  ItemName='" & vs1.TextMatrix(i, 0) & "'", CON
      If rs.Fields("itemgp").Value = "Semi Finish (R/D)" Or rs.Fields("itemgp").Value = "Semi Finish (Store)" Then
         vs3.TextMatrix(J, 0) = vs1.TextMatrix(i, 0)
         vs3.TextMatrix(J, 1) = vs1.TextMatrix(i, 1)
         vs3.TextMatrix(J, 2) = vs1.TextMatrix(i, 2)
         J = J + 1
      End If
   End If
        
   Next
    
    
End Sub
Private Sub vs1_LeaveCell()
   'Total1
End Sub

Private Sub vs2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs2.Col = 0 Then
        'cellposiVs3
        vs2.TextMatrix(vs2.RowSel, 0) = cboItemVs3.Text
     End If

End Sub

Private Sub vs2_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 46 Then
    vs2.RemoveItem (vs2.RowSel)
    'Total2
    TotalFinal
 End If

  
  If KeyCode = 13 Then
     
     If vs2.Col = 0 Then
        vs2.Editable = flexEDNone
        Vs3Frame.Visible = True
        cboItemVs3.Visible = True
        cboItemVs3.SetFocus
     Else
        vs2.Editable = flexEDKbdMouse
        'cellposiVs3
     End If

  End If

End Sub

Private Sub vs2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
        
          
 If vs2.Col = 0 Then
 
      vs2.Editable = flexEDNone
      Vs3Frame.Visible = True
      
      
          
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "select * from ItemMaster where " & stringyear & " and  ItemName='" & vs2.TextMatrix(vs2.RowSel, 0) & "'", CON
    If rs.EOF = False Then
       vs2.TextMatrix(vs2.RowSel, 1) = rs.Fields("Unit").Value
       SendKeys "{right}"
       SendKeys "{right}"
       Vs3Frame.Visible = False
       vs2.Editable = flexEDKbdMouse
       vs2.SetFocus
    Else
       vs2.TextMatrix(vs2.RowSel, 1) = "Kg"
       SendKeys "{right}"
       SendKeys "{right}"
       Vs3Frame.Visible = False
       vs2.Editable = flexEDKbdMouse
       vs2.SetFocus
    
    End If
    
 End If
 
    
    If vs2.Col = 2 Then
           
           SendKeys "{home}"
           SendKeys "{down}"
           Vs3Frame.TOP = Vs3Frame.TOP + 170
    End If
    
       
   'Total2

End If

End Sub
Private Sub vs2_LeaveCell()
   'Total2
End Sub
Private Sub vs3_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     If vs3.Col = 0 Then
        'cellposiVs2
        'vs3.TextMatrix(vs3.RowSel, 0) = cboitemvscboItemVs2.Text
     End If
 
End Sub

Private Sub vs3_KeyDown(KeyCode As Integer, Shift As Integer)
    
  If KeyCode = 46 Then
    vs3.RemoveItem (vs3.RowSel)
    'Total4
  End If
  
  If KeyCode = 13 Then
     If vs3.Col = 0 Then
        
        vs3.Editable = flexEDNone
        FrameVs2.Visible = True
        cboItemVs2.Visible = True
        cboItemVs2.SetFocus
     Else
        
        vs3.Editable = flexEDKbdMouse
        
     End If
  End If

End Sub

Private Sub vs3_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
 
 If vs3.Col = 0 Then
    vs3.Editable = flexEDNone
    FrameVs2.Visible = True
    cboItemVs2.SetFocus
    
    
 
          
    Set rs = New ADODB.Recordset
    If rs.State = 1 Then rs.Close
    rs.Open "select * from ItemMaster where " & stringyear & " and  ItemName='" & vs3.TextMatrix(vs3.RowSel, 0) & "'", CON
    If rs.EOF = False Then
       vs3.TextMatrix(vs3.RowSel, 1) = rs.Fields("Unit").Value
       SendKeys "{right}"
       SendKeys "{right}"
       FrameVs2.Visible = False
       vs3.Editable = flexEDKbdMouse
       vs3.SetFocus
    Else
       vs3.TextMatrix(vs3.RowSel, 1) = "Kg"
       SendKeys "{right}"
       SendKeys "{right}"
       FrameVs2.Visible = False
       vs3.Editable = flexEDKbdMouse
       vs3.SetFocus
    
    End If
    
 End If
    
 If vs3.Col = 2 Then
    
   If rs.State = 1 Then rs.Close
   rs.Open "select  OpeningStock from ItemMaster where " & stringyear & " and  ItemName='" & vs3.TextMatrix(vs.RowSel, 0) & "'", CON
   If rs.EOF = False Then
      If Val(rs.Fields(0).Value) < Val(vs3.TextMatrix(vs3.RowSel, 2)) Then
         MsgBox "Stock Less !!", vbInformation
         
      End If
   End If
    
    
    SendKeys "{home}"
    SendKeys "{down}"
    
    FrameVs2.TOP = FrameVs2.TOP + 170
    'AddItemInGrid2
 End If
    
    

 'Total4
 
End If
 
End Sub




