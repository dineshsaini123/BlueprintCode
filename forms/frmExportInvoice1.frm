VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmExportInvoice 
   Caption         =   "Form1"
   ClientHeight    =   10200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15030
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   15030
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtTotalPSC 
      Height          =   285
      Left            =   11700
      TabIndex        =   60
      Top             =   8040
      Width           =   1035
   End
   Begin VB.TextBox txtTotalCartoons 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11700
      TabIndex        =   56
      Top             =   8340
      Width           =   1035
   End
   Begin VB.TextBox txtCBM 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11700
      TabIndex        =   55
      Top             =   8640
      Width           =   1035
   End
   Begin VB.TextBox txtNetW 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11700
      TabIndex        =   54
      Top             =   9000
      Width           =   1035
   End
   Begin VB.TextBox txtEPSG 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   52
      Top             =   3120
      Width           =   2415
   End
   Begin VB.TextBox txtPreCarriageBy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6660
      TabIndex        =   50
      Top             =   3060
      Width           =   3585
   End
   Begin VB.TextBox txtPlaceofPrecarriage 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11760
      TabIndex        =   49
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox txtIECode 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   1980
      Width           =   2415
   End
   Begin VB.TextBox txtExporterPan 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   38
      Top             =   2340
      Width           =   2415
   End
   Begin VB.TextBox txtExportCurrency 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   37
      Top             =   2700
      Width           =   2415
   End
   Begin VB.TextBox txtFDestination 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6660
      TabIndex        =   36
      Top             =   2760
      Width           =   3585
   End
   Begin VB.TextBox txtino 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      TabIndex        =   19
      Top             =   300
      Width           =   1245
   End
   Begin VB.TextBox txtParty 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6660
      TabIndex        =   18
      Top             =   300
      Width           =   3690
   End
   Begin VB.TextBox txtAdd 
      Appearance      =   0  'Flat
      Height          =   585
      Left            =   6660
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   600
      Width           =   3690
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11700
      TabIndex        =   16
      Top             =   7740
      Width           =   1035
   End
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   300
      TabIndex        =   8
      Top             =   8400
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
         Picture         =   "frmExportInvoice.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Left            =   1290
         Picture         =   "frmExportInvoice.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
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
         Picture         =   "frmExportInvoice.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "frmExportInvoice.frx":23AC
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "frmExportInvoice.frx":27B9
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "frmExportInvoice.frx":339D
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
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
         Picture         =   "frmExportInvoice.frx":3F81
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   1230
      End
   End
   Begin VB.TextBox txtBuyerOrderNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6660
      TabIndex        =   7
      Top             =   1260
      Width           =   3585
   End
   Begin VB.TextBox txtBuyerBatchNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6660
      TabIndex        =   6
      Top             =   1560
      Width           =   3585
   End
   Begin VB.TextBox txtTermsOfDelivery 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6660
      TabIndex        =   5
      Top             =   2160
      Width           =   3585
   End
   Begin VB.TextBox txtBrand 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6660
      TabIndex        =   4
      Top             =   1860
      Width           =   3585
   End
   Begin VB.TextBox txtTermsOfPayment 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6660
      TabIndex        =   3
      Top             =   2460
      Width           =   3585
   End
   Begin VB.TextBox txtPortLoading 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11760
      TabIndex        =   2
      Top             =   1800
      Width           =   2205
   End
   Begin VB.TextBox txtPortDischarge 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11760
      TabIndex        =   1
      Top             =   2160
      Width           =   2205
   End
   Begin VB.TextBox txtNet 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   11700
      TabIndex        =   0
      Top             =   9360
      Width           =   1035
   End
   Begin Crystal.CrystalReport CR 
      Left            =   13260
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   4065
      Left            =   180
      TabIndex        =   20
      Top             =   3600
      Width           =   12585
      _cx             =   22199
      _cy             =   7170
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
      FormatString    =   $"frmExportInvoice.frx":4B65
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
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   4155
         Begin MSDataListLib.DataCombo cboItem 
            Height          =   2310
            Left            =   0
            TabIndex        =   22
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
   Begin MSMask.MaskEdBox dateInv 
      Height          =   315
      Left            =   3180
      TabIndex        =   23
      Top             =   300
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
      Left            =   3180
      TabIndex        =   24
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
      Left            =   3180
      TabIndex        =   25
      Top             =   1200
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
   Begin VB.Label Label7 
      Caption         =   "Total Pcs"
      Height          =   255
      Left            =   10320
      TabIndex        =   61
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Total Cartons"
      Height          =   195
      Left            =   10320
      TabIndex        =   59
      Top             =   8400
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Total CBM"
      Height          =   195
      Left            =   10320
      TabIndex        =   58
      Top             =   8700
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Net Weight (Kg.)"
      Height          =   255
      Left            =   10320
      TabIndex        =   57
      Top             =   9000
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "EPCG License No."
      Height          =   195
      Left            =   180
      TabIndex        =   53
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label ExportCurrency 
      AutoSize        =   -1  'True
      Caption         =   "Export Currency"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   51
      Top             =   2760
      Width           =   1125
   End
   Begin VB.Label CountryofDestination 
      AutoSize        =   -1  'True
      Caption         =   "Country of Final Destination"
      Height          =   195
      Index           =   1
      Left            =   4620
      TabIndex        =   48
      Top             =   2820
      Width           =   1935
   End
   Begin VB.Label BuyerBatchNum 
      AutoSize        =   -1  'True
      Caption         =   "Buyer's Batch No."
      Height          =   195
      Index           =   1
      Left            =   4620
      TabIndex        =   47
      Top             =   1560
      Width           =   1875
   End
   Begin VB.Label BuyersOrderNum 
      AutoSize        =   -1  'True
      Caption         =   "Buyer's Order No(s)."
      Height          =   195
      Index           =   1
      Left            =   4620
      TabIndex        =   46
      Top             =   1260
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "Name of Dealer/Buyer :"
      Height          =   300
      Index           =   2
      Left            =   4620
      TabIndex        =   44
      Top             =   300
      Width           =   1785
   End
   Begin VB.Line Line5 
      X1              =   11100
      X2              =   19920
      Y1              =   -540
      Y2              =   -540
   End
   Begin VB.Label OrderDate 
      AutoSize        =   -1  'True
      Caption         =   "Order Date(s)"
      Height          =   195
      Left            =   14400
      TabIndex        =   43
      Top             =   -360
      Width           =   945
   End
   Begin VB.Label ExporterPan 
      AutoSize        =   -1  'True
      Caption         =   "Exporter's Pan"
      Height          =   195
      Left            =   180
      TabIndex        =   42
      Top             =   2340
      Width           =   1020
   End
   Begin VB.Label ExportCurrency 
      AutoSize        =   -1  'True
      Caption         =   "Brand"
      Height          =   195
      Index           =   0
      Left            =   4620
      TabIndex        =   41
      Top             =   1860
      Width           =   1920
   End
   Begin VB.Label PreCarriageBy 
      AutoSize        =   -1  'True
      Caption         =   "Pre Carriage By"
      Height          =   195
      Left            =   4620
      TabIndex        =   40
      Top             =   3120
      Width           =   1875
   End
   Begin VB.Label PlaceofCarriage 
      AutoSize        =   -1  'True
      Caption         =   "Place of Recipt by PreCarrige"
      Height          =   390
      Left            =   10320
      TabIndex        =   39
      Top             =   1320
      Width           =   1815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Invoice No:"
      Height          =   270
      Index           =   0
      Left            =   210
      TabIndex        =   35
      Top             =   300
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "Date "
      Height          =   270
      Index           =   1
      Left            =   2640
      TabIndex        =   34
      Top             =   300
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Exporter's I.E. code"
      Height          =   300
      Index           =   4
      Left            =   180
      TabIndex        =   33
      Top             =   1980
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Date &&  Time of issue of invoice"
      Height          =   270
      Index           =   7
      Left            =   180
      TabIndex        =   32
      Top             =   840
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Date &&  Time of Dispatch"
      Height          =   270
      Index           =   8
      Left            =   180
      TabIndex        =   31
      Top             =   1200
      Width           =   2835
   End
   Begin VB.Label Label1 
      Caption         =   "Terms of Delivery"
      Height          =   210
      Index           =   11
      Left            =   4620
      TabIndex        =   30
      Top             =   2160
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "Port of Loading"
      Height          =   270
      Index           =   13
      Left            =   10320
      TabIndex        =   29
      Top             =   1800
      Width           =   1530
   End
   Begin VB.Label Label1 
      Caption         =   "Terms of Payment"
      Height          =   210
      Index           =   14
      Left            =   4620
      TabIndex        =   28
      Top             =   2460
      Width           =   1890
   End
   Begin VB.Label Label1 
      Caption         =   "Port of Discharge"
      Height          =   270
      Index           =   15
      Left            =   10320
      TabIndex        =   27
      Top             =   2100
      Width           =   1530
   End
   Begin VB.Label Label2 
      Caption         =   "Gross Weight (Kg.)"
      Height          =   255
      Left            =   10320
      TabIndex        =   26
      Top             =   9360
      Width           =   1395
   End
End
Attribute VB_Name = "frmExportInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim itemgp As String
'Dim rates As Double
'Dim i As Integer
'Dim Status As String
'Dim Item_Name As String
'Dim unit As String
'Dim qty As Integer
'Dim iitem1 As String
'Dim StockFlag As String
'Dim edit As Boolean
'Const gl As String = "SUNDRY DEBTORS"
'Const VAT As Double = 4.5
'Const VAT1 As Double = 2
'
'Dim total1 As Double
'Dim VAT_less As Double
'Dim VAT_Add As Double
'
'
'Dim header As String
'
'Const dis1 As Double = 25
'Const dis2 As Double = 9.09
'Const dis3 As Double = 7.5
'Const dis4 As Double = 5
'
'
'
'Private Sub cmdMain_Click()
'Unload Me
'End Sub
'Sub cellposi()
' 'VsFrame.Width = 3165
' VsFrame.TOP = vs.TOP + ((vs.CellTop)) - 1400
' VsFrame.Left = (vs.Left) - 200
'End Sub
'Sub Total()
'
'On Error Resume Next
'
'txtTotal.Text = 0
'
'
'total1 = 0
'
'
'For J = 1 To vs.Rows - 1
'If vs.TextMatrix(J, 1) <> "" Then
'txtTotal.Text = (Val(txtTotal.Text) + Val(vs.TextMatrix(J, 6)))
'End If
'Next
'
'
'txtTotal = Format(txtTotal, ".00")
'
'
'
'total1 = (Val(txtTotal) - Val(txtamount())
'
'txtamount(0) = Round((total1 * Val(txtRate(0)) / 100), 2)
'txtamount(0) = Format(txtamount(0), ".00")
'
'txtNet = Format((Val(total1) + Val(txtamount(0))), ".00")
'
'
'
'End Sub
'
'Sub cellposiVs()
' Vs1Frame.Width = 2500
' Vs1Frame.TOP = vs1.TOP + ((vs1.CellTop))
' Vs1Frame.Left = (vs1.Left) + 550
'End Sub
'Sub AddItemInGrid1()
''
''    Dim rs_4 As New ADODB.Recordset
''
''    rs_4.Open "select * from bm order by BKDESC", con, adOpenDynamic, adLockOptimistic
''
''    Set cboitemvs1.RowSource = rs_4
''    cboitemvs1.ListField = "BKDESC"
''    cboitemvs1.BoundColumn = "BKCODE"
''    cboitemvs1.ReFill
'
'End Sub
'Sub AddItemInGrid3()
'    'Adodc1.ConnectionString = "filedsn=Saru"
'    'Adodc1.CommandType = adCmdText
''
''    Dim rs_3 As New ADODB.Recordset
''
''    'rs_3.Open "select * from ItemMaster where (ItemGp='Finish Item' or ItemGp='Scrap' or ItemGp='Losses' or (ItemGp='Chunnus (Pointing)' or ItemGp='Chunnus (Sonicut)') or  Itemgp='Rejuction') order by ItemName", con, adOpenDynamic, adLockOptimistic
''    rs_3.Open "select * from ItemMaster order by ItemName", con, adOpenDynamic, adLockOptimistic
''
''    'Adodc1.Refresh
''    Set cboItemVs3.RowSource = rs_3
''    cboItemVs3.ListField = "ItemName"
''    cboItemVs3.BoundColumn = "ItemName"
''    cboItemVs3.ReFill
'
'End Sub
'Sub AddItemInGrid2()
''    'Adodc1.ConnectionString = "filedsn=Saru"
''    'Adodc1.CommandType = adCmdText
''    Dim rs_2 As New ADODB.Recordset
''
''    'rs_2.Open "select * from ItemMaster where (ItemGp='Semi Finish (R/D)' or ItemGp= 'Semi Finish (Store)' or Itemgp='Rejuction') order by ItemName", con, adOpenDynamic, adLockOptimistic
''    rs_2.Open "select * from ItemMaster order by ItemName", con, adOpenDynamic, adLockOptimistic
''
''    Set cboItemVs2.RowSource = rs_2
''    cboItemVs2.ListField = "ItemName"
''    cboItemVs2.BoundColumn = "ItemName"
''    cboItemVs2.ReFill
'
'End Sub
'Sub AddItemInGrid()
''    'Adodc1.ConnectionString = "filedsn=Saru"
''    'Adodc1.CommandType = adCmdText
''    Dim rs_1 As New ADODB.Recordset
''
''    'rs_1.Open "select * from ItemMaster where ItemGp='Raw Item' or ItemGp='Scrap' or (ItemGp='Chunnus (Pointing)' or ItemGp='Chunnus (Sonicut)') order by ItemName", con, adOpenDynamic, adLockOptimistic
''    rs_1.Open "select * from bm order by BKDESC", CON, adOpenDynamic, adLockOptimistic
''
''    Set cboItem.RowSource = rs_1
''    cboItem.ListField = "BKDESC"
''    cboItem.BoundColumn = "BKCODE"
''    cboItem.ReFill
'
'End Sub
'Private Sub cboFinishItem_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then txtqty.SetFocus
'End Sub
'
'
'Private Sub cbogodown_KeyDown(KeyCode As Integer, Shift As Integer)
' If KeyCode = 13 Then txtRemarks.SetFocus
'End Sub
'
'Private Sub cbogodown_LostFocus()
'If cboGodown = "" Then
'   MsgBox "Select Godown Name ..", vbCritical
'   cboGodown.SetFocus
'   Exit Sub
'End If
'End Sub
'
'Private Sub cboItem_KeyDown(KeyCode As Integer, Shift As Integer)
'     If KeyCode = 13 Then
'        cellposi
'        If cboItem.Text = "" Then
'        VsFrame.Visible = False
'        cmdSave_2.SetFocus
'        Exit Sub
'        End If
'        vs.TextMatrix(vs.RowSel, 0) = cboItem.Text
'        vs.TextMatrix(vs.RowSel, 6) = cboItem.BoundText
'
'
'        vs.SetFocus
'
'     ElseIf KeyCode = 27 Then
'
'          VsFrame.Visible = False
'
'     End If
'End Sub
'Sub saveInMaster()
'         On Error Resume Next
'
'         If rs.State = 1 Then rs.Close
'         rs.Open "select * from ItemMaster where ItemName='" & iitem1 & "'", CON, adOpenDynamic, adLockOptimistic
'         If rs.EOF = True Then
'            rs.AddNew
'            rs.Fields("ItemGp").Value = frmAddMaster.cboGp.Text
'            rs.Fields("ItemName").Value = iitem1
'            rs.Fields("Unit").Value = "Kg"
'            rs.Update
'         Else
'            MsgBox "This Item Already Exist !!", vbCritical
'            Exit Sub
'         End If
'         frmAddMaster.Visible = False
'
'End Sub
'Private Sub cboitemvs1_KeyDown(KeyCode As Integer, Shift As Integer)
'     If KeyCode = 13 Then
'        cellposiVs
'        vs1.TextMatrix(vs1.RowSel, 0) = cboitemvs1.Text
'
'        If rs.State = 1 Then rs.Close
'        rs.Open "select * from ItemMaster where ItemName='" & cboitemvs1.Text & "'", CON
'        If rs.EOF = True Then
'            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
'                frmAddMaster.Show 1
'                iitem1 = cboitemvs1.Text
'                saveInMaster
'                cboitemvs1.Text = ""
'                Vs1Frame.Visible = False
'                vs1.SetFocus
'             End If
'        End If
'        vs1.SetFocus
'     ElseIf KeyCode = 27 Then
'        Vs1Frame.Visible = False
'     End If
'End Sub
'
'
'Private Sub cboItemVs2_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'
'
'        'cellposiVs2
'        vs3.TextMatrix(vs3.RowSel, 0) = cboItemVs2.Text
'        Set rs = New ADODB.Recordset
'        If rs.State = 1 Then rs.Close
'        rs.Open "select * from ItemMaster where ItemName='" & cboItemVs2.Text & "'", CON
'        If rs.EOF = True Then
'            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
'                frmAddMaster.Show 1
'                iitem1 = cboItemVs2.Text
'                saveInMaster
'
'                cboItemVs2.Text = ""
'             End If
'        End If
'        vs3.SetFocus
'
'ElseIf KeyCode = 27 Then
'         FrameVs2.Visible = False
'End If
'End Sub
'
'Private Sub cboItemVs3_KeyDown(KeyCode As Integer, Shift As Integer)
'     If KeyCode = 13 Then
'        'cellposiVs3
'
'        vs2.TextMatrix(vs2.RowSel, 0) = cboItemVs3.Text
'        Set rs = New ADODB.Recordset
'        If rs.State = 1 Then rs.Close
'        rs.Open "select * from ItemMaster where ItemName='" & cboItemVs3.Text & "'", CON
'        If rs.EOF = True Then
'            If MsgBox("Want To Add In Master ?", vbQuestion + vbYesNo) = vbYes Then
'                Vs3Frame.Visible = False
'                iitem1 = cboItemVs3.Text
'                frmAddMaster.Show 1
'                saveInMaster
'                cboItemVs3.Text = ""
'                vs2.SetFocus
'             End If
'        End If
'        Vs3Frame.Visible = False
'        'cboItemVs3.Visible = False
'        vs2.SetFocus
'     ElseIf KeyCode = 27 Then
'        Vs3Frame.Visible = False
'     End If
'
'End Sub
'
'Private Sub cmdadd_Click()
' If rs.State = 1 Then rs.Close
' rs.Open "select HeatingNo from IssueMaster where HeatingDate >=datevalue('" & fromdate.Value & "') and HeatingDate <=datevalue('" & todate.Value & "') order by HeatingNo", CON
' ListHeatingNo.Clear
' If rs.EOF = False Then
'    While rs.EOF = False
'       ListHeatingNo.AddItem rs(0)
'       rs.MoveNext
'    Wend
' End If
'End Sub
'Private Sub cmdDelete_Click()
'
'''   If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
'''
'''      DeleteRecord txtHeating.Text, "HeatingNo", "IssueMaster"
'''      DeleteRecord txtHeating.Text, "HeatingNo", "IssueRawMetrial"
'''      Call cmdref_Click
'''
'''   End If
'End Sub
'Sub DeleteStock()
'
''''Dim rr As New ADODB.Recordset
''''Dim rs_u As New ADODB.Recordset
''''Dim openning As Double
''''
''''
''''
''''
''''
'''''================ Issue For Casting
''''
''''
'''' If StockFlag = "1" Then
''''
''''    If rs_u.State = 1 Then rs_u.Close
''''    rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
''''    If rs_u.EOF = False Then
''''        rs_u!qty = rs_u!qty + qty
''''        rs_u.Update
''''    End If
''''
'''' End If
''''
''''
''''
'''' '================ Receive For Casting
''''
'''' If StockFlag = "2" Then
''''
''''   If itemgp <> "Losses" Then
''''
''''        If rs_u.State = 1 Then rs_u.Close
''''        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
''''        If rs_u.EOF = False Then
''''            rs_u!qty = rs_u!qty - qty
''''            rs_u.Update
''''        End If
''''
''''    End If
''''
'''' End If
''''
''''
''''
'''' '================ Receive For Finish
''''
'''' If StockFlag = "3" Then
''''
''''   If itemgp <> "Losses" Then
''''
''''        If rs_u.State = 1 Then rs_u.Close
''''        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
''''        If rs_u.EOF = False Then
''''             rs_u!qty = rs_u!qty - qty
''''            rs_u.Update
''''        End If
''''
''''    End If
''''
'''' End If
''''
''''
''''
'''' '================ Issue For Finish
''''
'''' If StockFlag = "4" Then
''''
''''   If itemgp <> "Losses" Then
''''
''''        If rs_u.State = 1 Then rs_u.Close
''''        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", con, adOpenDynamic, adLockOptimistic
''''        If rs_u.EOF = False Then
''''             rs_u!qty = rs_u!qty + qty
''''            rs_u.Update
''''        End If
''''
''''    End If
''''
'''' End If
''''
'''' '====================================
''''
''''
''''
''''
''''
'
'
'
'
'End Sub
'Private Sub cmdexit_Click()
'Unload Me
'End Sub
'
'Private Sub cmdFatch_Click()
'AddSemifinish
''Total4
'
'End Sub
'
'Private Sub cmdFind_Click()
' Frame1.Visible = True
' fromdate.SetFocus
'End Sub
'
'Private Sub cmdModify_Click()
''   If MsgBox("Want To Modify ?", vbQuestion + vbYesNo) = vbYes Then
''
''
''      DeleteRecord txtHeating.Text, "HeatingNo", "IssueMaster"
''      DeleteRecord txtHeating.Text, "HeatingNo", "IssueRawMetrial"
''
''
''
''      'SaveData
''
''      UpdateIssue
''
''      Call cmdref_Click
''   End If
'End Sub
'Sub UpdateIssue()
'
'Dim rss As New ADODB.Recordset
'Dim search As New ADODB.Recordset
'
'If search.State = 1 Then search.Close
'search.Open "select ItemName,qty from Invoice where HeatNo='" & txtHeating.Text & "'", CON
'If search.EOF = False Then
'While search.EOF = False
'
'    If rss.State = 1 Then rss.Close
'    rss.Open "select * from IssueRawMetrial where HeatingNo=" & txtHeating.Text & " and ItemName='" & search.Fields(0).Value & "'", CON, adOpenDynamic, adLockOptimistic
'    If rss.EOF = False Then
'       rss.Fields("Issue").Value = (CDbl(rss.Fields("Issue").Value) + CDbl(search.Fields("qty").Value))
'       rss.Update
'    End If
'
'    search.MoveNext
'
'Wend
'
'End If
'
'End Sub
'
'Private Sub cmdref_Click()
'      txtHeating.Text = ""
'      txtParty.Text = ""
'
'      txtRemarks.Text = ""
'
'
'      txtTotal1.Text = 0
'      txtTotal2.Text = 0
'      txtTotal3.Text = 0
'      txtTotal4.Text = 0
'
'      txtSize.Text = ""
'      txtGrade.Text = ""
'      txtRawAndCasting.Text = 0
'
'      vs.Clear
'      vs1.Clear
'      vs2.Clear
'      vs3.Clear
'
'      setwidth
'      txtHeating.SetFocus
'      cmdDelete.Enabled = False
'      cmdModify.Enabled = False
'      cmdSave.Enabled = True
'
'      Record = ""
'
'End Sub
'
'
'Private Sub Command4_Click()
'   Unload Me
'End Sub
'Private Sub CmdSave_Click()
'
'
'
'
'    If rs.State = 1 Then rs.Close
'    rs.Open "select * from IssueMaster where HeatingNo=" & txtHeating.Text & "", CON
'    If rs.EOF = False Then
'       MsgBox "Heating No. Already Exist !!", vbInformation
'       Exit Sub
'    End If
'
'    If txtHeating.Text = "" Then
'       MsgBox "Please Enter Heating No !!", vbCritical
'       txtHeating.SetFocus
'       Exit Sub
'    End If
'
'
'    Set rs = New ADODB.Recordset
'    If rs.State = 1 Then rs.Close
'    rs.Open "select * from IssueMaster where HeatingNo=" & txtHeating.Text & "", CON
'    If rs.EOF = True Then
'       If MsgBox("Want To Save ?", vbQuestion + vbYesNo) = vbYes Then
'        '  SaveData
'       End If
'    Else
'          MsgBox "Dublicate Heating No !!", vbCritical
'    End If
'End Sub
'Sub ItemGpSearch(Str As String)
'
'    If rs1.State = 1 Then rs1.Close
'    rs1.Open "select ItemGp,Rate from ItemMaster where ItemName='" & Str & "'", CON
'    If rs1.EOF = False Then
'       itemgp = rs1.Fields(0).Value
'       rates = rs1.Fields(1).Value
'    End If
'
'End Sub
'Sub UpdateStock()
'    Dim rr As New ADODB.Recordset
'    Dim rs_u As New ADODB.Recordset
'    Dim openning As Double
'
'
'
'
' '================ Issue For Casting
'
'
' If StockFlag = "1" Then
'
'    If rs_u.State = 1 Then rs_u.Close
'    rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
'    If rs_u.EOF = True Then
'        rs_u.AddNew
'        rs_u!itemname = Item_Name
'        ItemGpSearch Item_Name
'        rs_u!itemgp = itemgp
'        rs_u!unit = unit
'        rs_u!rate = rates
'        rs_u!qty = (-1 * qty)
'        rs_u.Update
'     Else
'        rs_u!qty = rs_u!qty - qty
'        rs_u.Update
'    End If
'
' End If
'
' '====================================
'
'
' '================ Receive For Casting
'
' If StockFlag = "2" Then
'
'   If itemgp <> "Losses" Then
'
'        If rs_u.State = 1 Then rs_u.Close
'        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
'        If rs_u.EOF = True Then
'            rs_u.AddNew
'            rs_u!itemname = Item_Name
'            ItemGpSearch Item_Name
'            rs_u!itemgp = itemgp
'            rs_u!unit = unit
'            rs_u!rate = rates
'            rs_u!qty = qty
'            rs_u.Update
'         Else
'            rs_u!qty = rs_u!qty + qty
'            rs_u.Update
'        End If
'
'    End If
'
' End If
'
' '====================================
'
'
' '================ Receive For Finish
'
' If StockFlag = "3" Then
'
'   If itemgp <> "Losses" Then
'
'        If rs_u.State = 1 Then rs_u.Close
'        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
'        If rs_u.EOF = True Then
'            rs_u.AddNew
'            rs_u!itemname = Item_Name
'            ItemGpSearch Item_Name
'            rs_u!itemgp = itemgp
'            rs_u!unit = unit
'            rs_u!rate = rates
'            rs_u!qty = qty
'            rs_u.Update
'         Else
'            rs_u!qty = rs_u!qty + qty
'            rs_u.Update
'        End If
'
'    End If
'
' End If
'
' '====================================
'
'
'  '================ Issue For Finish
'
' If StockFlag = "4" Then
'
'   If itemgp <> "Losses" Then
'
'        If rs_u.State = 1 Then rs_u.Close
'        rs_u.Open "select * from Stock where ItemName='" & Item_Name & "' and ItemGp='" & itemgp & "'", CON, adOpenDynamic, adLockOptimistic
'        If rs_u.EOF = True Then
'            rs_u.AddNew
'            rs_u!itemname = Item_Name
'            ItemGpSearch Item_Name
'            rs_u!itemgp = itemgp
'            rs_u!unit = unit
'            rs_u!rate = rates
'            rs_u!qty = (-1 * qty)
'            rs_u.Update
'         Else
'            rs_u!qty = rs_u!qty - qty
'            rs_u.Update
'        End If
'
'    End If
'
' End If
'
' '====================================
'
'
'
'
'End Sub
'
'
'
'
'
'
'Private Sub cmdAdd_1_Click()
'
'
'Dim o As Object
'For Each o In Me
'If TypeOf o Is textbox Then
'o.Text = ""
'End If
'
'If TypeOf o Is MaskEdBox Then
'o.Text = "__/__/____"
'End If
'
'Next
'
'Check_Nontaxable.Value = False
'
'ptype.Caption = ""
'
'dateInv = Date
'
'dateInv.SetFocus
'txtRate(0) = VAT
'txtCenteral = "4/2006 CE Dt. 01/03/06 S.No.97"
'txtModePay = "BY Road"
'
'   vs.Clear
'   setwidth
'   cmdDelete_3.Enabled = False
'   cmdEdit_4.Enabled = True
'   'cmdPrint_7.Enabled = False
'   cmdSave_2.Enabled = True
'   cmdAdd_1.Enabled = True
'   cmdExit_12.Enabled = True
'   txtino.Text = MaxSNo("invoicea", "INVOICENO")
'End Sub
'
'
'
'Private Sub cmdDelete_3_Click()
'If MsgBox("Want To Delete ?", vbQuestion + vbYesNo) = vbYes Then
'
'   CON.BeginTrans
'   CON.Execute "delete from invoicea where invoiceNo =" & txtino & ""
'   CON.Execute "delete from invoiceb where invoiceNo =" & txtino & ""
'   CON.Execute "delete from invoicec where invoiceNo =" & txtino & ""
'   CON.CommitTrans
'
'   Call cmdAdd_1_Click
'End If
'cmdEdit_4.Enabled = False
'cmdDelete_3.Enabled = False
'End Sub
'
'Private Sub cmdEdit_4_Click()
'   cmdDelete_3.Enabled = True
'   cmdEdit_4.Enabled = False
'   cmdSave_2.Enabled = True
'   cmdAdd_1.Enabled = False
'   cmdExit_12.Enabled = True
'   edit = True
'End Sub
'
'Private Sub cmdExit_12_Click()
'Unload Me
'End Sub
'
'Private Sub cmdPrint_7_Click()
'
'CR.Reset
'CR.ReportFileName = App.Path & "/Reports/CHALLAN.rpt"
'CR.ReplaceSelectionFormula "{invoiceA.invoiceno}=" & txtHeating.Text & ""
'CR.WindowShowPrintSetupBtn = True
'CR.WindowState = crptMaximized
'CR.Action = 1
'
'End Sub
'
'
'''Sub searchData()
'''
'''If rs.State = 1 Then rs.Close
'''rs.Open "select * from invoicea where INVOICENO=" & txtHeating.Text & "", CON, adOpenDynamic, adLockOptimistic
'''If rs.EOF = False Then
'''txtParty.Text = rs.Fields("SUBLEDGER").Value
'''txtRemarks.Text = rs.Fields("Remarks").Value & ""
'''Text1.Text = rs.Fields("add1").Value & ""
'''Text2.Text = rs.Fields("add2").Value & ""
'''If Not IsNull(rs.Fields("godown").Value) Then
'''cboGodown = rs.Fields("godown").Value & ""
'''Else
'''cboGodown.ListIndex = -1
'''End If
'''
'''End If
'''
'''
'''
'''If rs.State = 1 Then rs.Close
'''rs.Open "select * from invoiceb where INVOICENO=" & txtHeating.Text & "", CON, adOpenDynamic, adLockOptimistic
'''For i = 1 To rs.RecordCount
'''If rs.EOF = False Then
'''vs.TextMatrix(i, 0) = rs.Fields("BOOKCODE").Value
'''vs.TextMatrix(i, 1) = rs.Fields("TBook").Value
'''vs.TextMatrix(i, 2) = rs.Fields("LoosBook").Value
'''vs.TextMatrix(i, 3) = rs.Fields("TotalBook").Value
'''vs.TextMatrix(i, 4) = rs.Fields("NetBook").Value
'''vs.TextMatrix(i, 5) = rs.Fields("remarks").Value & ""
'''vs.TextMatrix(i, 6) = rs.Fields("Book_Code").Value & ""
'''rs.MoveNext
'''End If
'''Next
'''
'''Total
'''
'''End Sub
'Private Sub cmdUndo_5_Click()
'   cmdDelete_3.Enabled = False
'   cmdEdit_4.Enabled = False
'   cmdPrint_7.Enabled = True
'   cmdSave_2.Enabled = False
'   cmdUndo_5.Enabled = False
'   cmdAdd_1.Enabled = True
'   cmdExit_12.Enabled = True
'End Sub
'
'
'
'Private Sub Dates_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then txtParty.SetFocus
'End Sub
'
'Private Sub cmdprint_Click()
'    CR.Reset
'    CR.Connect = constr
'    CR.ReportFileName = strrptpath & "\REPORTS\Exinvoice.rpt"
'    CR.ReplaceSelectionFormula "{invoicea.invoiceno} = " & txtino & " AND {invoicea.setupid} = " & main.setupid & " AND {invoicea.fyear} = '" & main.session & "'"
'
'    CR.Formulas(0) = "tax_label='" & tax(0) & "'"
'    CR.Formulas(1) = "tax_rate='" & txtRate(0) & "'"
'    CR.Formulas(2) = "tax_amount='" & txtamount(0) & "'"
'
'    CR.Formulas(3) = "tax_label1='" & tax(1) & "'"
'    CR.Formulas(4) = "tax_rate1='" & txtRate(1) & "'"
'    CR.Formulas(5) = "tax_amount1='" & Format(txtamount(1), ".00") & "'"
'    CR.Formulas(6) = "header='" & header & "'"
'
'    CR.WindowState = crptMaximized
'    CR.Action = 1
'End Sub
'
'Private Sub cmdSave_2_Click()
'
'
''On Error GoTo save:
'
'Dim n As Date
'Dim i As Integer
'Dim netrate As String
'Dim with_without As String
'
'
'If Check_Net.Value = 1 Then
'   netrate = "y"
'   Else
'   netrate = "n"
'End If
'
'If Option_with.Value = True Then
'with_without = 1
'Else
'with_without = 2
'End If
'
'
'If Not IsDate(dateIssue.Text) Then
'  MsgBox "Enter Issue Date ...", vbCritical
'  dateIssue.SetFocus
'  Exit Sub
'End If
'
'If Not IsDate(dateDispatch.Text) Then
'  MsgBox "Enter Dispatch Date ...", vbCritical
'  dateDispatch.SetFocus
'  Exit Sub
'End If
'
'
'If Not IsDate(dateRR.Text) Then
'  MsgBox "Enter RR/LR Date ...", vbCritical
'  dateRR.SetFocus
'  Exit Sub
'End If
'
'Dim nontax As String
'
'If Check_Nontaxable.Value = 0 Then
'    nontax = "n"
'  Else
'    nontax = "y"
'End If
'
'
'
'
'
'i = 1
'
'If edit = False Then
'
'            txtino = MaxSNo("invoicea", "INVOICENO")
'            CON.BeginTrans
'
'             CON.Execute "exec insertData_Invoicea " & txtino & ",'" & dateInv & "','" & dateIssue & "'," & _
'            "'" & dateDispatch & "','" & txtCenteral.Text & "','" & txtParty & "','" & txtDest & "'," & _
'            "'" & txtModePay & "','" & txtTrans & "','" & txtBoxes & "','" & txtWeight & "','" & txtRR & "','" & dateRR.Text & "'," & _
'            "'" & txtFrieght & "','" & txtWagon & "','" & Val(txtTotal) & "','" & Val(txtNet) & "','" & with_without & "','" & netrate & "','" & main.username & "','" & main.username & "','" & main.session & "'," & _
'            "" & main.setupid & ""
'
'            'CON.Execute "update invoicea  set NonTaxable = '" & nontax & "' where invoiceno=" & txtino & ""
'
'
'
'            'For i = 1 To vs.Rows - 1
'
'            'If vs.TextMatrix(i, 1) <> "" Then
'
'            'CON.Execute "exec insertData_Invoiceb " & txtino & ",'" & dateInv & "','" & txtParty & "'," & _
'            '"" & Val(vs.TextMatrix(i, 0)) & ",'" & vs.TextMatrix(i, 1) & "'," & Val(vs.TextMatrix(i, 3)) & "," & _
'            '"" & Val(vs.TextMatrix(i, 4)) & "," & Val(vs.TextMatrix(i, 5)) & "," & Val(vs.TextMatrix(i, 6)) & "," & _
'            '"'" & main.username & "','" & main.username & "','" & main.session & "','" & vs.TextMatrix(i, 2) & "'," & _
'            '"" & main.setupid & ""
'
'            'End If
'
'            'Next
'
''            For J = 0 To 1
''
''            If J = 0 Then
''              ST = "Debit"
''            ElseIf J = 1 Then
''            If Val(txtamount(J)) < 0 Then
''              ST = "Credit"
''            Else
''              ST = "Debit"
''            End If
''            End If
''
''
''            CON.Execute "insert into invoicec" & _
''            "(INVOICENO,INVOICEDate,GENLEDGER,GAmount,Rate,Amount,DebitOrCredit," & _
''            "text,fyear,createdby,createdon,updatedby,updatedon,setupid) values(" & _
''            "" & txtino & ",'" & Format(dateInv, "MM/dd/yyyy") & "','" & gl & "'," & Val(txtNet) & "," & Val(txtRate(J)) & "," & _
''            "" & Val(txtamount(J)) & ",'" & ST & "','" & tax(J) & "','" & main.session & "','" & main.username & "'," & _
''            "'" & Format(Date, "MM/DD/yyyy") & "','" & main.username & "','" & Format(Date, "MM/dd/yyyy") & "'," & main.setupid & ")"
''
''            Next
'
'
'            CON.CommitTrans
'
'
'
'
'Else
'
'            '-------------------------------------------
'
'
'
'            CON.BeginTrans
'
'
'            CON.Execute "delete from invoicea where invoiceNo =" & txtino & ""
'            CON.Execute "delete from invoiceb where invoiceNo =" & txtino & ""
'            CON.Execute "delete from invoicec where invoiceNo =" & txtino & ""
'
'
'
'            CON.Execute "exec insertData_Invoicea " & txtino & ",'" & dateInv & "','" & dateIssue & "'," & _
'            "'" & dateDispatch & "','" & txtCenteral.Text & "','" & txtParty & "','" & txtDest & "'," & _
'            "'" & txtModePay & "','" & txtTrans & "','" & txtBoxes & "','" & txtWeight & "','" & txtRR & "','" & dateRR.Text & "'," & _
'            "'" & txtFrieght & "','" & txtWagon & "','" & Val(txtTotal) & "','" & Val(txtNet) & "','" & with_without & "','" & netrate & "','" & main.username & "','" & main.username & "','" & main.session & "'," & _
'            "" & main.setupid & ""
'
'            CON.Execute "update invoicea  set NonTaxable = '" & nontax & "' where invoiceno=" & txtino & ""
'
'
'            For i = 1 To vs.Rows - 1
'
'            If vs.TextMatrix(i, 1) <> "" Then
'
'            CON.Execute "exec insertData_Invoiceb " & txtino & ",'" & dateInv & "','" & txtParty & "'," & _
'            "" & Val(vs.TextMatrix(i, 0)) & ",'" & vs.TextMatrix(i, 1) & "'," & Val(vs.TextMatrix(i, 3)) & "," & _
'            "" & Val(vs.TextMatrix(i, 4)) & "," & Val(vs.TextMatrix(i, 5)) & "," & Val(vs.TextMatrix(i, 6)) & "," & _
'            "'" & main.username & "','" & main.username & "','" & main.session & "','" & vs.TextMatrix(i, 2) & "'," & _
'            "" & main.setupid & ""
'
'            End If
'
'            Next
'
'            For J = 0 To 1
'
'            If J = 0 Then
'              ST = "Debit"
'            ElseIf J = 1 Then
'            If Val(txtamount(J)) < 0 Then
'              ST = "Credit"
'            Else
'              ST = "Debit"
'            End If
'            End If
'
'
'            CON.Execute "insert into invoicec" & _
'            "(INVOICENO,INVOICEDate,GENLEDGER,GAmount,Rate,Amount,DebitOrCredit," & _
'            "text,fyear,createdby,createdon,updatedby,updatedon,setupid) values(" & _
'            "" & txtino & ",'" & Format(dateInv, "MM/dd/yyyy") & "','" & gl & "'," & Val(txtNet) & "," & Val(txtRate(J)) & "," & _
'            "" & Val(txtamount(J)) & ",'" & ST & "','" & tax(J) & "','" & main.session & "','" & main.username & "'," & _
'            "'" & Format(Date, "MM/DD/yyyy") & "','" & main.username & "','" & Format(Date, "MM/dd/yyyy") & "'," & main.setupid & ")"
'
'            Next
'
'
'            CON.CommitTrans
'
'
'            edit = False
'
''---------------------------------------------
'
'End If
'
'
'cmdEdit_4.Enabled = True
'cmdDelete_3.Enabled = True
'Call cmdAdd_1_Click
'
''Exit Sub
'
'
''save:
''
''CON.RollbackTrans
''If Err.Number = "-2147217900" Then
''   MsgBox "Duplicate Data ...", vbCritical
''   txtCode.SetFocus
''End If
'
'
'
'
'End Sub
'
'Private Sub dateInv_LostFocus()
'dateIssue.Text = dateInv.Text
'dateDispatch.Text = dateInv.Text
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'
'
'If KeyCode = 27 Then
'     If MsgBox("Want To Exit ?", vbQuestion + vbYesNo) = vbYes Then
'        Unload Me
'     End If
' End If
'
' If KeyCode = 13 Then
'
' If UCase(Trim(VB.Screen.ActiveControl.Name)) <> UCase(Trim("vs")) Then
'    SendKeys "{tab}"
'    HIT
' End If
'
' End If
'
'
'
'
'
'End Sub
'Sub SearchData()
'
'vs.Clear
'setwidth
'
'
'If rs.State = 1 Then rs.Close
'
'st1 = "select INVOICENO,INVOICEDATE,IssueDate,DisPatchDate,CentralExise,SUBLEDGER,STATION," & _
'"ModeOfPayment,THROUGH,BUNDLES,WEIGHT,BILTYNO,BILTYDate,FREIGHT,TXT1,[with_withoutFormc],[NetRate],NonTaxable from invoicea where invoiceno=" & txtino & ""
'
'rs.Open st1, CON
'If rs.EOF = False Then
'
'  dateInv = rs!invoicedate
'  dateIssue = Format(rs!IssueDate, "dd/MM/yyyy")
'  dateDispatch = Format(rs!DisPatchDate, "dd/MM/yyyy")
'  txtCenteral = rs!CentralExise
'  txtParty = rs!subledger
'  txtDest = rs!station
'  txtModePay = rs!ModeOfPayment
'  txtTrans = rs!through
'  txtBoxes = rs!bundles
'  txtWeight = rs!weight
'  dateRR = rs!BILTYDATE
'  txtRR = rs!biltyno
'  txtWagon = rs!txt1
'  txtFrieght = rs!freight
'
'  If rs![with_withoutFormc] = 1 Then
'     Option_with.Value = True
'  Else
'     Option_without.Value = True
'  End If
'  If rs![netrate] = "y" Then
'     Check_Net.Value = 1
'  Else
'     Check_Net.Value = 0
'  End If
'
'  If rs!NonTaxable = "n" Then
'     Check_Nontaxable.Value = 0
'  Else
'     Check_Nontaxable.Value = 1
'  End If
'
'
'
'  If rs.State = 1 Then rs.Close
'   rs.Open "SELECT [SUBLEDGER],address1,address2,[City],[District],[State],[bcityid],typeofcust " & _
'   "from [ExportData].[dbo].[SubledgerQry] WHERE SUBLEDGER ='" & txtParty & "'", CON
'   If rs.EOF = False Then
'     txtAdd = rs![subledger] & " " & vbCrLf & rs!address1 & " " & rs!address1 & vbCrLf & rs![CITY] + vbCrLf + rs![District] & "," & rs![State]
'     Me.ptype.Caption = rs!TypeOfCust
'   End If
'
'
'End If
'
'
'
'If rs.State = 1 Then rs.Close
'rs.Open "select * from invoiceb where INVOICENO=" & txtino.Text & " order by printorder", CON, adOpenDynamic, adLockOptimistic
'For i = 1 To rs.RecordCount
'If rs.EOF = False Then
'
'
'
'
'    vs.TextMatrix(i, 0) = rs.Fields("printorder").Value
'    vs.TextMatrix(i, 1) = rs.Fields("BOOKCODE").Value
'
'    vs.TextMatrix(i, 3) = rs.Fields("Quantity").Value
'    vs.TextMatrix(i, 4) = rs.Fields("Rate").Value
'    vs.TextMatrix(i, 5) = Format(rs.Fields("NetRate").Value, ".00")
'    vs.TextMatrix(i, 6) = Format(rs.Fields("Amount").Value, ".00")
'
'
'
'    If rs1.State = 1 Then rs1.Close
'    rs1.Open "select ProductQuality,TypeofProduct,rulling,rate,NoofPages from copymaster " & _
'    "where bookno='" & vs.TextMatrix(i, 1) & "'", CON
'    If rs1.EOF = False Then
'       vs.TextMatrix(i, 2) = rs1!TypeofProduct + " (" + rs1!rulling + ")" + Str(rs1!NoofPages) + " " + Str(rs1!rate) + " " + rs1!ProductQuality
'    End If
'
'rs.MoveNext
'End If
'Next
'
'
'
'If rs.State = 1 Then rs.Close
'rs.Open "select rate,amount from invoicec where INVOICENO=" & txtino.Text & " order by auto", CON, adOpenDynamic, adLockOptimistic
'If rs.EOF = False Then
'   txtRate(0) = rs(0)
'   rs.MoveNext
'   txtamount(1) = Format(rs(1), ".00")
'End If
'
''-----------------------
'
'If rs.State = 1 Then rs.Close
'rs.Open "select State,tinno from SubledgerQry where subledger='" & txtParty & "'", CON
'If rs.EOF = False Then
'    ST = rs(0)
'    If LCase(ST) = "u.p." Then
'     If Len(rs!tinno) > 0 Then
'       header = "TAX-INVOICE"
'     Else
'       header = "SALE-INVOICE"
'     End If
'
'    Else
'       header = "SALE-INVOICE"
'    End If
'End If
'
''----------------------
'
'Total
'
'
'End Sub
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
'Private Sub Form_Load()
'
' setwidth
'
' dateInv.Text = Format(Date, "dd/MM/yyyy")
'
' txtino = MaxSNo("invoicea", "INVOICENO")
'
' 'txtRate(0) = VAT
'
'
'
' withForm = VAT1
' withoutForm = VAT
'
'End Sub
'Sub setwidth()
'vs.Cols = 7
'vs.FormatString = "S.No.|^Item Code|<Item Name|>Quantity|>MRP|>Net Rate|>Net Amount"
'vs.ColWidth(0) = 300
'vs.ColWidth(1) = 1000
'vs.ColWidth(2) = 6500
'vs.ColWidth(3) = 1000
'vs.ColWidth(4) = 800
'vs.ColWidth(5) = 1000
'vs.ColWidth(6) = 1200
'End Sub
'Private Sub fromdate_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = 13 Then todate.SetFocus
'End Sub
'Private Sub HeatingDate_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = 13 Then txtParty.SetFocus
'End Sub
'
'Private Sub ListHeatingNo_Click()
'  Call cmdref_Click
'  SearchData
'  TotalFinal
'  'Frame1.Visible = False
'End Sub
'Private Sub todate_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = 13 Then Call cmdadd_Click
'End Sub
'Private Sub txtGrade_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Then txtRemarks.SetFocus
'End Sub
'
'Private Sub txtHeating_GotFocus()
'If PopUpValue1 <> "" Then
'txtHeating.Text = PopUpValue1
'Dates.Value = PopUpValue2
'vs.Clear
'setwidth
'SearchData
'PopUpValue1 = ""
'PopUpValue2 = ""
'End If
'End Sub
'
'Private Sub txtHeating_KeyDown(KeyCode As Integer, Shift As Integer)
''If KeyCode = 113 Then
''popuplist2 "select INVOICENO,INVOICEDATE,SUBLEDGER from invoicea order by INVOICENO", CON
''End If
''
''If KeyCode = 13 Then
''SearchData
''End If
'
'End Sub
'
'Private Sub txtHeating_KeyPress(KeyAscii As Integer)
'
'   If KeyAscii = 13 Then
'
'   Dates.SetFocus
'
'
'  End If
'
'
'End Sub
'
'
'Private Sub Option_with_Click()
''If Option_with.Value = True Then
''  txtRate(0) = withForm
''Else
''  txtRate(0) = withoutForm
''End If
'FatchTaxFromSate
'End Sub
'
'Private Sub Option_without_Click()
''If Option_with.Value = True Then
''  txtRate(0) = withForm
''Else
''  txtRate(0) = withoutForm
''End If
'
'FatchTaxFromSate
'End Sub
'
'
'Private Sub txtamount_Change(Index As Integer)
'On Error Resume Next
'
'total1 = (Val(txtTotal) - Val(txtamount(1)))
'
'txtamount(0) = Round((total1 * Val(txtRate(0)) / 100), 3)
'txtamount(0) = Format(txtamount(0), ".00")
'
'txtNet = Format((Val(total1) + Val(txtamount(0))), ".00")
'
'End Sub
'
'Private Sub txtamount_LostFocus(Index As Integer)
'txtamount(1) = Format(txtamount(1), ".00")
'End Sub
'
'Private Sub txtino_GotFocus()
'HIT
'End Sub
'
'Private Sub txtino_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'  SearchData
'End If
'End Sub
'
'Private Sub txtParty_GotFocus()
'HIT
'If PopUpValue1 <> "" Then
'   txtParty = PopUpValue2
'
'
'  If rs.State = 1 Then rs.Close
'   rs.Open "SELECT [SUBLEDGER],address1,address2,[City],[District],[State],[bcityid],typeofcust " & _
'   "from [ExportData].[dbo].[SubledgerQry] WHERE SUBLEDGER ='" & txtParty & "'", CON
'   If rs.EOF = False Then
'     txtAdd = rs![subledger] & " " & vbCrLf & rs!address1 & " " & rs!address1 & vbCrLf & rs![CITY] + vbCrLf + rs![District] & "," & rs![State]
'     Me.ptype.Caption = rs!TypeOfCust
'   End If
'
'
'
'   FatchTaxFromSate
'
'   txtDest.SetFocus
'
'   PopUpValue1 = ""
'   PopUpValue2 = ""
'   PopUpValue3 = ""
'   popupvalue4 = ""
'
'End If
'
'End Sub
'
'Private Sub txtParty_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then Exit Sub
'tblNo = 1
'frmSearchItem.Show
'End Sub
'
'Private Sub txtParty_LostFocus()
'   Record = ""
'End Sub
'Private Sub txtQty_GotFocus()
'     txtqty.SelLength = 10
'End Sub
'
'Private Sub txtQty_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then txtRemarks.SetFocus
'End Sub
'Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'Private Sub txtSize_KeyPress(KeyAscii As Integer)
' If KeyAscii = 13 Then txtGrade.SetFocus
'End Sub
'
'Private Sub vs_AfterEdit(ByVal row As Long, ByVal col As Long)
'     If vs.col = 0 Then
'        cellposi
'        vs.TextMatrix(vs.RowSel, 0) = cboItem.Text
'     End If
'End Sub
'
'Private Sub VS_KeyDown(KeyCode As Integer, Shift As Integer)
'
'
'
'  If KeyCode = 115 Then
'  If MsgBox("Want to delete ?", vbQuestion + vbYesNo) = vbYes Then
'    vs.RemoveItem (vs.RowSel)
'    Total
'  End If
'  End If
'
'  If KeyCode = 13 Then
'
'     If vs.col = 0 Then
'        vs.Editable = flexEDNone
'        VsFrame.Visible = True
'        cboItem.SetFocus
'     Else
'        vs.Editable = flexEDKbdMouse
'        cellposi
'     End If
'
'  End If
'
'
'
'
'
'End Sub
'
'Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)
'
'Dim Item As String
'
'
'If KeyCode = 13 Then
'
' If vs.col = 1 Then
'
'
'
'    If rs.State = 1 Then rs.Close
'    rs.Open "select ProductQuality,TypeofProduct,rulling,rate,NoofPages from copymaster " & _
'    "where bookno='" & vs.TextMatrix(vs.RowSel, 1) & "'", CON
'    If rs.EOF = False Then
'
'          Item = rs!TypeofProduct + " (" + rs!rulling + ")" + Str(rs!NoofPages) + " " + Str(rs!rate) + " " + rs!ProductQuality
'
'          vs.TextMatrix(vs.RowSel, 2) = Item
'
'          vs.TextMatrix(vs.RowSel, 4) = rs.Fields("Rate").Value
'
'    End If
'
'    SendKeys "{right}"
'    SendKeys "{right}"
'
' End If
'
'
' If Check_Net.Value = 1 Then
'
'
' If vs.col = 3 Then
'    vs.TextMatrix(vs.RowSel, 6) = Format((Val(vs.TextMatrix(vs.RowSel, 3)) * Val(vs.TextMatrix(vs.RowSel, 5))), ".00")
'
'    SendKeys "{right}"
'    SendKeys "{right}"
'    Total
' End If
'
' If vs.col = 5 Then
'
'    vs.TextMatrix(vs.RowSel, 6) = Format((Val(vs.TextMatrix(vs.RowSel, 3)) * Val(vs.TextMatrix(vs.RowSel, 5))), ".00")
'    SendKeys "{home}"
'    SendKeys "{down}"
'
'    vs.TextMatrix(vs.RowSel, 0) = vs.row
'
'    Total
'
' End If
'
'
'
'Else
'
'
'
'
' If vs.col = 3 Then
'
'
'
'    Dim amt1, amt2, amt3, amt4 As Double
'
'    amt1 = (Val(vs.TextMatrix(vs.RowSel, 4)) - ((Val(vs.TextMatrix(vs.RowSel, 4)) * dis1) / 100))
'    amt2 = (amt1 - ((amt1 * dis2) / 100))
'    amt3 = (amt2 - ((amt2 * dis3) / 100))
'
'    If ptype.Caption = "Supper Stockist" Then
'       amt4 = (amt3 - ((amt3 * dis4) / 100))
'    Else
'       amt4 = amt3
'    End If
'
'    ss = 0
'
'
'    FatchTaxFromSate
'    If Check_Nontaxable.Value = 0 Then
'            If Option_with.Value = True Then
'
'           ss = ((amt4 * VAT_less) / 100)
'           amt4 = Round((amt4 - ss), 2)
'        Else
'           ss = ((amt4 * VAT_less) / 100)
'           amt4 = Round((amt4 - ss), 2)
'        End If
'    End If
'
'
'
'    vs.TextMatrix(vs.RowSel, 5) = Format(Round(amt4, 2), ".00")
'    vs.TextMatrix(vs.RowSel, 6) = (Val(vs.TextMatrix(vs.RowSel, 3)) * Val(vs.TextMatrix(vs.RowSel, 5)))
'
'    vs.TextMatrix(vs.RowSel, 6) = Format(vs.TextMatrix(vs.RowSel, 6), ".00")
'
'    vs.TextMatrix(vs.RowSel, 0) = vs.row
'    SendKeys "{home}"
'    SendKeys "{down}"
'
'
'    Total
' End If
'
'
'
' End If
'
'
'
'
'
'
'
'
'End If
'
'End Sub
'Sub FatchTaxFromSate()
'
'Dim ST As String
'Dim with_without As String
'
'If Option_with.Value = True Then
'   with_without = Option_with.Caption
'Else
'   with_without = Option_without.Caption
'End If
'
'
'
'
'If rs.State = 1 Then rs.Close
'rs.Open "select State,tinno from SubledgerQry where subledger='" & txtParty & "'", CON
'If rs.EOF = False Then
'   ST = rs(0)
'If LCase(ST) = "u.p." Then
'   tax(0) = "VAT"
' If Len(rs!tinno) > 0 Then
'   'header = "TAX-INVOICE"
' Else
'   'header = "SALE-INVOICE"
' End If
'
'Else
'   'header = "SALE-INVOICE"
'   tax(0) = "CST"
'End If
'
'End If
'
'
'If rs.State = 1 Then rs.Close
'rs.Open "select add_val,less_val from [state_tax_list] where statename='" & ST & "' and with_without='" & with_without & "'", CON
'If rs.EOF = False Then
'   VAT_Add = rs(0)
'   VAT_less = rs(1)
'   txtRate(0) = VAT_Add
'End If
'
'If Check_Nontaxable.Value = 1 Then
'   txtRate(0) = 0
'End If
'
'
'End Sub
'
'Private Sub vs_KeyUpEdit(ByVal row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
'
''If KeyCode = 13 Then
''
''If vs.Col = 1 Then
''If Not (KeyCode >= 48 And KeyCode <= 57) Then
''MsgBox "Enter Only Numeric Value !!", vbInformation
''vs.TextMatrix(vs.RowSel, 1) = ""
''vs.SetFocus
''Exit Sub
''End If
''End If
''
''
''If vs.Col = 2 Then
''If Not (KeyCode >= 48 And KeyCode <= 57) Then
''MsgBox "Enter Only Numeric Value !!", vbInformation
''vs.TextMatrix(vs.RowSel, 2) = ""
''vs.SetFocus
''Exit Sub
''End If
''End If
''
''
''If vs.Col = 3 Then
''If Not (KeyCode >= 48 And KeyCode <= 57) Then
''MsgBox "Enter Only Numeric Value !!", vbInformation
''vs.TextMatrix(vs.RowSel, 3) = ""
''vs.SetFocus
''Exit Sub
''End If
''End If
''
''
''If vs.Col = 4 Then
''If Not (KeyCode >= 48 And KeyCode <= 57) Then
''MsgBox "Enter Only Numeric Value !!", vbInformation
''vs.TextMatrix(vs.RowSel, 4) = ""
''vs.SetFocus
''Exit Sub
''End If
''End If
''
''
''
''
''End If
'
'End Sub
'
'Private Sub vs_LeaveCell()
'  Total
'End Sub
'
'Private Sub vs1_AfterEdit(ByVal row As Long, ByVal col As Long)
'     If vs1.col = 0 Then
'        cellposiVs
'        vs1.TextMatrix(vs1.RowSel, 0) = cboitemvs1.Text
'     End If
'
'End Sub
'Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode = 46 Then
'    vs1.RemoveItem (vs1.RowSel)
'    'Total1
'    TotalFinal
'  End If
'
'  If KeyCode = 13 Then
'     If vs1.col = 0 Then
'        vs1.Editable = flexEDNone
'        Vs1Frame.Visible = True
'        cboitemvs1.Visible = True
'        cboitemvs1.SetFocus
'     Else
'        vs1.Editable = flexEDKbdMouse
'        cellposiVs
'     End If
'  End If
'End Sub
'
'Private Sub vs1_KeyUp(KeyCode As Integer, Shift As Integer)
'
'If KeyCode = 13 Then
'
' If vs1.col = 0 Then
'    vs1.Editable = flexEDNone
'    Vs1Frame.Visible = True
'    cboitemvs1.SetFocus
'
'    Set rs = New ADODB.Recordset
'    If rs.State = 1 Then rs.Close
'    rs.Open "select * from ItemMaster where ItemName='" & vs1.TextMatrix(vs1.RowSel, 0) & "'", CON
'    If rs.EOF = False Then
'       vs1.TextMatrix(vs1.RowSel, 1) = rs.Fields("Unit").Value
'       SendKeys "{right}"
'       SendKeys "{right}"
'       Vs1Frame.Visible = False
'       vs1.Editable = flexEDKbdMouse
'       vs1.SetFocus
'    Else
'       vs1.TextMatrix(vs1.RowSel, 1) = "Kg"
'       SendKeys "{right}"
'       SendKeys "{right}"
'       Vs1Frame.Visible = False
'       vs1.Editable = flexEDKbdMouse
'       vs1.SetFocus
'
'    End If
'
' End If
'
' If vs1.col = 2 Then
'
'    SendKeys "{home}"
'    SendKeys "{down}"
'
'    AddItemInGrid1
' End If
'
'
'
' 'Total1
' TotalFinal
'
'End If
'
'
'End Sub
'Sub AddSemifinish()
'   Dim J As Integer
'
'   J = 1
'
'   vs3.Clear
'   For i = 1 To vs1.Rows - 1
'
'   If vs1.TextMatrix(i, 0) <> "" Then
'      If rs.State = 1 Then rs.Close
'      rs.Open "select * from ItemMaster where ItemName='" & vs1.TextMatrix(i, 0) & "'", CON
'      If rs.Fields("itemgp").Value = "Semi Finish (R/D)" Or rs.Fields("itemgp").Value = "Semi Finish (Store)" Then
'         vs3.TextMatrix(J, 0) = vs1.TextMatrix(i, 0)
'         vs3.TextMatrix(J, 1) = vs1.TextMatrix(i, 1)
'         vs3.TextMatrix(J, 2) = vs1.TextMatrix(i, 2)
'         J = J + 1
'      End If
'   End If
'
'   Next
'
'
'End Sub
'Private Sub vs1_LeaveCell()
'   'Total1
'End Sub
'
'Private Sub vs2_AfterEdit(ByVal row As Long, ByVal col As Long)
'     If vs2.col = 0 Then
'        'cellposiVs3
'        vs2.TextMatrix(vs2.RowSel, 0) = cboItemVs3.Text
'     End If
'
'End Sub
'
'Private Sub vs2_KeyDown(KeyCode As Integer, Shift As Integer)
' If KeyCode = 46 Then
'    vs2.RemoveItem (vs2.RowSel)
'    'Total2
'    TotalFinal
' End If
'
'
'  If KeyCode = 13 Then
'
'     If vs2.col = 0 Then
'        vs2.Editable = flexEDNone
'        Vs3Frame.Visible = True
'        cboItemVs3.Visible = True
'        cboItemVs3.SetFocus
'     Else
'        vs2.Editable = flexEDKbdMouse
'        'cellposiVs3
'     End If
'
'  End If
'
'End Sub
'
'Private Sub vs2_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'
'
' If vs2.col = 0 Then
'
'      vs2.Editable = flexEDNone
'      Vs3Frame.Visible = True
'
'
'
'    Set rs = New ADODB.Recordset
'    If rs.State = 1 Then rs.Close
'    rs.Open "select * from ItemMaster where ItemName='" & vs2.TextMatrix(vs2.RowSel, 0) & "'", CON
'    If rs.EOF = False Then
'       vs2.TextMatrix(vs2.RowSel, 1) = rs.Fields("Unit").Value
'       SendKeys "{right}"
'       SendKeys "{right}"
'       Vs3Frame.Visible = False
'       vs2.Editable = flexEDKbdMouse
'       vs2.SetFocus
'    Else
'       vs2.TextMatrix(vs2.RowSel, 1) = "Kg"
'       SendKeys "{right}"
'       SendKeys "{right}"
'       Vs3Frame.Visible = False
'       vs2.Editable = flexEDKbdMouse
'       vs2.SetFocus
'
'    End If
'
' End If
'
'
'    If vs2.col = 2 Then
'
'           SendKeys "{home}"
'           SendKeys "{down}"
'           Vs3Frame.TOP = Vs3Frame.TOP + 170
'    End If
'
'
'   'Total2
'
'End If
'
'End Sub
'Private Sub vs2_LeaveCell()
'   'Total2
'End Sub
'Private Sub vs3_AfterEdit(ByVal row As Long, ByVal col As Long)
'     If vs3.col = 0 Then
'        'cellposiVs2
'        'vs3.TextMatrix(vs3.RowSel, 0) = cboitemvscboItemVs2.Text
'     End If
'
'End Sub
'
'Private Sub vs3_KeyDown(KeyCode As Integer, Shift As Integer)
'
'  If KeyCode = 46 Then
'    vs3.RemoveItem (vs3.RowSel)
'    'Total4
'  End If
'
'  If KeyCode = 13 Then
'     If vs3.col = 0 Then
'
'        vs3.Editable = flexEDNone
'        FrameVs2.Visible = True
'        cboItemVs2.Visible = True
'        cboItemVs2.SetFocus
'     Else
'
'        vs3.Editable = flexEDKbdMouse
'
'     End If
'  End If
'
'End Sub
'
'Private Sub vs3_KeyUp(KeyCode As Integer, Shift As Integer)
'
'If KeyCode = 13 Then
'
' If vs3.col = 0 Then
'    vs3.Editable = flexEDNone
'    FrameVs2.Visible = True
'    cboItemVs2.SetFocus
'
'
'
'
'    Set rs = New ADODB.Recordset
'    If rs.State = 1 Then rs.Close
'    rs.Open "select * from ItemMaster where ItemName='" & vs3.TextMatrix(vs3.RowSel, 0) & "'", CON
'    If rs.EOF = False Then
'       vs3.TextMatrix(vs3.RowSel, 1) = rs.Fields("Unit").Value
'       SendKeys "{right}"
'       SendKeys "{right}"
'       FrameVs2.Visible = False
'       vs3.Editable = flexEDKbdMouse
'       vs3.SetFocus
'    Else
'       vs3.TextMatrix(vs3.RowSel, 1) = "Kg"
'       SendKeys "{right}"
'       SendKeys "{right}"
'       FrameVs2.Visible = False
'       vs3.Editable = flexEDKbdMouse
'       vs3.SetFocus
'
'    End If
'
' End If
'
' If vs3.col = 2 Then
'
'   If rs.State = 1 Then rs.Close
'   rs.Open "select  OpeningStock from ItemMaster where ItemName='" & vs3.TextMatrix(vs.RowSel, 0) & "'", CON
'   If rs.EOF = False Then
'      If Val(rs.Fields(0).Value) < Val(vs3.TextMatrix(vs3.RowSel, 2)) Then
'         MsgBox "Stock Less !!", vbInformation
'
'      End If
'   End If
'
'
'    SendKeys "{home}"
'    SendKeys "{down}"
'
'    FrameVs2.TOP = FrameVs2.TOP + 170
'    'AddItemInGrid2
' End If
'
'
'
' 'Total4
'
'End If
'
'End Sub
'
'
'
'
