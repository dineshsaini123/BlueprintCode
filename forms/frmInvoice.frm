VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmInvoice 
   Caption         =   "Sales Invoice"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10410
   ScaleWidth      =   12825
   WindowState     =   2  'Maximized
   Begin VB.Frame buttonFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   1980
      TabIndex        =   22
      Top             =   8400
      Width           =   8385
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFC0&
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
         Left            =   4680
         Picture         =   "frmInvoice.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   135
         Width           =   1140
      End
      Begin VB.CommandButton cmdExit_12 
         BackColor       =   &H00FFFFC0&
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
         Left            =   7065
         Picture         =   "frmInvoice.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdEdit_4 
         BackColor       =   &H00FFFFC0&
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
         Left            =   3525
         Picture         =   "frmInvoice.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   135
         Width           =   1140
      End
      Begin VB.CommandButton cmdDelete_3 
         BackColor       =   &H00FFFFC0&
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
         Left            =   2340
         Picture         =   "frmInvoice.frx":1BD5
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   1185
      End
      Begin VB.CommandButton cmdSave_2 
         BackColor       =   &H00FFFFC0&
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
         Left            =   1200
         Picture         =   "frmInvoice.frx":27B9
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   135
         Width           =   1140
      End
      Begin VB.CommandButton cmdAdd_1 
         BackColor       =   &H00FFFFC0&
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
         Picture         =   "frmInvoice.frx":339D
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   135
         Width           =   1140
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Print"
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
         Left            =   5850
         Picture         =   "frmInvoice.frx":3F81
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.ComboBox cboItem 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   1320
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   21
      Top             =   4980
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1860
      TabIndex        =   16
      Top             =   3240
      Width           =   3675
   End
   Begin VB.ComboBox cboGodown 
      Height          =   315
      Left            =   1860
      TabIndex        =   14
      Top             =   2820
      Width           =   3675
   End
   Begin VB.TextBox Text3 
      Height          =   765
      Left            =   1860
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   1980
      Width           =   4155
   End
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   1860
      TabIndex        =   8
      Top             =   1620
      Width           =   4155
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   1800
      TabIndex        =   7
      Top             =   660
      Width           =   1155
   End
   Begin VB.TextBox txtOrderNo 
      Height          =   345
      Left            =   4860
      TabIndex        =   3
      Top             =   600
      Width           =   1155
   End
   Begin MSMask.MaskEdBox txtOrderDate 
      Height          =   315
      Left            =   4860
      TabIndex        =   1
      Top             =   1080
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Top             =   1140
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   315
      Left            =   2220
      TabIndex        =   17
      Top             =   3720
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   315
      Left            =   2220
      TabIndex        =   19
      Top             =   4080
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   3270
      Left            =   420
      TabIndex        =   30
      Top             =   4500
      Width           =   10005
      _cx             =   17648
      _cy             =   5768
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16744576
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
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   325
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
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount :"
      Height          =   240
      Left            =   7560
      TabIndex        =   34
      Top             =   7980
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Qty :"
      Height          =   240
      Left            =   5670
      TabIndex        =   33
      Top             =   7980
      Width           =   825
   End
   Begin VB.Label lblQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6390
      TabIndex        =   32
      Top             =   7980
      Width           =   1050
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8685
      TabIndex        =   31
      Top             =   7995
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "Date of Dispatch :"
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
      Index           =   9
      Left            =   480
      TabIndex        =   20
      Top             =   4080
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "Date of Issue :"
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
      Index           =   8
      Left            =   480
      TabIndex        =   18
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Agent Name :"
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
      Index           =   7
      Left            =   480
      TabIndex        =   15
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Godown :"
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
      Index           =   6
      Left            =   480
      TabIndex        =   13
      Top             =   2820
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Address :"
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
      Index           =   5
      Left            =   480
      TabIndex        =   12
      Top             =   1980
      Width           =   1275
   End
   Begin VB.Label header 
      BackColor       =   &H8000000D&
      Caption         =   "    Sales Invoice"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   10755
   End
   Begin VB.Label Label1 
      Caption         =   "Party Name :"
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
      Index           =   4
      Left            =   480
      TabIndex        =   9
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Invoice No."
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
      Index           =   3
      Left            =   480
      TabIndex        =   6
      Top             =   660
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Invoice Date"
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
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   1140
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Order No."
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
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Order Date"
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
      Index           =   0
      Left            =   3720
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_12_Click()
Unload Me
End Sub

Private Sub Form_Load()
formatVSGrid
header(0).TOP = MainMenu.TOP + 60
header(0).Left = MainMenu.Left
header(0).Width = MainMenu.Width

End Sub
Sub formatVSGrid()

    vs.Clear

    vs.Rows = 2
    vs.Cols = 8

    vs.FormatString = "SNo.|Item Code|Item Name|Quantity|M.R.P.|Net Rate|Net Amount"

    For K = 0 To 6
        vs.Cell(flexcpFontBold, 0, K) = True
    Next

    For K = 0 To 6
        vs.Cell(flexcpForeColor, 0, K) = vbWhite
    Next

    vs.ColWidth(0) = 800
    vs.ColWidth(1) = 1200
    vs.ColWidth(2) = 3000
    vs.ColWidth(3) = 1050
    vs.ColWidth(4) = 1200
    vs.ColWidth(5) = 1200
    vs.ColWidth(6) = 1500
    vs.ColWidth(7) = 0


End Sub
