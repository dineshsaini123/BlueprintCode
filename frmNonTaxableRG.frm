VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmNonTaxableRG 
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19140
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   19140
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdShow 
      Caption         =   "&View"
      Height          =   495
      Left            =   3900
      TabIndex        =   5
      Top             =   780
      Width           =   1215
   End
   Begin VB.CommandButton Command1_print 
      Caption         =   "&Print"
      Height          =   495
      Left            =   5220
      TabIndex        =   2
      Top             =   780
      Width           =   1215
   End
   Begin VB.CommandButton Command2_exit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   6540
      TabIndex        =   1
      Top             =   780
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dt2 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   780
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72089601
      CurrentDate     =   41274
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   780
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72089601
      CurrentDate     =   41274
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   7455
      Left            =   60
      TabIndex        =   6
      Top             =   1320
      Width           =   18135
      _cx             =   31988
      _cy             =   13150
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorFixed  =   16710321
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   2
      FixedCols       =   0
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   780
      Width           =   555
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Non Taxable RG -1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   18015
   End
End
Attribute VB_Name = "frmNonTaxableRG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub vsINI()
   
   
 vs.Cols = 20
 
 vs.MergeCells = flexMergeFixedOnly
 vs.WordWrap = True
 
  With vs
   
   .RowHeight(1) = 800
   
   vs.MergeRow(0) = True
   vs.MergeCol(0) = True
   vs.MergeCol(1) = True
   vs.MergeCol(2) = True
   
   .MergeCol(9) = True
   .MergeCol(10) = True
   .MergeCol(11) = True
   .MergeCol(12) = True
   
   .MergeCol(13) = True
   .MergeCol(14) = True
   .MergeCol(15) = True
   
    .MergeCol(17) = True
   .MergeCol(18) = True
   .MergeCol(19) = True
   
  
   
   
  
    
   For K = 3 To 8
      vs.ColAlignment(K) = flexAlignCenterCenter
   Next
  
  
     For K = 13 To 16
      vs.ColAlignment(K) = flexAlignCenterCenter
   Next

   
   .MergeCells = flexMergeFree
   .MergeRow(0) = True
   
   .MergeCells = flexMergeFree
   .MergeCol(0) = True
   .MergeCol(1) = True
   .MergeCol(2) = True
   .MergeCol(3) = True
   
   
   vs.TextMatrix(0, 0) = "Date"
   vs.TextMatrix(0, 1) = "Opening Balance"
   vs.TextMatrix(0, 2) = "Quantity Manufactured"
   vs.TextMatrix(0, 3) = "Total (2+3)"
   
   vs.TextMatrix(0, 4) = "For Home Use (Domestic Sales)"
   vs.TextMatrix(0, 5) = "For Home Use (Domestic Sales)"
   
   vs.TextMatrix(0, 6) = "For Export Under Claim for Rebate of Duty"
   vs.TextMatrix(0, 7) = "For Export Under Claim for Rebate of Duty"
   vs.TextMatrix(0, 8) = "For Export Under Claim for Rebate of Duty"
   
   .TextMatrix(0, 9) = "For Other Factories or Warehouse Under Bond"
   
   
   .TextMatrix(0, 10) = "For other Purpose"
   .TextMatrix(0, 11) = "For other Purpose"
   .TextMatrix(0, 12) = "For other Purpose"


   .TextMatrix(0, 13) = "Duty Payable & Paid"
   .TextMatrix(0, 14) = "Duty Payable & Paid"
   .TextMatrix(0, 15) = "Duty Payable & Paid"

  
   .TextMatrix(0, 16) = "Total Duty"
  
   .TextMatrix(0, 17) = "closing balance"
   .TextMatrix(0, 18) = "closing balance"
   
   .TextMatrix(0, 19) = "Remark(Bill No.)"
   
   
   
   
   
   vs.TextMatrix(1, 0) = "Date"
   vs.TextMatrix(1, 1) = "Opening Balance"
   vs.TextMatrix(1, 2) = "Quantity Manufactured"
   vs.TextMatrix(1, 3) = "Total (2+3)"
   
   vs.TextMatrix(1, 4) = "Quantity"
   vs.TextMatrix(1, 5) = "Value"
   
   .TextMatrix(1, 6) = "Quantity"
   .TextMatrix(1, 7) = "Value"
   .TextMatrix(1, 8) = "For Export Under Bond"
   
   
   .TextMatrix(1, 9) = "For Other Factories or Warehouse Under Bond"
   
   
   .TextMatrix(1, 10) = "Purpose"
   .TextMatrix(1, 11) = "Quantity"
   .TextMatrix(1, 12) = "Rate"
   
   
   
   .TextMatrix(1, 13) = "BED @2%"
   .TextMatrix(1, 14) = "Edu. Cess @0.02%"
   .TextMatrix(1, 15) = "S.H.Edu. Cess @0.01%"
   
   .TextMatrix(1, 16) = "Total Duty"
   
   
   .TextMatrix(1, 17) = "In Finishing Room"
   .TextMatrix(1, 18) = "In Bonded Store Room"
   
   .TextMatrix(1, 19) = "Remark(Bill No.)"

   
  
 
 End With
  
  
 For i = 1 To 19
    vs.ColWidth(i) = 1500
 Next



dt1.Value = Date
dt2.Value = Date
 
   
   
End Sub

Private Sub Command2_exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
vsINI
End Sub


