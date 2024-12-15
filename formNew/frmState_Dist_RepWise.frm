VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmState_Dist_RepWise 
   Caption         =   "Sale Summry"
   ClientHeight    =   9672
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   15432
   Icon            =   "frmState_Dist_RepWise.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9672
   ScaleWidth      =   15432
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1_saleret 
      Caption         =   "Sale Ret. As On"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8880
      TabIndex        =   45
      Top             =   270
      Width           =   1740
   End
   Begin VB.TextBox txtGP 
      Height          =   315
      Left            =   4500
      TabIndex        =   43
      Top             =   2280
      Width           =   930
   End
   Begin VB.TextBox txtbk 
      Height          =   315
      Left            =   3015
      TabIndex        =   41
      Top             =   2280
      Width           =   1470
   End
   Begin VB.TextBox txtbkName 
      Height          =   315
      Left            =   60
      TabIndex        =   40
      Top             =   2280
      Width           =   2955
   End
   Begin VB.Frame frmSaleSummary 
      Height          =   2220
      Left            =   5460
      TabIndex        =   37
      Top             =   630
      Width           =   5775
      Begin VSFlex7Ctl.VSFlexGrid vs1 
         Height          =   2010
         Left            =   45
         TabIndex        =   38
         Top             =   135
         Width           =   5700
         _cx             =   10054
         _cy             =   3545
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
         BackColorFixed  =   12582847
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
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   2
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmState_Dist_RepWise.frx":000C
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
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   225
            Left            =   14715
            MultiLine       =   -1  'True
            TabIndex        =   39
            Top             =   4020
            Width           =   195
         End
      End
   End
   Begin VB.Frame frmbilty 
      Height          =   1725
      Left            =   45
      TabIndex        =   27
      Top             =   630
      Visible         =   0   'False
      Width           =   5190
      Begin VB.TextBox txtRep 
         Height          =   285
         Left            =   810
         TabIndex        =   31
         Top             =   1170
         Width           =   4290
      End
      Begin VB.CheckBox Check1_crnote 
         Caption         =   "Bilty Return (Without Credit Note)"
         Height          =   240
         Left            =   810
         TabIndex        =   32
         Top             =   1440
         Width           =   3165
      End
      Begin VB.TextBox txtBilty_State 
         Height          =   285
         Left            =   810
         TabIndex        =   30
         Top             =   855
         Width           =   4290
      End
      Begin VB.TextBox txtBilty_Dist 
         Height          =   285
         Left            =   810
         TabIndex        =   29
         Top             =   540
         Width           =   4290
      End
      Begin VB.TextBox txtBilty_Party 
         Height          =   285
         Left            =   810
         TabIndex        =   28
         Top             =   225
         Width           =   4290
      End
      Begin VB.Label Label7 
         Caption         =   "Rep.:"
         Height          =   285
         Left            =   135
         TabIndex        =   36
         Top             =   1170
         Width           =   690
      End
      Begin VB.Label Label6 
         Caption         =   "State:"
         Height          =   285
         Left            =   135
         TabIndex        =   35
         Top             =   855
         Width           =   690
      End
      Begin VB.Label Label5 
         Caption         =   "District :"
         Height          =   285
         Left            =   135
         TabIndex        =   34
         Top             =   585
         Width           =   690
      End
      Begin VB.Label Label4 
         Caption         =   "Party :"
         Height          =   240
         Left            =   135
         TabIndex        =   33
         Top             =   225
         Width           =   645
      End
   End
   Begin VB.TextBox txtTot1_2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7590
      TabIndex        =   26
      Top             =   9045
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtTot1_1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6270
      TabIndex        =   25
      Top             =   9045
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtTot1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8910
      TabIndex        =   23
      Top             =   9045
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtTot3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   11655
      TabIndex        =   22
      Top             =   9045
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txtTot2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10290
      TabIndex        =   21
      Top             =   9045
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSComCtl2.DTPicker txtFrom 
      Height          =   312
      Left            =   5544
      TabIndex        =   19
      Top             =   240
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   550
      _Version        =   393216
      Format          =   173080577
      CurrentDate     =   42409
   End
   Begin VB.CheckBox Check1_godwon 
      Caption         =   "Godown"
      Height          =   315
      Left            =   60
      TabIndex        =   18
      Top             =   1620
      Width           =   1515
   End
   Begin VB.TextBox txtParty1 
      Height          =   315
      Left            =   60
      TabIndex        =   17
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtParty 
      Height          =   315
      Left            =   45
      TabIndex        =   15
      Top             =   1170
      Width           =   3900
   End
   Begin VB.CommandButton cmdRepQty 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Export to Excel"
      Height          =   555
      Left            =   12615
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   825
      Width           =   2400
   End
   Begin VB.ListBox cmbAgentName 
      Appearance      =   0  'Flat
      Height          =   1752
      Left            =   5484
      Style           =   1  'Checkbox
      TabIndex        =   10
      Top             =   855
      Width           =   3315
   End
   Begin VB.CheckBox Check1_selectAll_Rep 
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   9
      Top             =   900
      Width           =   1335
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   13380
      TabIndex        =   7
      Top             =   9045
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.ComboBox cboType 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   330
      ItemData        =   "frmState_Dist_RepWise.frx":0082
      Left            =   60
      List            =   "frmState_Dist_RepWise.frx":00D1
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   5160
   End
   Begin VB.CommandButton cmdExit_12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      Height          =   615
      Left            =   12615
      Picture         =   "frmState_Dist_RepWise.frx":0417
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1425
      Width           =   2430
   End
   Begin VB.CommandButton cmdPrint_7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print"
      Enabled         =   0   'False
      Height          =   615
      Left            =   13836
      Picture         =   "frmState_Dist_RepWise.frx":0FFB
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   165
      Width           =   1176
   End
   Begin VB.CommandButton cmdAdd_1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
      Height          =   615
      Left            =   12612
      Picture         =   "frmState_Dist_RepWise.frx":1BDF
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   165
      Width           =   1176
   End
   Begin VSFlex7Ctl.VSFlexGrid vs 
      Height          =   6192
      Left            =   120
      TabIndex        =   0
      Top             =   2856
      Width           =   15120
      _cx             =   26670
      _cy             =   10922
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
      BackColorFixed  =   12582847
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1000
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   450
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmState_Dist_RepWise.frx":27C3
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
      Editable        =   1
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
      Begin VB.TextBox txtNote 
         Appearance      =   0  'Flat
         Height          =   225
         Left            =   15075
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   4020
         Width           =   195
      End
   End
   Begin MSComCtl2.DTPicker dateAsOn 
      Height          =   312
      Left            =   7164
      TabIndex        =   13
      Top             =   240
      Width           =   1416
      _ExtentX        =   2498
      _ExtentY        =   550
      _Version        =   393216
      Format          =   173080577
      CurrentDate     =   42409
   End
   Begin MSComCtl2.DTPicker txtSaleRetDate 
      Height          =   312
      Left            =   10632
      TabIndex        =   46
      Top             =   276
      Visible         =   0   'False
      Width           =   1308
      _ExtentX        =   2307
      _ExtentY        =   550
      _Version        =   393216
      Format          =   173080577
      CurrentDate     =   42409
   End
   Begin VB.Label lblGP 
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
      Height          =   315
      Left            =   4500
      TabIndex        =   44
      Top             =   2025
      Width           =   855
   End
   Begin VB.Label lblbk 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Name"
      Height          =   315
      Left            =   3105
      TabIndex        =   42
      Top             =   2025
      Width           =   1215
   End
   Begin VB.Label lblBookName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   315
      Left            =   60
      TabIndex        =   20
      Top             =   1980
      Width           =   1755
   End
   Begin VB.Label lblParty 
      BackStyle       =   0  'Transparent
      Caption         =   "Party Name"
      Height          =   315
      Left            =   60
      TabIndex        =   16
      Top             =   960
      Width           =   1755
   End
   Begin VB.Label lblAson 
      Caption         =   "To"
      Height          =   252
      Left            =   6900
      TabIndex        =   14
      Top             =   240
      Width           =   276
   End
   Begin VB.Label lblRep 
      BackStyle       =   0  'Transparent
      Caption         =   "Representative :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   5544
      TabIndex        =   11
      Top             =   660
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      Height          =   255
      Left            =   765
      TabIndex        =   8
      Top             =   9045
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Type Report :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   1875
   End
   Begin VB.Label Label3 
      BackColor       =   &H00BFFFBF&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   9090
      Width           =   15015
   End
End
Attribute VB_Name = "frmState_Dist_RepWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rtype As String
Private Sub cboType_Click()

lblbk.Visible = False
txtbk.Visible = False

lblRep.Visible = False
cmbAgentName.Visible = False
Check1_selectAll_Rep.Visible = False
cmdRepQty.Visible = False


txtFrom.Visible = True
lblAson.Visible = True
dateAson.Visible = True
frmSaleSummary.Visible = False


lblParty.Visible = False
txtParty.Visible = False
txtParty1.Visible = False
Check1_godwon.Visible = False


lblBookName.Visible = False
txtbkName.Visible = False

frmbilty.Visible = False
cmdRepQty.Visible = False

lblGP.Visible = False
txtgp.Visible = False

If cboType.text = "Bilty Return Status..." Then
    frmbilty.Visible = True
    rtype = "Bilty Return Status..."
    lblBookName.Visible = True
    txtbkName.Visible = True
    
ElseIf (cboType.text = "Consolidated Sales Summary...") Then

 txtFrom.Visible = False
 lblAson.Visible = False
 dateAson.Visible = False
 frmSaleSummary.Visible = True
 cmdRepQty.Visible = True
 rtype = "Consolidated Sales Summary..."
 
 lblBookName.Visible = True
 txtbkName.Visible = True

 
vs1.rows = 6
 
 If RS.State = 1 Then RS.close
 RS.Open "select id,head,[from],to from salesummary order by id", CCON
 For I = 0 To 5
     vs1.TextMatrix(I, 0) = RS!id
     vs1.TextMatrix(I, 1) = RS!head
     vs1.TextMatrix(I, 2) = RS!from
     If RS!to = "Till Date" Then
     vs1.TextMatrix(I, 3) = Format(Date, "dd/MM/yyyy")
     Else
     vs1.TextMatrix(I, 3) = RS!to
     End If
     RS.MoveNext
 Next
 
 vs1.ColWidth(0) = 700
 vs1.ColWidth(1) = 1850
 vs1.ColWidth(2) = 1500
 vs1.ColWidth(3) = 1500

ElseIf (cboType.text = "Consolidated Sales Summary Rep. Wise ...") Then

 txtFrom.Visible = False
 lblAson.Visible = False
 dateAson.Visible = False
 frmSaleSummary.Visible = True
 cmdRepQty.Visible = True
 rtype = "Consolidated Sales Summary Rep. Wise ..."
 lblBookName.Visible = True
 txtbkName.Visible = True

 
vs1.rows = 6
 
 If RS.State = 1 Then RS.close
 RS.Open "select id,head,[from],to from salesummary order by id", CCON
 For I = 0 To 5
     vs1.TextMatrix(I, 0) = RS!id
     vs1.TextMatrix(I, 1) = RS!head
     vs1.TextMatrix(I, 2) = RS!from
     If RS!to = "Till Date" Then
     vs1.TextMatrix(I, 3) = Format(Date, "dd/MM/yyyy")
     Else
     vs1.TextMatrix(I, 3) = RS!to
     End If
     RS.MoveNext
 Next
 
 vs1.ColWidth(0) = 700
 vs1.ColWidth(1) = 1850
 vs1.ColWidth(2) = 1500
 vs1.ColWidth(3) = 1500
    
ElseIf (cboType.text = "Representative & Book Wise Net Sales(Amt.)") Then

    lblRep.Visible = True
    cmbAgentName.Visible = True
    Check1_selectAll_Rep.Visible = True
    cmdRepQty.Visible = True
    cmdPrint_7.Enabled = True
    rtype = "bookwise"
    lblAson.Visible = True
    dateAson.Visible = True
    
    lblBookName.Visible = True
    txtbkName.Visible = True
    
ElseIf cboType = "Representative & Book Wise Sales Return" Then
   lblRep.Visible = True
    cmbAgentName.Visible = True
    Check1_selectAll_Rep.Visible = True
    cmdRepQty.Visible = True
    cmdPrint_7.Enabled = True
    rtype = "bookwiseret"
    lblAson.Visible = True
    dateAson.Visible = True
    
    lblBookName.Visible = True
    txtbkName.Visible = True
    
ElseIf cboType = "Rep. Wise & Book & Bill Wise Sale..." Then

   lblbk.Visible = True
   txtbk.Visible = True


    lblParty.Visible = True
    txtParty.Visible = True
    txtParty1.Visible = False
    lblParty.Caption = "Rep. Name"
    rtype = "Rep. Wise & Book & Bill Wise Sale..."
    lblbk.Caption = "State "
    lblBookName.Visible = True
    txtbkName.Visible = True
ElseIf cboType = "Tital Wise & Party Wise Sale & Sale Ret. Qty" Then

   lblbk.Visible = False
   txtbk.Visible = False

    cmdRepQty.Visible = True
    lblParty.Visible = False
    txtParty.Visible = False
    txtParty1.Visible = False
    lblParty.Caption = "Rep. Name"
    
    rtype = "Tital Wise & Party Wise Sale & Sale Ret. Qty"
    lblbk.Caption = "State "
    lblBookName.Visible = True
    txtbkName.Visible = True


ElseIf cboType = "Book Wise & School Wise Net Sale.." Then
   lblParty.Visible = True
    txtParty.Visible = True
    txtParty1.Visible = False
    lblParty.Caption = "State"
    rtype = "Book Wise & School Wise Net Sale.."
    
    lblbk.Visible = True
    txtbk.Visible = True

    lblGP.Visible = True
    txtgp.Visible = True

    
    lblBookName.Visible = True
    txtbkName.Visible = True

ElseIf (cboType.text = "Book Wise Sales") Then
    cmdPrint_7.Enabled = True
    rtype = "Book Wise Sales"
    
    lblBookName.Visible = True
    txtbkName.Visible = True
    
    
ElseIf (cboType.text = "State Wise") Then
    cmdPrint_7.Enabled = True
    lblBookName.Visible = True
    txtbkName.Visible = True

ElseIf cboType.text = "Party Wise & Book Wise Gross Sales" Then
    
    
    lblParty.Visible = True
    txtParty.Visible = True
    txtParty1.Visible = False
    lblParty.Caption = "Party"
    cmdPrint_7.Enabled = True
    Check1_godwon.Visible = True
    
    txtbkName.Visible = True
    lblBookName.Visible = True
    
    rtype = "grosssale"

ElseIf cboType.text = "Party Wise & Book Wise Sale & Return" Then
   
    lblParty.Visible = True
    txtParty.Visible = True
    txtParty1.Visible = False
    lblParty.Caption = "Party"
    cmdPrint_7.Enabled = True
    Check1_godwon.Visible = True
    
    txtbkName.Visible = True
    lblBookName.Visible = True

   
ElseIf cboType.text = "Party Payment Details" Then
    lblParty.Visible = True
    txtParty.Visible = True
    txtParty1.Visible = False
    lblParty.Caption = "State"
    rtype = "Party Payment Details"
    
    lblBookName.Visible = False
    txtbkName.Visible = False
    
ElseIf cboType.text = "School Wise & Book Wise Net Sale.." Then

    lblParty.Visible = True
    txtParty.Visible = True
    txtParty1.Visible = False
    lblParty.Caption = "State"
    rtype = "SchoolWiseSale"
    
    lblBookName.Visible = True
    txtbkName.Visible = True
    
    lblGP.Visible = True
    txtgp.Visible = True

    
ElseIf cboType.text = "State Wise & Book Wise Gross Sales.." Then
    
    lblParty.Visible = True
    txtParty.Visible = True
    txtParty1.Visible = False
    lblParty.Caption = "State"
    rtype = "grosssale_Statewise"
    
    lblBookName.Visible = True
    txtbkName.Visible = True
    
  
ElseIf cboType.text = "Party Credit Register..." Then
      lblParty.Visible = True
    txtParty.Visible = True
    txtParty1.Visible = False
    lblParty.Caption = "State"
    rtype = "CreditRegister"
    
ElseIf cboType.text = "Party Wise & Rep.wise Net Sale.." Then
    lblRep.Visible = True
    cmbAgentName.Visible = True
    Check1_selectAll_Rep.Visible = True
    cmdRepQty.Visible = False
    cmdPrint_7.Enabled = True
    rtype = "Party Wise & Rep.wise"
    lblAson.Visible = True
    dateAson.Visible = True
    lblBookName.Visible = True
    txtbkName.Visible = True
    
ElseIf cboType.text = "Representative & Book Wise Net Sales(Qty.)" Then
    lblRep.Visible = True
    cmbAgentName.Visible = True
    Check1_selectAll_Rep.Visible = True
    cmdRepQty.Visible = True
    cmdPrint_7.Enabled = True
    rtype = "bookwiseQty"
    lblAson.Visible = True
    dateAson.Visible = True
    
    lblBookName.Visible = True
    txtbkName.Visible = True


ElseIf cboType = "Representative & Bill Wise Net Sales" Then
    rtype = "billwise"
    cmdPrint_7.Enabled = True
    lblBookName.Visible = True
    txtbkName.Visible = True
ElseIf cboType = "Party Wise Area Wise Net Quantity Sale..." Then
       rtype = "Party Wise Area Wise Net Quantity Sale..."
       cmdPrint_7.Enabled = True
           lblBookName.Visible = True
    txtbkName.Visible = True

ElseIf cboType = "Stock In Hand As On" Then
    rtype = "Stock In Hand As On"
    cmdRepQty.Visible = True
    
    dateAson.Visible = True
    lblAson.Visible = True
ElseIf cboType = "Rep Wise" Then
    rtype = "Rep Wise"
    dateAson.Visible = True
    lblAson.Visible = True
     lblBookName.Visible = True
    txtbkName.Visible = True

ElseIf cboType = "District Wise" Then
    rtype = "District Wise"
    dateAson.Visible = True
    lblAson.Visible = True
   txtbkName.Visible = True
   lblBookName.Visible = True

ElseIf cboType = "Party Wise Area Wise Net Sale..." Then
   rtype = "PartyWiseArea"
   lblBookName.Visible = True
   txtbkName.Visible = True
ElseIf cboType = "Party Wise Area & Rep. Wise Net Sale..." Then
   rtype = "Party Wise Area & Rep. Wise Net Sale..."
   lblBookName.Visible = True
   txtbkName.Visible = True
ElseIf cboType = "Book Wise Sales(Area Wise)" Then
   rtype = "Book Wise Sales(Area Wise)"
   
   txtbkName.Visible = True
   lblBookName.Visible = True

ElseIf cboType = "Rep.Wise & Title Wise Net Qty Summary.." Then
    
    lblbk.Visible = False
    txtbk.Visible = False

    cmdRepQty.Visible = False
    lblParty.Visible = False
    txtParty.Visible = False
    txtParty1.Visible = False
    lblParty.Caption = "Rep. Name"
    
    rtype = "Rep.Wise & Title Wise Net Qty Summary.."
    lblbk.Caption = "State "
    lblBookName.Visible = True
    txtbkName.Visible = True
Else
   rtype = "all"
End If


End Sub

Private Sub Check1_book_Click()
For J = 0 To List1_books.ListCount - 1
    List1_books.Selected(J) = False
Next
If Check1_book.value = 1 Then
  For J = 0 To List1_books.ListCount - 1
      List1_books.Selected(J) = True
  Next
End If
End Sub

Private Sub Check1_godwon_Click()
If Check1_godwon.value = 1 Then
lblParty.Caption = "Godown"
Else
lblParty.Caption = "Party"
End If
End Sub

Private Sub Check1_saleret_Click()
If Check1_saleret.value = 1 Then
   txtSaleRetDate.Visible = True
Else
   txtSaleRetDate.Visible = False
End If
End Sub

Private Sub Check1_selectAll_Rep_Click()
For J = 0 To cmbAgentName.ListCount - 1
    cmbAgentName.Selected(J) = False
Next
If Check1_selectAll_Rep.value = 1 Then
  For J = 0 To cmbAgentName.ListCount - 1
      cmbAgentName.Selected(J) = True
  Next
End If
End Sub
Sub fillStockGrid()
Dim rs_data As New ADODB.Recordset
Dim rs_issue As New ADODB.Recordset
Dim rs_sale As New ADODB.Recordset
Dim rs_sale1 As New ADODB.Recordset

Dim I As Integer

Dim godown_rec As Long
Dim godown_issue As Long

I = 1

If RS.State = 1 Then RS.close
RS.Open "select sum(Quantity),BOOKCODE from stocksummaryQry where convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" Godown='M' and issue_Receive='Receive' group by BOOKCODE", con, adOpenKeyset

If rs_issue.State = 1 Then rs_issue.close
rs_issue.Open "select sum(Quantity),BOOKCODE from stocksummaryQry where convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" Godown='M' and issue_Receive='Issue' group by BOOKCODE", con, adOpenKeyset

If rs_sale.State = 1 Then rs_sale.close
rs_sale.Open "SELECT BOOKCODE,sum([QUANTITY]) FROM INVOICEB where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) group by BOOKCODE", con, adOpenKeyset, adLockReadOnly

If rs_sale1.State = 1 Then rs_sale1.close
rs_sale1.Open "SELECT BOOKCODE,sum([QUANTITY]) FROM INVOICEB_free where convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) group by BOOKCODE", con, adOpenKeyset, adLockReadOnly





If rs_data.State = 1 Then rs_data.close
rs_data.Open "select BOOKCODE,bookname from  Books where kitcode='n' and " & stringyear & "  group by BOOKCODE,bookname", con, adOpenKeyset
If rs_data.EOF = False Then
vs.rows = rs_data.RecordCount + 2
End If

While rs_data.EOF = False
   
   
  vs.TextMatrix(I, 0) = rs_data!Bookcode
  vs.TextMatrix(I, 1) = rs_data!Bookname
    
  godown_rec = 0
  godown_issue = 0

 
 RS.MoveFirst
 RS.Find "BOOKCODE='" & rs_data.Fields("BOOKCODE").value & "'"
 If RS.EOF = False Then
 godown_rec = RS(0)
 End If
 
 rs_issue.MoveFirst
 rs_issue.Find "BOOKCODE='" & rs_data.Fields("BOOKCODE").value & "'"
 If rs_issue.EOF = False Then
 godown_issue = rs_issue(0)
 End If
 
   
 rs_sale.MoveFirst
 rs_sale.Find "BOOKCODE='" & rs_data.Fields("BOOKCODE").value & "'"
 If rs_sale.EOF = False Then
 vs.TextMatrix(I, 8) = rs_sale(1)
 End If
 
 If rs_sale1.EOF = False Then
 
 rs_sale1.MoveFirst
 rs_sale1.Find "BOOKCODE='" & rs_data.Fields("BOOKCODE").value & "'"
 If rs_sale1.EOF = False Then
 vs.TextMatrix(I, 8) = Val(vs.TextMatrix(I, 8)) + rs_sale1(1)
 End If
 
 End If
 
   
 vs.TextMatrix(I, 2) = (godown_rec - godown_issue)
  
 I = I + 1
 rs_data.MoveNext
 
Wend



'===================================================================================
'===================================================================================
con.Execute "exec BookStockSummary 'ALL','NS'"
'===================================================================================
'===================================================================================


If RS.State = 1 Then RS.close
RS.Open "select sum(Quantity),BOOKCODE from stocksummaryQry where convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" Godown='NS' and issue_Receive='Receive' group by BOOKCODE", con, adOpenKeyset

If rs_issue.State = 1 Then rs_issue.close
rs_issue.Open "select sum(Quantity),BOOKCODE from stocksummaryQry where convert(smalldatetime,Dates,103)<= convert(smalldatetime,'" & dateAson & "',103) and " & _
" Godown='NS' and issue_Receive='Issue' group by BOOKCODE", con, adOpenKeyset


For I = 1 To vs.rows - 1
  
If vs.TextMatrix(I, 0) <> "" Then

   godown_rec = 0
   godown_issue = 0
   
   
   If RS.EOF = False Then
   RS.MoveFirst
   RS.Find "BOOKCODE='" & vs.TextMatrix(I, 0) & "'"
   If RS.EOF = False Then
        godown_rec = RS(0)
   End If
   End If
 
   If rs_issue.EOF = False Then
   rs_issue.MoveFirst
   rs_issue.Find "BOOKCODE='" & vs.TextMatrix(I, 0) & "'"
   If rs_issue.EOF = False Then
     godown_issue = rs_issue(0)
   End If
   End If
   
   

   
   vs.TextMatrix(I, 3) = (godown_rec - godown_issue)
   
   vs.TextMatrix(I, 4) = 0
   vs.TextMatrix(I, 5) = ((Val(vs.TextMatrix(I, 2)) + Val(vs.TextMatrix(I, 3))) - vs.TextMatrix(I, 4))
  
   vs.TextMatrix(I, 6) = 0
   

  
   godown_rec = 0
   godown_issue = 0
   
   If rs_data.State = 1 Then rs_data.close
   rs_data.Open "SELECT INVOICENO,INVOICEDATE,BOOKCODE,sum([QUANTITY]) as Qty FROM ORDERB where BOOKCODE='" & vs.TextMatrix(I, 0) & "' and convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103)  group by INVOICENO,INVOICEDATE,BOOKCODE"
   While rs_data.EOF = False
    godown_rec = godown_rec + rs_data(3)
    rs_data.MoveNext
   Wend

   If rs_data.State = 1 Then rs_data.close
   rs_data.Open "SELECT OrderNo,INVOICEDATE,BOOKCODE,sum(QUANTITY) as qty1 FROM invoiceBQry where OrderNo>0 and BOOKCODE='" & vs.TextMatrix(I, 0) & "' and convert(smalldatetime,invoicedate,103)<= convert(smalldatetime,'" & dateAson & "',103) group by OrderNo,INVOICEDATE,BOOKCODE"
   While rs_data.EOF = False
    godown_issue = godown_issue + rs_data(3)
    rs_data.MoveNext
   Wend
   
   vs.TextMatrix(I, 6) = (godown_rec - godown_issue)
   vs.TextMatrix(I, 7) = (vs.TextMatrix(I, 5) - vs.TextMatrix(I, 6))
 
End If

Next


End Sub
Sub txtAlign(ST As String)
  
Label2.Visible = True

txtTot1.Visible = False
txtTot2.Visible = False
txtTot3.Visible = False
txtTot1_1.Visible = False
txtTot1_2.Visible = False


If ST = "state" Then
   txtTot1.Visible = False
   txtTot2.Visible = False
   txtTot3.Visible = False
   txtTot2.Left = 7560
   txtTot3.Left = 10260
   txtTotal.Visible = False
ElseIf ST = "dist" Then
   txtTot1.Visible = False
   txtTot2.Visible = True
   txtTot3.Visible = True
   txtTot2.Left = 7560
   txtTot3.Left = 10260
   txtTotal.Visible = True
ElseIf ST = "rep" Then
   txtTot1.Visible = True
   txtTot2.Visible = True
   txtTot3.Visible = True
   
   txtTot1.Left = 6300
   txtTot2.Left = 8460
   txtTot3.Left = 10680
   txtTotal.Visible = True
ElseIf ST = "bookwise" Then
    txtTot1.Visible = True
    txtTot2.Visible = True
    txtTot3.Visible = True
    txtTot1_1.Visible = True
    txtTot1_2.Visible = True
    txtTotal.Visible = True
    
   txtTot1_1.Left = 6180
   txtTot1_2.Left = 7500
   txtTot1.Left = 8820
   txtTot2.Left = 10200
   txtTot3.Left = 11520
   
ElseIf ST = "partywisearea" Then

    txtTot1.Visible = False
    txtTot2.Visible = False
    txtTot3.Visible = False
    txtTotal.Visible = False
    
    '================
    
    txtTot1_1.Visible = False
    txtTot1_2.Visible = False
  
    
   'txtTot1_1.Left = 6180
   'txtTot1_2.Left = 7500
   txtTot1.Left = 8980
   txtTot2.Left = 10300
   txtTot3.Left = 11600
ElseIf ST = "partywiserep" Then
    txtTotal.Visible = True
End If
  
End Sub

Private Sub cmdAdd_1_Click()

Screen.MousePointer = vbHourglass

vs.Clear
Dim str11 As String
Dim st_ As String

Dim net1, netsale As Double
Dim T1_, t2_, t3_ As Double
Dim rs_ As New ADODB.Recordset


Label3.Visible = True
Label2.Visible = True

txtTot1_1 = 0
txtTot1_2 = 0
txtTot1 = 0
txtTot2 = 0
txtTot3 = 0
txtTotal = 0

txtTot1_1.Visible = False
txtTot1_2.Visible = False
txtTot1.Visible = False
txtTot2.Visible = False
txtTot3.Visible = False
txtTotal.Visible = False




Dim str_date As String
str_date = "(invoicedate>=convert(smalldatetime,'" & txtFrom.value & "',103) and invoicedate<=convert(smalldatetime,'" & dateAson.value & "',103))"


Dim s As String
s = ""
For I = 0 To cmbAgentName.ListCount - 1
If cmbAgentName.Selected(I) = True Then
    If s = "" Then
       s = "AGENTNAME='" & cmbAgentName.List(I) & "'"
    Else
       s = s & " Or " & "AGENTNAME='" & cmbAgentName.List(I) & "'"
    End If
End If
Next


'txtTotal.Visible = True
Label2.Visible = True


If RS.State = 1 Then RS.close
Select Case cboType

Case "Party Payment Details"

    cmdRepQty.Visible = True
    
    txtTot2.Visible = True
    txtTot3.Visible = True

    txtTot2 = ""
    txtTot3 = ""
    
   T1_ = 0
   t2_ = 0
   
   Dim din As Integer
   
   din = 0
    

   vs.Cols = 9
   k1 = 1
   vs.rows = 1
     
   
   con.Execute "delete from tmpPayment"
   con.Execute "insert into tmpPayment(VoucherType , VoucherDate, Genledger, SubLedger, amount, DebitorCredit,DESCRIPTION,crno,PAYTYPE) select VoucherType , VoucherDate, Genledger, SubLedger, amount, DebitorCredit,DESCRIPTION,vsno,PAYTYPE from VOUCHERS where (GenLedger='SUNDRY DEBTORS' and  VoucherDate>=convert(smalldatetime, '" & txtFrom.value & "'  ,103) and VoucherDate<=convert(smalldatetime, '" & dateAson.value & "'  ,103) and SUBLEDGER not like '%IMPREST A/C%')"
      
    vs.TextMatrix(0, 0) = "Id"
    vs.TextMatrix(0, 1) = "Code"
    vs.TextMatrix(0, 2) = "Party Name"
    vs.TextMatrix(0, 3) = "State"
    vs.TextMatrix(0, 4) = "Date"
    vs.TextMatrix(0, 5) = "Cheque"
    vs.TextMatrix(0, 6) = "Debit"
    vs.TextMatrix(0, 7) = "Credit"
    vs.TextMatrix(0, 8) = "Payment(New/Old)"
    
    If MsgBox("For New Payment (Select yes Button)... " & vbCrLf & "Or For Both payment (Select No Button) ...", vbQuestion + vbYesNo) = vbYes Then
       str_ = "select crno,b.code,a.subledger,b.states ,a.Amount,a.debitorcredit,a.DESCRIPTION ,c.PayType,a.VoucherDate from tmpPayment  as a inner join SLEDGER as b on (a.SubLedger =b.SUBLEDGER) inner join VOUCHERS as c on (a.crNo = c.vsno) where c.PayType='n' order by a.subledger,a.VoucherDate"
    Else
       str_ = "select crno,b.code,a.subledger,b.states ,a.Amount,a.debitorcredit,a.DESCRIPTION ,c.PayType,a.VoucherDate from tmpPayment as a inner join SLEDGER as b on (a.SubLedger =b.SUBLEDGER) inner join VOUCHERS as c on (a.crNo = c.vsno) order by a.subledger,a.VoucherDate"
    End If
    
    If rs1.State = 1 Then rs1.close
    rs1.Open str_, con
    For J = 1 To rs1.RecordCount
    
        DoEvents
        DoEvents
        vs.rows = vs.rows + 1
        vs.TextMatrix(J, 0) = rs1!crno
        vs.TextMatrix(J, 1) = rs1!Code
        vs.TextMatrix(J, 2) = rs1(2)
        vs.TextMatrix(J, 3) = rs1(3)   'state
        
        vs.TextMatrix(J, 4) = rs1(8) 'date
        
        vs.TextMatrix(J, 5) = rs1(6) & "" 'narr
        
        If rs1(5) = "D" Then
          vs.TextMatrix(J, 6) = rs1(4)  'amount
        Else
          vs.TextMatrix(J, 7) = rs1(4)  'amount
        End If
        
        If Not IsNull(rs1(7)) Then
           vs.TextMatrix(J, 8) = UCase(rs1(7))
        Else
           vs.TextMatrix(J, 8) = ""
        End If
        
        
        
        If din = 0 Then
          vs.Cell(flexcpBackColor, J, 0) = vbWhite
          vs.Cell(flexcpBackColor, J, 1) = vbWhite
          vs.Cell(flexcpBackColor, J, 2) = vbWhite
          vs.Cell(flexcpBackColor, J, 3) = vbWhite
          vs.Cell(flexcpBackColor, J, 4) = vbWhite
          vs.Cell(flexcpBackColor, J, 5) = vbWhite
          vs.Cell(flexcpBackColor, J, 6) = vbWhite
          vs.Cell(flexcpBackColor, J, 7) = vbWhite
          vs.Cell(flexcpBackColor, J, 8) = vbWhite
        Else
          vs.Cell(flexcpBackColor, J, 0) = &HC0FFFF
          vs.Cell(flexcpBackColor, J, 1) = &HC0FFFF
          vs.Cell(flexcpBackColor, J, 2) = &HC0FFFF
          vs.Cell(flexcpBackColor, J, 3) = &HC0FFFF
          vs.Cell(flexcpBackColor, J, 4) = &HC0FFFF
          vs.Cell(flexcpBackColor, J, 5) = &HC0FFFF
          vs.Cell(flexcpBackColor, J, 6) = &HC0FFFF
          vs.Cell(flexcpBackColor, J, 7) = &HC0FFFF
          vs.Cell(flexcpBackColor, J, 8) = &HC0FFFF
        End If

        
        
        
        
        rs1.MoveNext
        
        DoEvents
        DoEvents
         
       If rs1.EOF = False Then
       If vs.TextMatrix(J, 0) <> rs1!Code Then
          If din = 0 Then
             din = 1
          Else
            din = 0
          End If
       End If
       End If

        
        T1_ = T1_ + IIf(vs.TextMatrix(J, 6) = "", 0, vs.TextMatrix(J, 6))
        t2_ = t2_ + IIf(vs.TextMatrix(J, 7) = "", 0, vs.TextMatrix(J, 7))

        
      Next
      
    
    txtTot2.text = T1_
    txtTot3.text = t2_
    
    vs.ColComboList(8) = "N|O"
    
    cmdPrint_7.Enabled = True
    vs.ColWidth(0) = 700
    vs.ColWidth(1) = 700
    vs.ColWidth(2) = 3700
    vs.ColWidth(3) = 1500
    vs.ColWidth(4) = 1100
    
    vs.ColWidth(5) = 2400
    vs.ColWidth(6) = 1300
    vs.ColWidth(7) = 1200
    vs.ColWidth(8) = 1000
    

   
    Screen.MousePointer = vbDefault
    Exit Sub



Case "Book Wise Sales(Area Wise)"
     
     
    txtTot1_1 = ""
    txtTot1_2 = ""
    txtTot1 = ""
    txtTot2 = ""
    txtTot3 = ""
    txtTotal = ""
     
     
     vs.Cols = 10
     k1 = 1
     vs.rows = 2
     
     vs.TextMatrix(0, 0) = "SNo."
     vs.TextMatrix(0, 1) = "Area"
     vs.TextMatrix(0, 2) = "Boook Code"
     vs.TextMatrix(0, 3) = "Boook Name"
     vs.TextMatrix(0, 4) = "Sales Amt"
     vs.TextMatrix(0, 5) = "Sales Qty"
     vs.TextMatrix(0, 6) = "SalesRet Amt"
     vs.TextMatrix(0, 7) = "SalesRet Qty"
     vs.TextMatrix(0, 8) = "Net Qty"
     vs.TextMatrix(0, 9) = "Net Amt"
     
    '-----------------------------------
     con.Execute "delete from tmpsale1"
     
     If txtbkName = "" Then
        con.Execute "insert into tmpsale1(bookcode,bookname,dist,Qty_sale,NetAmt_sale) SELECT BOOKCODE,BOOKNAME,District,sum(QUANTITY) as Qty,sum(NETAMOUNT) as NetAmt FROM invoiceBQry  where " & str_date & " group by  BOOKCODE,BOOKNAME,District"
        con.Execute "insert into tmpsale1(bookcode,bookname,dist,Qty_sale,NetAmt_sale) SELECT BOOKCODE,BOOKNAME,District,sum(QUANTITY) as Qty,sum(NETAMOUNT) as NetAmt FROM cashBQry  where " & str_date & " group by  BOOKCODE,BOOKNAME,District"
        con.Execute "insert into tmpsale1(bookcode,bookname,dist,Qty_saleret,NetAmt_saleret) SELECT BOOKCODE,BOOKNAME,District,sum(QUANTITY) as Qty,sum(NETAMOUNT) as NetAmt FROM CreditbQry where " & str_date & " group by  BOOKCODE,BOOKNAME,District"
     Else
       con.Execute "insert into tmpsale1(bookcode,bookname,dist,Qty_sale,NetAmt_sale) SELECT BOOKCODE,BOOKNAME,District,sum(QUANTITY) as Qty,sum(NETAMOUNT) as NetAmt FROM invoiceBQry  where " & str_date & " and groupcode='" & txtbkName.text & "' group by  BOOKCODE,BOOKNAME,District"
       con.Execute "insert into tmpsale1(bookcode,bookname,dist,Qty_sale,NetAmt_sale) SELECT BOOKCODE,BOOKNAME,District,sum(QUANTITY) as Qty,sum(NETAMOUNT) as NetAmt FROM cashBQry  where " & str_date & " and groupcode='" & txtbkName.text & "' group by  BOOKCODE,BOOKNAME,District"
       con.Execute "insert into tmpsale1(bookcode,bookname,dist,Qty_saleret,NetAmt_saleret) SELECT BOOKCODE,BOOKNAME,District,sum(QUANTITY) as Qty,sum(NETAMOUNT) as NetAmt FROM CreditbQry where " & str_date & " and groupcode='" & txtbkName.text & "' group by  BOOKCODE,BOOKNAME,District"
     End If

    If rs1.State = 1 Then rs1.close
    'rs1.Open "select bookcode,bookname,dist,sum(Qty_sale),sum(NetAmt_sale),sum(Qty_saleret),sum(NetAmt_saleret) from tmpsale group by bookcode,bookname,dist"
    Set rs1 = con.Execute("exec searchList '" & "bkwise" & "'")
    For J = 1 To rs1.RecordCount
    
        DoEvents
        DoEvents
        vs.rows = vs.rows + 1
        vs.TextMatrix(J, 0) = J
        vs.TextMatrix(J, 1) = rs1!dist
        vs.TextMatrix(J, 2) = rs1!Bookcode & ""
        vs.TextMatrix(J, 3) = rs1!Bookname & ""
        If Not IsNull(rs1(3)) Then
        vs.TextMatrix(J, 4) = rs1(4)
        Else
        vs.TextMatrix(J, 4) = 0
        End If
        
        If Not IsNull(rs1(4)) Then
        vs.TextMatrix(J, 5) = rs1(3)
        Else
        vs.TextMatrix(J, 5) = 0
        End If
    
        If Not IsNull(rs1(5)) Then
        vs.TextMatrix(J, 6) = rs1(6)
        Else
        vs.TextMatrix(J, 6) = 0
        End If
        
        If Not IsNull(rs1(6)) Then
        vs.TextMatrix(J, 7) = rs1(5)
        Else
        vs.TextMatrix(J, 7) = 0
        End If
    
    
    
    
        vs.TextMatrix(J, 8) = Round((vs.TextMatrix(J, 5) - vs.TextMatrix(J, 7)), 0)
        vs.TextMatrix(J, 9) = (vs.TextMatrix(J, 4) - vs.TextMatrix(J, 6))
        
        
        txtTot1_1 = Val(txtTot1_1) + Val(vs.TextMatrix(J, 4))
        txtTot1_2 = Val(txtTot1_2) + Val(vs.TextMatrix(J, 5))
        txtTot1 = Val(txtTot1) + Val(vs.TextMatrix(J, 6))
        txtTot2 = Val(txtTot2) + Val(vs.TextMatrix(J, 7))
        txtTot3 = Val(txtTot3) + Val(vs.TextMatrix(J, 8))
        txtTotal = Val(txtTotal) + Val(vs.TextMatrix(J, 9))

        DoEvents
        DoEvents
    
    
       rs1.MoveNext
    Next
     
     
     
     
      If Val(txtTot1_1) > 0 Then txtTot1_1 = Round(txtTot1_1, 0)
      If Val(txtTot1_2) > 0 Then txtTot1_2 = Round(txtTot1_2, 0)
      If Val(txtTot1) > 0 Then txtTot1 = Round(txtTot1, 0)
      If Val(txtTot2) > 0 Then txtTot2 = Round(txtTot2, 0)
      If Val(txtTot3) > 0 Then txtTot3 = Round(txtTot3, 0)
      If Val(txtTotal) > 0 Then txtTotal = Round(txtTotal, 0)
     
     
     
    cmdPrint_7.Enabled = True
    vs.ColWidth(0) = 800
    vs.ColWidth(1) = 1200
    vs.ColWidth(2) = 1000
    vs.ColWidth(3) = 3000
    vs.ColWidth(4) = 1300
    
    vs.ColWidth(5) = 1300
    vs.ColWidth(6) = 1300
    vs.ColWidth(7) = 1300
    vs.ColWidth(8) = 1300
    
    txtAlign "bookwise"
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
Case "Book Wise (Sale & Sale Return Bill List)..."
            
            Dim kk1 As Integer
            cmdRepQty.Visible = True
            
            vs.Cols = 6
            vs.rows = 1
            vs.TextMatrix(0, 0) = "BookCode"
            vs.TextMatrix(0, 1) = "BookName"
            
       If MsgBox("Want to Refresh Again ?", vbQuestion + vbYesNo) = vbYes Then
           
           Screen.MousePointer = vbHourglass
           con.Execute "delete tmpSale_SaleR"
           con.Execute "exec tmpdataForSale_and_Ret '" & txtFrom.value & "','" & dateAson.value & "'"
           con.Execute "insert into tmpSale_SaleR(bookcode,BookName) select distinct BOOKCODE,Bookname from tmpINVB_CrB"

            
            kk1 = 1
             If rs1.State = 1 Then rs1.close
            rs1.Open "select  top 5 Fyear,status_ from tmpINVB_CrB group by Fyear,status_ order by Fyear", con, adOpenDynamic, adLockOptimistic
            While rs1.EOF = False
                    If kk1 <= 2 Then
                         con.Execute "exec updateBookwise_BillList  '" & rs1!fyear & "','1', 'I'"
                         con.Execute "exec updateBookwise_BillList  '" & rs1!fyear & "','1', 'CI'"
                         vs.TextMatrix(0, 2) = "Sale - " & rs1!fyear
                         vs.TextMatrix(0, 3) = "Sale -Ret. " & rs1!fyear

                    Else
                         con.Execute "exec updateBookwise_BillList  '" & rs1!fyear & "','2', 'I'"
                         con.Execute "exec updateBookwise_BillList  '" & rs1!fyear & "','2', 'CI'"
                         vs.TextMatrix(0, 4) = "Sale - " & rs1!fyear
                         vs.TextMatrix(0, 5) = "Sale -Ret. " & rs1!fyear

                    End If
                kk1 = kk1 + 1
                rs1.MoveNext
                
            Wend
            
            Screen.MousePointer = vbDefault
            
         Else
         
             
            kk1 = 1
             If rs1.State = 1 Then rs1.close
            rs1.Open "select  top 5 Fyear,status_ from tmpINVB_CrB group by Fyear,status_ order by Fyear", con, adOpenDynamic, adLockOptimistic
            While rs1.EOF = False
                    If kk1 <= 2 Then
                         vs.TextMatrix(0, 2) = "Sale - " & rs1!fyear
                         vs.TextMatrix(0, 3) = "Sale -Ret. " & rs1!fyear
                    Else
                         vs.TextMatrix(0, 4) = "Sale - " & rs1!fyear
                         vs.TextMatrix(0, 5) = "Sale -Ret. " & rs1!fyear

                    End If
                kk1 = kk1 + 1
                rs1.MoveNext
                
            Wend
            
         End If
            
            
            
            'If rs1.State = 1 Then rs1.close
            'rs1.Open "select INVOICENO,status_,Fyear from tmpINVB_CrB order by fyear", con, adOpenDynamic, adLockOptimistic
            
            
            str11 = "SELECT distinct BOOKCODE,BOOkName,Sale1,SaleR1,Sale2,SaleR2  from  tmpSale_SaleR"
            If RS.State = 1 Then RS.close
            RS.Open str11, con, adOpenKeyset, adLockOptimistic
            For J = 1 To RS.RecordCount - 1
        
                DoEvents
                DoEvents
                vs.rows = vs.rows + 1
                
                vs.TextMatrix(J, 0) = RS(0)
                vs.TextMatrix(J, 1) = RS(1)
                 vs.TextMatrix(J, 2) = RS(2) & ""
                vs.TextMatrix(J, 3) = RS(3) & ""
                vs.TextMatrix(J, 4) = RS(4) & ""
                vs.TextMatrix(J, 5) = RS(5) & ""
                
                
                
                RS.MoveNext
           Next
            
            vs.ColWidth(0) = 800
            vs.ColWidth(1) = 2500
            vs.ColWidth(2) = 3500
            vs.ColWidth(3) = 3500
            vs.ColWidth(4) = 3500
            vs.ColWidth(5) = 3500
              
            Screen.MousePointer = vbDefault
            
            
            Exit Sub


Case "Consolidated Sales Summary..."

    Label3.Visible = False
    Label2.Visible = False
    
     txtTot1_1 = 0
     txtTot1_2 = 0
     txtTot1 = 0
     txtTot2 = 0
     txtTot3 = 0
     
   
     
     
     
     
     con.Execute "exec NetSaleSummary  '" & vs1.TextMatrix(0, 2) & "','" & vs1.TextMatrix(0, 3) & "','" & vs1.TextMatrix(1, 2) & "','" & vs1.TextMatrix(1, 3) & "','" & vs1.TextMatrix(2, 2) & "','" & vs1.TextMatrix(2, 3) & "' ,'" & vs1.TextMatrix(3, 2) & "','" & vs1.TextMatrix(3, 3) & "' ,'" & vs1.TextMatrix(4, 2) & "','" & vs1.TextMatrix(4, 3) & "','" & vs1.TextMatrix(5, 2) & "','" & vs1.TextMatrix(5, 3) & "','" & txtbkName.text & "'"
     '----------------------------------------------------------
     
     vs.Cols = 10
     vs.rows = 1
     
     vs.TextMatrix(0, 0) = "SN."
     vs.TextMatrix(0, 1) = "PCode"
     vs.TextMatrix(0, 2) = "Party Name"
     vs.TextMatrix(0, 3) = "District"
     vs.TextMatrix(0, 4) = "State"
     vs.TextMatrix(0, 5) = "" & vs1.TextMatrix(0, 1)
     vs.TextMatrix(0, 6) = "" & vs1.TextMatrix(1, 1)
     vs.TextMatrix(0, 7) = "" & vs1.TextMatrix(2, 1)
     vs.TextMatrix(0, 8) = "" & vs1.TextMatrix(3, 1)
     vs.TextMatrix(0, 9) = "Adjustment"    '"" & vs1.TextMatrix(4, 1)
     
     
    vs.ColWidth(0) = 700
    vs.ColWidth(1) = 1000
    vs.ColWidth(2) = 2400
    vs.ColWidth(3) = 2400
    
    vs.ColWidth(4) = 1200
    vs.ColWidth(5) = 1200
    vs.ColWidth(6) = 1250
    vs.ColWidth(7) = 1250
    
    st1_ = "SELECT pcode,SUBLEDGER,district,States,sum(Sale_Pr) as SaleP,sum(Sale_Curr) as SaleC,sum(SaleR_Pr)  as Return_P,sum(SaleR_Curr) as Return_C " & _
                    ",sum(Adjust) as Adjust FROM [tmpSaleSummary] group by PCode ,SUBLEDGER,district,States "
    
    
    
    If rs1.State = 1 Then rs1.close
    rs1.Open st1_, con
    For J = 1 To rs1.RecordCount
    
        DoEvents
        DoEvents
        vs.rows = vs.rows + 1
        
        vs.TextMatrix(J, 0) = J
        vs.TextMatrix(J, 1) = rs1!pcode
        vs.TextMatrix(J, 2) = rs1!subledger
        vs.TextMatrix(J, 3) = rs1!District & ""
        vs.TextMatrix(J, 4) = rs1!states & ""
        
        If Not IsNull(rs1(4)) Then
             vs.TextMatrix(J, 5) = Round(rs1(4), 2)
        Else
             vs.TextMatrix(J, 5) = 0
        End If
        
        If Not IsNull(rs1(5)) Then
        vs.TextMatrix(J, 6) = Round(rs1(5), 2)
        Else
        vs.TextMatrix(J, 6) = 0
        End If
    
        If Not IsNull(rs1(6)) Then
        vs.TextMatrix(J, 7) = Round(rs1(6), 2)
        Else
        vs.TextMatrix(J, 7) = 0
        End If
        
        If Not IsNull(rs1(7)) Then
        vs.TextMatrix(J, 8) = Round(rs1(7), 2)
        Else
        vs.TextMatrix(J, 8) = 0
        End If
    
        If Not IsNull(rs1(8)) Then
        vs.TextMatrix(J, 9) = rs1(8)
        Else
        vs.TextMatrix(J, 9) = 0
        End If
    
           
           
        txtTot1_1 = Val(txtTot1_1) + Val(vs.TextMatrix(J, 5))
        txtTot1_2 = Val(txtTot1_2) + Val(vs.TextMatrix(J, 6))
        txtTot1 = Val(txtTot1) + Val(vs.TextMatrix(J, 7))
        txtTot2 = Val(txtTot2) + Val(vs.TextMatrix(J, 8))
        txtTot3 = Val(txtTot3) + Val(vs.TextMatrix(J, 9))
    
    
        DoEvents
        DoEvents
    
    
       rs1.MoveNext
    Next
     
     
''    vs.Rows = vs.Rows + 1
''
''     vs.TextMatrix(J, 5) = txtTot1_1
''     vs.TextMatrix(J, 6) = txtTot1_2
''     vs.TextMatrix(J, 7) = txtTot1
''     vs.TextMatrix(J, 8) = txtTot2
''      vs.TextMatrix(J, 9) = txtTot3
       
    
      
    Screen.MousePointer = vbDefault
    Exit Sub
     
Case "Consolidated Sales Summary Rep. Wise ..."

    Label3.Visible = False
    Label2.Visible = False
    
     txtTot1_1 = 0
     txtTot1_2 = 0
     txtTot1 = 0
     txtTot2 = 0
     txtTot3 = 0
    
     con.Execute "exec NetSaleSummary  '" & vs1.TextMatrix(0, 2) & "','" & vs1.TextMatrix(0, 3) & "','" & vs1.TextMatrix(1, 2) & "','" & vs1.TextMatrix(1, 3) & "','" & vs1.TextMatrix(2, 2) & "','" & vs1.TextMatrix(2, 3) & "' ,'" & vs1.TextMatrix(3, 2) & "','" & vs1.TextMatrix(3, 3) & "' ,'" & vs1.TextMatrix(4, 2) & "','" & vs1.TextMatrix(4, 3) & "','" & vs1.TextMatrix(5, 2) & "','" & vs1.TextMatrix(5, 3) & "','" & txtbkName.text & "'"
     '----------------------------------------------------------
     
     vs.Cols = 7
     vs.rows = 1
     
     vs.TextMatrix(0, 0) = "SN."
     vs.TextMatrix(0, 1) = "Rep. Name"
     'vs.TextMatrix(0, 2) = "District"
     'vs.TextMatrix(0, 3) = "State"
     vs.TextMatrix(0, 2) = "" & vs1.TextMatrix(0, 1)
     vs.TextMatrix(0, 3) = "" & vs1.TextMatrix(1, 1)
     vs.TextMatrix(0, 4) = "" & vs1.TextMatrix(2, 1)
     vs.TextMatrix(0, 5) = "" & vs1.TextMatrix(3, 1)
     vs.TextMatrix(0, 6) = "ADJUSTMENT"
     
     
    vs.ColWidth(0) = 800
    vs.ColWidth(1) = 3800
    vs.ColWidth(2) = 1800
    vs.ColWidth(3) = 1800
    
    vs.ColWidth(4) = 1700
    vs.ColWidth(5) = 1700
    vs.ColWidth(6) = 1750
    'vs.ColWidth(7) = 1750
    
    st1_ = "SELECT repname,sum(Sale_Pr) as SaleP,sum(Sale_Curr) as SaleC,sum(SaleR_Pr)  as Return_P,sum(SaleR_Curr) as Return_C " & _
                    ",sum(Adjust) as Adjust FROM [tmpSaleSummary] group by repname "
    
    
    
    If rs1.State = 1 Then rs1.close
    rs1.Open st1_, con
    For J = 1 To rs1.RecordCount
    
        DoEvents
        DoEvents
        vs.rows = vs.rows + 1
        
        vs.TextMatrix(J, 0) = J
        vs.TextMatrix(J, 1) = rs1!RepName
        
        If Not IsNull(rs1(1)) Then                                   'sale-1
             vs.TextMatrix(J, 2) = Round(rs1(1), 2)
        Else
             vs.TextMatrix(J, 2) = 0
        End If
        
        If Not IsNull(rs1(2)) Then                                   'sale-2
        vs.TextMatrix(J, 3) = Round(rs1(2), 2)
        Else
        vs.TextMatrix(J, 3) = 0
        End If
    
        If Not IsNull(rs1(3)) Then                                  'saleret-1
        vs.TextMatrix(J, 4) = Round(rs1(3), 2)
        Else
        vs.TextMatrix(J, 4) = 0
        End If
        
        If Not IsNull(rs1(4)) Then                                 'saleret-2
        vs.TextMatrix(J, 5) = Round(rs1(4), 2)
        Else
        vs.TextMatrix(J, 5) = 0
        End If
        
        If Not IsNull(rs1(5)) Then                                  'Adj
        vs.TextMatrix(J, 6) = Round(rs1(5), 2)
        Else
        vs.TextMatrix(J, 6) = 0
        End If

    
       
           
           
        txtTot1_1 = Val(txtTot1_1) + Val(vs.TextMatrix(J, 4))
        txtTot1_2 = Val(txtTot1_2) + Val(vs.TextMatrix(J, 5))
        txtTot1 = Val(txtTot1) + Val(vs.TextMatrix(J, 6))
        'txtTot2 = Val(txtTot2) + Val(vs.TextMatrix(J, 7))
        'txtTot3 = Val(txtTot3) + Val(vs.TextMatrix(J, 8))
    
    
        DoEvents
        DoEvents
    
    
       rs1.MoveNext
    Next
     
     

    
      
    Screen.MousePointer = vbDefault
    Exit Sub
     

Case "Bilty Return Status..."
            
            Label2.Visible = False
            cmdRepQty.Visible = True
            
            vs.Cols = 8
            vs.TextMatrix(0, 0) = "GRNo."
            vs.TextMatrix(0, 1) = "GRDate"
            vs.TextMatrix(0, 2) = "No of BDL"
            vs.TextMatrix(0, 3) = "Rec. Date"
            vs.TextMatrix(0, 4) = "CR. No."
            vs.TextMatrix(0, 5) = "School"
            vs.TextMatrix(0, 6) = "RepName"
            vs.TextMatrix(0, 7) = "PartyName"
            vs.rows = 2
            
          st_ = ""
          
          If txtBilty_Party.text <> "" Then
             st_ = "Particulars='" & txtBilty_Party.text & "'"
          End If
          
          If txtBilty_Dist.text <> "" Then
             If st_ <> "" Then
                 st_ = st_ & " and  district='" & txtBilty_Dist.text & "'"
             Else
                st_ = "district='" & txtBilty_Dist.text & "'"
             End If
          End If
          
          
          If txtBilty_State.text <> "" Then
             If st_ <> "" Then
                 st_ = st_ & " and  states='" & txtBilty_State.text & "'"
             Else
                st_ = "states='" & txtBilty_State.text & "'"
             End If
          End If
          
          
          If txtRep.text <> "" Then
             If st_ <> "" Then
                 st_ = st_ & " and  AgentName='" & txtRep.text & "'"
             Else
                st_ = "AgentName='" & txtRep.text & "'"
             End If
          End If
          
          
          If Check1_crnote.value = 1 Then
             If st_ <> "" Then
                 st_ = st_ & " and  rno is  null"
             Else
                st_ = " rno is  null "
             End If
          End If
          
          
          
          If st_ <> "" Then
              st_ = "(" & st_ & ")"
          End If
          
            
            
            
            
            
            If st_ <> "" Then
            str11 = "SELECT GR,GR_DT,BDL,Recd_dt,RNO,ScName,AgentName,district,states,Particulars FROM biltyreturnQry where " & st_
            Else
            str11 = "SELECT GR,GR_DT,BDL,Recd_dt,RNO,ScName,AgentName,district,states,Particulars FROM biltyreturnQry"
            End If
           
           If RS.State = 1 Then RS.close
           RS.Open str11, con, adOpenKeyset, adLockOptimistic
           For J = 1 To RS.RecordCount - 1
        
                DoEvents
                DoEvents
                vs.TextMatrix(J, 0) = RS(0) & ""
                vs.TextMatrix(J, 1) = RS(1) & ""
                vs.TextMatrix(J, 2) = RS(2) & ""
                vs.TextMatrix(J, 3) = RS(3) & ""
                vs.TextMatrix(J, 4) = RS(4) & ""
                
              '===========================
                If vs.TextMatrix(J, 4) = "" Then
                For k1 = 0 To 6
                    vs.Cell(flexcpBackColor, J, k1) = vbCyan
                    DoEvents
                Next
                End If
              '===========================
              
                vs.TextMatrix(J, 5) = RS(5) & ""
                vs.TextMatrix(J, 6) = RS(6) & ""
                
                vs.TextMatrix(J, 7) = RS!Particulars & ""
                
                vs.rows = vs.rows + 1
                RS.MoveNext
           Next
            
            vs.ColWidth(0) = 1000
            vs.ColWidth(1) = 1200
            vs.ColWidth(2) = 900
            vs.ColWidth(3) = 1000
            vs.ColWidth(4) = 900
            vs.ColWidth(5) = 4000
            vs.ColWidth(6) = 1400
            vs.ColWidth(7) = 1500
            
            
            Screen.MousePointer = vbDefault
            
            Exit Sub

Case "State Wise":
     
     Label2.Visible = False
     txtTotal.Visible = False

     
     vs.Cols = 8
     
     vs.TextMatrix(0, 0) = "SNo."
     vs.TextMatrix(0, 1) = "State"
     vs.TextMatrix(0, 2) = "Sales"
     vs.TextMatrix(0, 3) = "Sales Return"
     vs.TextMatrix(0, 4) = "Net Sales"
     
     vs.TextMatrix(0, 5) = "SalesQty"
     vs.TextMatrix(0, 6) = "Return Qty"
     vs.TextMatrix(0, 7) = "Net Ret.Qty"
     
     T1_ = 0
     t2_ = 0
     t3_ = 0
     
     
     
     ''str11 = "SELECT State,sum(NetSale) FROM StateWisteNetSale where " & str_date & " group by State"
     str11 = "SELECT distinct [states] FROM SLEDGER where len(states)>0 order by states"
     If RS.State = 1 Then RS.close
     RS.Open str11, con, adOpenKeyset, adLockOptimistic
     For J = 1 To RS.RecordCount
        
        DoEvents
        DoEvents
        
        vs.TextMatrix(J, 0) = J
        vs.TextMatrix(J, 1) = RS(0)
        
        
        If rs1.State = 1 Then rs1.close
        
        If txtbkName.text = "" Then
        rs1.Open "SELECT sum(NetSale),sum(qty) FROM StateWisteNetSale where state='" & RS(0) & "' and " & str_date & "", con, adOpenKeyset, adLockOptimistic
        Else
        rs1.Open "SELECT sum(NetSale),sum(qty) FROM StateWisteNetSale where groupcode ='" & txtbkName.text & "' and state='" & RS(0) & "' and " & str_date & "", con, adOpenKeyset, adLockOptimistic
        End If
        
        If Not IsNull(rs1(0)) Then
           vs.TextMatrix(J, 2) = Round(rs1(0), 0)
        Else
           vs.TextMatrix(J, 2) = 0
        End If

        If Not IsNull(rs1(1)) Then
           vs.TextMatrix(J, 5) = Round(rs1(1), 0)
        Else
           vs.TextMatrix(J, 5) = 0
        End If

        
        
        If rs1.State = 1 Then rs1.close
        If txtbkName.text = "" Then
        rs1.Open "SELECT sum(NetSaleret),sum(qty) FROM StateWisteNetSaleret where state='" & RS(0) & "' and " & str_date & "", con, adOpenKeyset, adLockOptimistic
        Else
        rs1.Open "SELECT sum(NetSaleret),sum(qty) FROM StateWisteNetSaleret where groupcode ='" & txtbkName.text & "' and state='" & RS(0) & "' and " & str_date & "", con, adOpenKeyset, adLockOptimistic
        End If
        
        If Not IsNull(rs1(0)) Then
           vs.TextMatrix(J, 3) = Round(rs1(0), 0)
        Else
           vs.TextMatrix(J, 3) = 0
        End If
        
        If Not IsNull(rs1(1)) Then
           vs.TextMatrix(J, 6) = Round(rs1(1), 0)
        Else
           vs.TextMatrix(J, 6) = 0
        End If
        
        
        vs.TextMatrix(J, 4) = (Val(vs.TextMatrix(J, 2)) - Val(vs.TextMatrix(J, 3)))
        vs.TextMatrix(J, 7) = (Val(vs.TextMatrix(J, 5)) - Val(vs.TextMatrix(J, 6)))
        
        DoEvents
        DoEvents

        
        
        RS.MoveNext
     Next
     
    txtTotal = 0
    T1_ = 0
    t2_ = 0
    t3_ = 0
    
    For I = 1 To vs.rows - 1
    If vs.TextMatrix(I, 0) <> "" Then
        txtTotal = txtTotal + Val(vs.TextMatrix(I, 4))
        T1_ = T1_ + Val(vs.TextMatrix(I, 2))
        t2_ = t2_ + Val(vs.TextMatrix(I, 3))
        t3_ = t3_ + Val(vs.TextMatrix(I, 4))

    End If
    Next
    
      
    txtTot2 = T1_
    txtTot3 = t2_
    
    txtTot2.Visible = False
    txtTot3.Visible = False
    txtTotal.Visible = False
    
    vs.ColWidth(0) = 1100
    vs.ColWidth(1) = 4000
    vs.ColWidth(2) = 1500
    vs.ColWidth(3) = 1500
    vs.ColWidth(4) = 1500
    
    vs.ColWidth(5) = 1500
    vs.ColWidth(6) = 1500
    vs.ColWidth(7) = 1500
    
    Label3.Visible = False
    
    cmdPrint_7.Enabled = True
    Screen.MousePointer = vbDefault
    
    txtAlign "state"
    
    Exit Sub
    
    
Case "Tital Wise & Party Wise Sale & Sale Ret. Qty":
    
    
   con.Execute "delete from tmpsale_ret"
    
   If txtbkName.text <> "" Then
       con.Execute "insert into tmpsale_ret(Pcode,Party,district,states,ScName,BCode,BName,QtySale,QtySaleRet) " & _
       " SELECT substring(SUBLEDGER,1,5) as PCode,Party,District,states,ScName,BOOKCODE,BOOKNAME ,sum(QUANTITY) as Qty,0 " & _
       " from invoiceBQry  where bookcode='" & txtbkName.text & "' and  " & str_date & " group by SUBLEDGER,District,states,Party,ScName,BOOKCODE,BOOKNAME "
    
       'New Code For Cash
       con.Execute "insert into tmpsale_ret(Pcode,Party,district,states,ScName,BCode,BName,QtySale,QtySaleRet) " & _
       " SELECT substring(SUBLEDGER,1,5) as PCode,Party,District,states,ScName,BOOKCODE,BOOKNAME ,sum(QUANTITY) as Qty,0 " & _
       " from CashBQry  where bookcode='" & txtbkName.text & "' and  " & str_date & " group by SUBLEDGER,District,states,Party,ScName,BOOKCODE,BOOKNAME "
    
       con.Execute "insert into tmpsale_ret(Pcode,Party,district,states,ScName,BCode,BName,QtySaleRet,QtySale) " & _
       " SELECT substring(SUBLEDGER,1,5) as PCode,DESCFORINVOICE as Party,DISTCODE as District,states,ScName,BOOKCODE,BOOKNAME ,sum(QTY) as Qty,0 " & _
       " from CreditBSchoolWise  where bookcode='" & txtbkName.text & "' and  " & str_date & " group by SUBLEDGER,DISTCODE,states,DESCFORINVOICE,ScName,BOOKCODE,BOOKNAME "
   End If
   
      
    
    
    
     st_ = ""
     
     vs.Cols = 8
     vs.TextMatrix(0, 0) = "PCode"
     vs.TextMatrix(0, 1) = "Party"
     vs.TextMatrix(0, 2) = "District"
     vs.TextMatrix(0, 3) = "State"
     vs.TextMatrix(0, 4) = "School Name"
     'vs.TextMatrix(0, 5) = "BCode"
     'vs.TextMatrix(0, 6) = "BName"
     vs.TextMatrix(0, 5) = "Qty(Sale)"
     vs.TextMatrix(0, 6) = "Qty(SaleRet)"
     vs.TextMatrix(0, 7) = "Net Sales"
     
     vs.rows = 1
     k1 = 1
     
     
     str11 = "SELECT Pcode,Party,district,states,ScName,BCode,BName,QtySale,QtySaleRet  FROM tmpsale_ret"
     If RS.State = 1 Then RS.close
     RS.Open str11, con, adOpenKeyset, adLockOptimistic
     For J = 1 To RS.RecordCount
        
            DoEvents
            DoEvents
            vs.rows = vs.rows + 1
            vs.TextMatrix(k1, 0) = RS!pcode
            vs.TextMatrix(k1, 1) = RS!party
            vs.TextMatrix(k1, 2) = RS!District
            vs.TextMatrix(k1, 3) = RS!states
            vs.TextMatrix(k1, 4) = RS!scname
            'vs.TextMatrix(k1, 5) = RS!bcode
            'vs.TextMatrix(k1, 6) = RS!BName
            vs.TextMatrix(k1, 5) = RS!QtySale
            vs.TextMatrix(k1, 6) = RS!QtySaleRet
            vs.TextMatrix(k1, 7) = (RS!QtySale - RS!QtySaleRet)
            k1 = k1 + 1
            
            
            
            DoEvents
            DoEvents

        RS.MoveNext
     Next
     
    
  

    vs.ColWidth(0) = 800
    vs.ColWidth(1) = 2500
    vs.ColWidth(2) = 2000
    vs.ColWidth(3) = 2000
    vs.ColWidth(4) = 4000
    vs.ColWidth(5) = 1000
    vs.ColWidth(6) = 1000
    vs.ColWidth(7) = 1000
    
    cmdPrint_7.Enabled = True
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

    
    
    
    
    
     
Case "District Wise":

     st_ = ""
     
     vs.Cols = 6
     vs.TextMatrix(0, 0) = "SNo."
     vs.TextMatrix(0, 1) = "District"
     vs.TextMatrix(0, 2) = "State"
     vs.TextMatrix(0, 3) = "Net Sales"
     vs.TextMatrix(0, 4) = "Sales Return"
     vs.TextMatrix(0, 5) = "Net Sales"
   
     
     k1 = 1
     
     ''If rs_.State = 1 Then rs_.close
     ''rs_.Open "select distinct DISTCODE,states from sledger where gledger='SUNDRY DEBTORS'", con, adOpenDynamic, adLockOptimistic
     
     'str11 = "SELECT District,sum(NETAMOUNT) as NetSale FROM invoiceaQry where " & str_date & " group by district"
     'str11 = "SELECT distinct DISTCODE FROM [SLEDGER] where len(DISTCODE)>0 order by DISTCODE"
    
     
     If RS.State = 1 Then RS.close
     'RS.Open "SELECT District,states,invoicedate,groupcode,NetAmt,Qty,category FROM DistrictwiseSales", con, adOpenKeyset, adLockOptimistic
     DoEvents
     Set RS = con.Execute("exec DisWiseSalesSp '" & txtbkName.text & "','" & txtFrom.value & "','" & dateAson.value & "'")
     DoEvents

     
     For J = 1 To RS.RecordCount
        
        
        netsale = RS(2)
        net1 = RS(3)
        
        
         
'        If txtbkname = "" Then
'           rs1.Open "SELECT sum(NETAMOUNT) FROM invoicebQry where District='" & RS(0) & "' and " & str_date & "", con, adOpenKeyset, adLockOptimistic
'        Else
'           rs1.Open "SELECT sum(NETAMOUNT) FROM invoicebQry where GROUPCODE='" & txtbkname.Text & "' and District='" & RS(0) & "' and " & str_date & "", con, adOpenKeyset, adLockOptimistic
'        End If
'
'        If Not IsNull(rs1(0)) Then
'            netsale = Round(rs1(0), 2)
'        Else
'           netsale = 0
'        End If
'
'
'        If rs1.State = 1 Then rs1.close
'         If txtbkname = "" Then
'        rs1.Open "SELECT sum(NETAMOUNT) FROM creditbQry where District='" & RS(0) & "' and " & str_date & "", con, adOpenKeyset, adLockOptimistic
'        Else
'        rs1.Open "SELECT sum(NETAMOUNT) FROM creditbQry where GROUPCODE='" & txtbkname.Text & "' and District='" & RS(0) & "' and " & str_date & "", con, adOpenKeyset, adLockOptimistic
'        End If
'        If Not IsNull(rs1(0)) Then
'           net1 = Round(rs1(0), 2)
'        Else
'           net1 = 0
'        End If
        
        
        
        
        If (netsale > 0 Or net1 > 0) Then
        
'            rs_.MoveFirst
'            rs_.Find "distcode='" & RS(0) & "'"
'            If rs_.EOF = False Then
'            st_ = rs_!states
'            Else
'            st_ = "-"
'            End If

            DoEvents
            DoEvents
            
            vs.TextMatrix(k1, 0) = k1
            vs.TextMatrix(k1, 1) = RS(0)
            'vs.TextMatrix(k1, 2) = st_
            vs.TextMatrix(k1, 2) = RS(1)
            vs.TextMatrix(k1, 3) = netsale
            vs.TextMatrix(k1, 4) = net1
            vs.TextMatrix(k1, 5) = (Val(vs.TextMatrix(k1, 3)) - Val(vs.TextMatrix(k1, 4)))
            k1 = k1 + 1
            
            DoEvents
            DoEvents

        End If
        
        RS.MoveNext
     Next
     
    
    txtTotal = 0
    T1_ = 0
    t2_ = 0
    For I = 1 To vs.rows - 1
    If vs.TextMatrix(I, 0) <> "" Then
       txtTotal = txtTotal + Val(vs.TextMatrix(I, 5))
       T1_ = T1_ + Val(vs.TextMatrix(I, 3))
       t2_ = t2_ + Val(vs.TextMatrix(I, 4))
       
    End If
    Next

    vs.ColWidth(0) = 800
    vs.ColWidth(1) = 2500
    vs.ColWidth(2) = 2500
    vs.ColWidth(3) = 2800
    vs.ColWidth(4) = 2800
    vs.ColWidth(5) = 2800
    
    cmdPrint_7.Enabled = True
    
    txtTot2 = Round(T1_, 0)
    txtTot3 = Round(t2_, 0)
    txtTotal = Round(txtTotal, 0)
    
    
    txtAlign "dist"
    Screen.MousePointer = vbDefault
    
    Exit Sub



Case "Rep Wise":

     
     
     vs.Cols = 8
     vs.TextMatrix(0, 0) = "SNo."
     vs.TextMatrix(0, 1) = "Representative"
     vs.TextMatrix(0, 2) = "Sale"
     vs.TextMatrix(0, 3) = "Sale Return"
     vs.TextMatrix(0, 4) = "Credit Note"
     vs.TextMatrix(0, 5) = "Debit Note"
     vs.TextMatrix(0, 6) = "Sp.Amt."
     vs.TextMatrix(0, 7) = "Net Sales"
     
     k1 = 1
     DebitFromRepSale
     
     '==============================================================================================
     con.Execute "delete from tmpRepwiseSale where uid=" & UId & ""
     
     'sale  & cash counter
     con.Execute "insert into tmpRepwiseSale(RepName,INVOICEDATE,Sale,SaleRet,DrNote,CrNote,NetSale,GP,uid)" & _
     " SELECT  Representative,INVOICEDATE,NetSale,0,0,0,0,GROUPCODE," & UId & " FROM RepWisteNetSale "

     con.Execute "insert into tmpRepwiseSale(RepName,INVOICEDATE,Sale,SaleRet,DrNote,CrNote,NetSale,GP,uid)" & _
     " SELECT  AgentName,INVOICEDATE,0,NetAmount,0,0,0,GROUPCODE," & UId & " FROM RepwiseBookWiseSale_Ret"
     
     If txtbkName <> "" Then
         con.Execute "insert into tmpRepwiseSale(RepName,INVOICEDATE,Sale,SaleRet,CrNote,DrNote,NetSale,GP,uid)" & _
         " SELECT  AgentName,invoicedate,0,0,amount,0,0,'" & txtbkName.text & "'," & UId & " FROM PartyCreditRegisternew where  saletype='" & txtbkName.text & "'  and " & debitForAgn & ""
         
         
        con.Execute "insert into tmpRepwiseSale(RepName,INVOICEDATE,Sale,SaleRet,DrNote,CrNote,NetSale,GP,uid) " & _
        " SELECT debitRegister.Agentname,debitRegister.invoicedate,0,0,debitRegister.amount,0,0,'" & txtbkName.text & "'," & UId & " FROM debitRegister LEFT OUTER JOIN dbo.debitNotDet ON debitRegister.DNN = debitNotDet.DNN " & _
        " where (debitNotDet.RepName Is Null and debitRegister.saletype='" & txtbkName.text & "'  and " & debitForAgn & ")"
        
         con.Execute "insert into tmpRepwiseSale(RepName,INVOICEDATE,Sale,SaleRet,DrNote,CrNote,NetSale,GP,uid) " & _
        " SELECT debitNotDet.RepName,debitRegister.invoicedate,0,0,debitNotDet.Amount,0,0,'" & txtbkName.text & "'," & UId & " FROM debitRegister LEFT OUTER JOIN dbo.debitNotDet ON debitRegister.DNN = debitNotDet.DNN " & _
        " where (debitNotDet.RepName Is not Null and debitRegister.saletype='" & txtbkName.text & "'  and " & debitForAgn & ")"
       
        
        
    Else
    
     con.Execute "insert into tmpRepwiseSale(RepName,INVOICEDATE,Sale,SaleRet,CrNote,DrNote,NetSale,GP,uid)" & _
     " SELECT  AgentName,invoicedate,0,0,amount,0,0,'-'," & UId & " FROM PartyCreditRegisternew where  " & debitForAgn & ""
    
         
    con.Execute "insert into tmpRepwiseSale(RepName,INVOICEDATE,Sale,SaleRet,DrNote,CrNote,NetSale,GP,uid) " & _
    " SELECT debitRegister.Agentname,debitRegister.invoicedate,0,0,debitRegister.amount,0,0,'-'," & UId & " FROM debitRegister LEFT OUTER JOIN dbo.debitNotDet ON debitRegister.DNN = debitNotDet.DNN " & _
    " where (debitNotDet.RepName Is Null  and " & debitForAgn & ")"
    
     con.Execute "insert into tmpRepwiseSale(RepName,INVOICEDATE,Sale,SaleRet,DrNote,CrNote,NetSale,GP,uid) " & _
    " SELECT debitNotDet.RepName,debitRegister.invoicedate,0,0,debitNotDet.Amount,0,0,'-'," & UId & " FROM debitRegister LEFT OUTER JOIN dbo.debitNotDet ON debitRegister.DNN = debitNotDet.DNN " & _
    " where (debitNotDet.RepName Is not Null and " & debitForAgn & ")"

        
    
   End If
     
    
     
     If txtbkName <> "" Then
        str11 = "SELECT RepName,sum(Sale),sum(Saleret),sum(DrNote),sum(CrNote) FROM tmpRepwiseSale where len(RepName)>1 and  uid = " & UId & " and gp ='" & txtbkName.text & "' and " & str_date & " group by RepName order by repname"
     Else
        str11 = "SELECT RepName,sum(Sale),sum(Saleret),sum(DrNote),sum(CrNote) FROM tmpRepwiseSale where len(RepName)>1 and uid = " & UId & " and  " & str_date & " group by RepName order by repname"
     End If
     
     dt_ = "(ddate>=convert(smalldatetime,'" & txtFrom.value & "',103) and ddate<=convert(smalldatetime,'" & dateAson.value & "',103))"
     If rs1.State = 1 Then rs1.close
     rs1.Open "SELECT RepName,sum(NetBalance+AdvAmt) as SpAmt FROM DonnationMain group by RepName", con, adOpenDynamic, adLockOptimistic
   

     RS.Open str11, con, adOpenKeyset, adLockOptimistic
     For J = 1 To RS.RecordCount

        T1_ = 0
        t2_ = 0
        net1 = 0

        'If rs1.State = 1 Then rs1.close
        'rs1.Open "SELECT sum(NetSale) FROM RepWisteNetSale where Representative='" & RS(0) & "' and " & str_date & "", con, adOpenKeyset, adLockOptimistic
        'If Not IsNull(rs1(0)) Then
           T1_ = Round(RS(1), 0)
        'Else
        '   T1_ = 0
        'End If


        'If rs1.State = 1 Then rs1.close
        'rs1.Open "SELECT sum(NetSale) FROM RepWisteNetSale_ret where Representative='" & RS(0) & "' and " & str_date & "", con, adOpenKeyset, adLockOptimistic
        'If Not IsNull(rs1(0)) Then
           t2_ = Round(RS(2), 0)
        'Else
        '   t2_ = 0
        'End If


        
        
        'net1 = RS(3) - RS(4)
        
        If net1 < 0 Then
        '   net1 = Abs(net1)
        End If

        'If debitForAgn <> "" Then
        '    If rs1.State = 1 Then rs1.close
        '    rs1.Open "SELECT sum(Amount) FROM PartyCreditRegisternew where " & debitForAgn & " and RepName='" & RS(0) & "' and " & str_date & "", con, adOpenKeyset, adLockOptimistic
        '    If Not IsNull(rs1(0)) Then
        '       net1 = rs1(0)
        '    End If

        '    If rs1.State = 1 Then rs1.close
        '    rs1.Open "SELECT sum(Amount) FROM debitRegister where " & debitForAgn & " and Agentname='" & RS(0) & "' and " & str_date & "", con, adOpenKeyset, adLockOptimistic
        '    If Not IsNull(rs1(0)) Then
        '       net1 = net1 - rs1(0)
        '    End If
        'End If
        
        If "GIRIJESH MANI TRIPATHI" = RS(0) Then
         '  MsgBox "A"
        End If
        
        
        ''If (T1_ > 0 Or t2_ > 0 Or net1 > 0) Then
            DoEvents

            vs.TextMatrix(k1, 0) = k1
            vs.TextMatrix(k1, 1) = RS(0)
            vs.TextMatrix(k1, 2) = T1_
            vs.TextMatrix(k1, 3) = t2_
            vs.TextMatrix(k1, 4) = Round(RS(4), 0)
            vs.TextMatrix(k1, 5) = Round(RS(3), 0)
            
            'If RS(0) = "ABHIJEET PANDEY" Then
            '   MsgBox "aaa"
            'End If
            
            If rs1.EOF = False Then
                rs1.MoveFirst
                rs1.Find "repname='" & RS(0) & "'"
                If rs1.EOF = False Then
                   vs.TextMatrix(k1, 6) = Round(rs1(1), 0)
                Else
                   vs.TextMatrix(k1, 6) = 0
                End If
            End If
            
            net1 = ((T1_ + Round(RS(3), 0)) - (Val(vs.TextMatrix(k1, 6)) + t2_ + Round(RS(4), 0)))
            vs.TextMatrix(k1, 7) = Round(net1, 2)
            k1 = k1 + 1

            DoEvents
            DoEvents
            DoEvents

        ''End If





        RS.MoveNext
     Next
     
     
    txtTotal = 0
    T1_ = 0
    t2_ = 0
    t3_ = 0
    
    For I = 1 To vs.rows - 1
    If vs.TextMatrix(I, 0) <> "" Then
    
    
    T1_ = T1_ + Val(vs.TextMatrix(I, 2))
    t2_ = t2_ + Val(vs.TextMatrix(I, 3))
    t3_ = t3_ + Val(vs.TextMatrix(I, 4))
    txtTotal = txtTotal + Val(vs.TextMatrix(I, 6))

    
    End If
    Next
    
    
    vs.ColWidth(0) = 1100
    vs.ColWidth(1) = 4700
    vs.ColWidth(2) = 1600
    vs.ColWidth(3) = 2100
    vs.ColWidth(4) = 2200
    vs.ColWidth(5) = 950
    vs.ColWidth(6) = 950
    vs.ColWidth(7) = 1100
    
    txtTot1 = T1_
    txtTot2 = t2_
    txtTot3 = t3_
    
    txtAlign "rep"
     
    cmdPrint_7.Enabled = True
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
    
    

Case "Book Wise Sales":
     
     Dim Q1, sum1 As Double
     
     vs.Cols = 9
     k1 = 1
          
     vs.rows = 2
     vs.TextMatrix(0, 0) = "SNo."
     vs.TextMatrix(0, 1) = "Boook Code"
     vs.TextMatrix(0, 2) = "Boook Name"
     
     vs.TextMatrix(0, 3) = "Sales Amt"
     vs.TextMatrix(0, 4) = "Sales Qty"
     
     vs.TextMatrix(0, 5) = "SalesRet Amt"
     vs.TextMatrix(0, 6) = "SalesRet Qty"
     
     vs.TextMatrix(0, 7) = "Net Amt"
     vs.TextMatrix(0, 8) = "Net Qty"
               
     
     If RS.State = 1 Then RS.close
     
     If txtbkName <> "" Then
         DoEvents
         Set RS = con.Execute("exec BookwiseSalesSp '" & txtbkName.text & "','" & txtFrom.value & "','" & dateAson.value & "'")
         DoEvents
     Else
         DoEvents
         Set RS = con.Execute("exec BookwiseSalesSp '','" & txtFrom.value & "','" & dateAson.value & "'")
         DoEvents
     End If
    
     For J = 1 To RS.RecordCount
         
        
            T1_ = RS(3)
            netsale = RS(4)
            
            Q1 = RS(5)
            sum1 = RS(6)
            

      
        
        
        
        If (netsale > 0 Or sum1 > 0) Then
        
            DoEvents
            DoEvents
            
            vs.TextMatrix(k1, 0) = k1
            vs.TextMatrix(k1, 1) = RS(0)
            vs.TextMatrix(k1, 2) = RS(1)
            
            vs.TextMatrix(k1, 3) = Round(netsale, 2)
            vs.TextMatrix(k1, 4) = T1_
            
            vs.TextMatrix(k1, 5) = Round(sum1, 2)
            vs.TextMatrix(k1, 6) = Q1
            
            
            vs.TextMatrix(k1, 7) = Round((netsale - sum1), 0)
            vs.TextMatrix(k1, 8) = (T1_ - Q1)
            
            txtTot1_1 = Val(txtTot1_1) + Val(vs.TextMatrix(k1, 3))
            txtTot1_2 = Val(txtTot1_2) + Val(vs.TextMatrix(k1, 4))
            txtTot1 = Val(txtTot1) + Val(vs.TextMatrix(k1, 5))
            txtTot2 = Val(txtTot2) + Val(vs.TextMatrix(k1, 6))
            txtTot3 = Val(txtTot3) + Val(vs.TextMatrix(k1, 7))
            txtTotal = Val(txtTotal) + Val(vs.TextMatrix(k1, 8))
            
            k1 = k1 + 1
            vs.rows = vs.rows + 1
            DoEvents
            DoEvents
        
        End If
        
        RS.MoveNext
     Next
     
     
        txtTot1_1 = Round(txtTot1_1, 0)
        txtTot1_2 = Round(txtTot1_2, 0)
        txtTot1 = Round(txtTot1, 0)
        txtTot2 = Round(txtTot2, 0)
        txtTot3 = Round(txtTot3, 0)
        txtTotal = Round(txtTotal, 0)
      
     

    cmdPrint_7.Enabled = True
    'vs.FormatString = "SNo.|<Boook Code|<Boook Name|>Net Sales|>Net Qty"
    vs.ColWidth(0) = 850
    vs.ColWidth(1) = 1200
    vs.ColWidth(2) = 4000
    vs.ColWidth(3) = 1300
    vs.ColWidth(4) = 1300
    
    vs.ColWidth(5) = 1300
    vs.ColWidth(6) = 1300
    vs.ColWidth(7) = 1300
    vs.ColWidth(8) = 1300
    
    txtAlign "bookwise"
    
    Screen.MousePointer = vbDefault
    Exit Sub


Case "Party Wise Area Wise Net Sale..."
      
     DebitFromRepSale
     
     Dim netamt_ret, sum2 As Double
     Dim T1 As Integer
     Dim rs_10 As New ADODB.Recordset
     
     sum2 = 0
     T1 = 1
     vs.Cols = 9
     vs.rows = 1
     
     vs.TextMatrix(0, 0) = "SNo."
     vs.TextMatrix(0, 1) = "Customer Name"
     vs.TextMatrix(0, 2) = "Area"
     vs.TextMatrix(0, 3) = "State"
     vs.TextMatrix(0, 4) = "Sales"
     vs.TextMatrix(0, 5) = "Sales Ret."
     vs.TextMatrix(0, 6) = "Adjustment"
     vs.TextMatrix(0, 7) = "Net Sales"
     vs.TextMatrix(0, 8) = "Gross Sales"
     
     
''''New Coding
   con.Execute "delete from tmpPartyWiseSale"
   
   
   If txtbkName.text = "" Then
   
        con.Execute "insert into tmpPartyWiseSale(INVOICEDATE,subledger,District,states,SaleRet)" & _
       " SELECT INVOICEDATE,SUBLEDGER,District,states,sum(NETAMOUNT) as Amt  FROM CreditbQry group by INVOICEDATE,SUBLEDGER,District,states"
    
        If debitForAgn <> "" Then
    
        con.Execute "insert into tmpPartyWiseSale(INVOICEDATE,subledger,District,states,crAmt)" & _
        " SELECT PartyCreditRegister.invoicedate, PartyCreditRegister.SUBLEDGER,SLEDGER.DISTCODE, SLEDGER.states," & _
        "PartyCreditRegister.netamount FROM PartyCreditRegister   INNER JOIN SLEDGER ON PartyCreditRegister.SUBLEDGER = SLEDGER.SUBLEDGER where " & debitForAgn & ""
    
        con.Execute "insert into tmpPartyWiseSale(INVOICEDATE,subledger,District,states,DrAmt)" & _
        " SELECT debitRegister.invoicedate, debitRegister.SUBLEDGER,SLEDGER.DISTCODE, SLEDGER.states," & _
        "debitRegister.Amount FROM debitRegister   INNER JOIN SLEDGER ON debitRegister.SUBLEDGER = SLEDGER.SUBLEDGER where " & debitForAgn & ""
    
       End If
    
       con.Execute "insert into tmpPartyWiseSale(INVOICEDATE,subledger,District,states,Sale,grossSale)" & _
       " SELECT INVOICEDATE,SUBLEDGER,District,states,sum(NETAMOUNT) as Amt,sum(amount)  FROM invoiceBQry group by INVOICEDATE,SUBLEDGER,District,states"
   
       con.Execute "insert into tmpPartyWiseSale(INVOICEDATE,subledger,District,states,Sale,grossSale)" & _
       " SELECT INVOICEDATE,SUBLEDGER,District,states,sum(NETAMOUNT) as Amt,sum(amount)  FROM cashBQry  group by INVOICEDATE,SUBLEDGER,District,states"
   
   Else
        con.Execute "insert into tmpPartyWiseSale(INVOICEDATE,subledger,District,states,SaleRet)" & _
       " SELECT INVOICEDATE,SUBLEDGER,District,states,sum(NETAMOUNT) as Amt  FROM CreditbQry where groupcode='" & txtbkName.text & "' group by INVOICEDATE,SUBLEDGER,District,states"
    
        If debitForAgn <> "" Then
    
        con.Execute "insert into tmpPartyWiseSale(INVOICEDATE,subledger,District,states,crAmt)" & _
        " SELECT PartyCreditRegister.invoicedate, PartyCreditRegister.SUBLEDGER,SLEDGER.DISTCODE, SLEDGER.states," & _
        "PartyCreditRegister.netamount FROM PartyCreditRegister   INNER JOIN SLEDGER ON PartyCreditRegister.SUBLEDGER = SLEDGER.SUBLEDGER where saletype='" & txtbkName.text & "' and " & debitForAgn & ""
    
        con.Execute "insert into tmpPartyWiseSale(INVOICEDATE,subledger,District,states,DrAmt)" & _
        " SELECT debitRegister.invoicedate, debitRegister.SUBLEDGER,SLEDGER.DISTCODE, SLEDGER.states," & _
        "debitRegister.Amount FROM debitRegister   INNER JOIN SLEDGER ON debitRegister.SUBLEDGER = SLEDGER.SUBLEDGER where saletype='" & txtbkName.text & "' and " & debitForAgn & ""
    
       End If
    
       con.Execute "insert into tmpPartyWiseSale(INVOICEDATE,subledger,District,states,Sale,grossSale)" & _
       " SELECT INVOICEDATE,SUBLEDGER,District,states,sum(NETAMOUNT) as Amt,sum(amount)  FROM invoiceBQry where groupcode='" & txtbkName.text & "' group by INVOICEDATE,SUBLEDGER,District,states"
       
       con.Execute "insert into tmpPartyWiseSale(INVOICEDATE,subledger,District,states,Sale,grossSale)" & _
       " SELECT INVOICEDATE,SUBLEDGER,District,states,sum(NETAMOUNT) as Amt,sum(amount)  FROM cashBQry where groupcode='" & txtbkName.text & "' group by INVOICEDATE,SUBLEDGER,District,states"

   End If
   


    

     'str11 = "SELECT party,ADDRESS3 AS City,states,SUBLEDGER,code FROM SLEDGER where gledger='SUNDRY DEBTORS'  order by SUBLEDGER"
     str11 = "SELECT SUBLEDGER as party,District AS City,states,substring(SUBLEDGER,1,5) as Code,sum(Sale),sum(SaleRet),sum(Adj),sum(CrAmt),sum(DrAmt),sum(grossSale) FROM tmpPartyWiseSale where " & str_date & "  group by SUBLEDGER,District,states"

     If RS.State = 1 Then RS.close
     RS.Open str11, con, adOpenKeyset, adLockOptimistic
     For J = 1 To RS.RecordCount



        netamt_ret = 0
        'If rs1.State = 1 Then rs1.close
        'rs1.Open "select sum(netamount) from CreditbQry where " & str_date & " and subledger='" & RS!SUBLEDGER & "'", con
        If Not IsNull(RS(5)) Then
           netamt_ret = Round(RS(5), 0)
        End If

        net1 = 0
        'If debitForAgn <> "" Then
        '    If rs1.State = 1 Then rs1.close
        '    rs1.Open "SELECT sum(netAmount) FROM PartyCreditRegister where " & debitForAgn & " and SUBLEDGER='" & RS!SUBLEDGER & "' and " & str_date & "", con, adOpenKeyset, adLockOptimistic
            If Not IsNull(RS(7)) Then
               net1 = RS(7)
            End If
        '    If rs1.State = 1 Then rs1.close
        '    rs1.Open "SELECT sum(Amount) FROM debitRegister where " & debitForAgn & " and SUBLEDGER='" & RS!SUBLEDGER & "' and " & str_date & "", con, adOpenKeyset, adLockOptimistic
            If Not IsNull(RS(8)) Then
               net1 = net1 - RS(8)
           End If
        'End If

       netsale = 0
       'str11 = "SELECt sum(NetAmount) FROM invoiceBQry where " & str_date & " and SUBLEDGER='" & RS!SUBLEDGER & "'"
        'f rs_10.State = 1 Then rs_10.close
       'rs_10.Open str11, con, adOpenKeyset, adLockOptimistic
       If Not IsNull(RS(4)) Then
          sum2 = (RS(4) - (netamt_ret + net1))
          netsale = RS(4)
       Else
        sum2 = (netamt_ret + net1) * -1
       End If
       
       

        If (netsale > 0 Or netamt_ret > 0 Or net1 > 0) Then
            DoEvents
            vs.rows = vs.rows + 1
            vs.TextMatrix(T1, 5) = 0
            vs.TextMatrix(T1, 0) = J
            vs.TextMatrix(T1, 1) = Trim(RS!Code) & " - " & Trim(Mid(RS!party, 6))
            vs.TextMatrix(T1, 2) = RS!city & ""
            vs.TextMatrix(T1, 3) = RS!states & ""
            vs.TextMatrix(T1, 4) = Round(netsale, 0)
            vs.TextMatrix(T1, 5) = netamt_ret
            vs.TextMatrix(T1, 6) = Round(net1, 0)
            vs.TextMatrix(T1, 7) = Round(sum2, 0)      '(Round(rs_10(3), 0) - (netamt_ret + Val(vs.TextMatrix(T1, 6))))
            If Not IsNull(RS(9)) Then
               vs.TextMatrix(T1, 8) = Round(RS(9), 0)    ' gross sale
            Else
               vs.TextMatrix(T1, 8) = 0
            End If
            T1 = T1 + 1
            DoEvents
            DoEvents
        End If

        RS.MoveNext
     Next
     
     
    txtTotal = 0
    txtTot1 = 0
    txtTot2 = 0
    txtTot3 = 0
    For I = 1 To vs.rows - 1
    If vs.TextMatrix(I, 0) <> "" Then
       txtTotal = txtTotal + Val(vs.TextMatrix(I, 7))
       
       txtTot1 = txtTot1 + Val(vs.TextMatrix(I, 4))
       txtTot2 = txtTot2 + Val(vs.TextMatrix(I, 5))
       txtTot3 = txtTot3 + Val(vs.TextMatrix(I, 6))
    
    End If
    Next
    
    rtype = "PartyWiseArea"
    cmdPrint_7.Enabled = True
    
    vs.ColWidth(0) = 700
    vs.ColWidth(1) = 4200
    vs.ColWidth(2) = 2000
    vs.ColWidth(3) = 2000
    vs.ColWidth(4) = 1200
    vs.ColWidth(5) = 1200
    vs.ColWidth(6) = 1200
    vs.ColWidth(7) = 1200
    vs.ColWidth(8) = 1200
    
    debitForAgn = ""
    
    txtAlign "partywisearea"
    
    Screen.MousePointer = vbDefault
    Exit Sub

Case "Party Wise Area & Rep. Wise Net Sale..."
      
     DebitFromRepSale
     
     Dim netamt_ret_, sum2_ As Double
     Dim dr_, cr_ As Double
     Dim T1_1 As Integer
     Dim rs_101 As New ADODB.Recordset
     Dim saleQty, saleRetQty As Double
     
     sum2 = 0
     T1 = 1
     vs.Cols = 9 + 3
     vs.rows = 1
     
     vs.TextMatrix(0, 0) = "SNo."
     vs.TextMatrix(0, 1) = "Customer Name"
     vs.TextMatrix(0, 2) = "Rep. Name"
     vs.TextMatrix(0, 3) = "Area"
     vs.TextMatrix(0, 4) = "State"
     vs.TextMatrix(0, 5) = "Sales"
     vs.TextMatrix(0, 6) = "Sales Ret."
     vs.TextMatrix(0, 7) = "Adjustment"
     vs.TextMatrix(0, 8) = "Net Sales"
     vs.TextMatrix(0, 9) = "SalesQty"
     vs.TextMatrix(0, 10) = "SRetQty"
     vs.TextMatrix(0, 11) = "NetQty"
     
     con.Execute "delete from tmptbl where len(subledger)>0"
     
     If rs_.State = 1 Then rs_.close
     rs_.Open "select distinct subledger,ADDRESS3 as area,states from sledger where gledger='SUNDRY DEBTORS'", con, adOpenDynamic, adLockOptimistic
     
     
    If txtbkName.text = "" Then
     
     con.Execute "insert into tmptbl(INVOICEDATE,SUBLEDGER,RepName,NETAMOUNT,Type_,QUANTITY)" & _
     "select INVOICEDATE,SUBLEDGER,agentname,sum(NETAMOUNT),'Ret',sum(QUANTITY) from CreditbQry group by INVOICEDATE,SUBLEDGER,agentname"
        
     If debitForAgn <> "" Then
        
        
        con.Execute "insert into tmptbl(INVOICEDATE,SUBLEDGER,RepName,NETAMOUNT,Type_,QUANTITY)" & _
        " select INVOICEDATE,SUBLEDGER,RepName,sum(Amount),'CR',0 from PartyCreditRegisterNew where " & debitForAgn & " group by INVOICEDATE,SUBLEDGER,RepName"
        
        
        con.Execute "insert into tmptbl(INVOICEDATE,SUBLEDGER,RepName,NETAMOUNT,Type_,QUANTITY)" & _
        " SELECT debitRegister.invoicedate,debitRegister.PGLD,debitRegister.Agentname,sum(debitRegister.amount),'DR',0 FROM debitRegister LEFT OUTER JOIN dbo.debitNotDet ON debitRegister.DNN = debitNotDet.DNN " & _
        " Where debitNotDet.RepName Is Null and " & debitForAgn & " group by INVOICEDATE,PGLD,agentname"
        
        con.Execute "insert into tmptbl(INVOICEDATE,SUBLEDGER,RepName,NETAMOUNT,Type_,QUANTITY)" & _
        " SELECT debitRegister.invoicedate,debitRegister.PGLD,debitNotDet.repname,sum(debitNotDet.Amount),'DR',0 FROM debitRegister LEFT OUTER JOIN dbo.debitNotDet ON debitRegister.DNN = debitNotDet.DNN " & _
        " Where debitNotDet.RepName Is not Null and " & debitForAgn & " group by INVOICEDATE,PGLD,debitNotDet.RepName"
        
        
     End If
       
     con.Execute "insert into tmptbl(INVOICEDATE,SUBLEDGER,RepName,NETAMOUNT,Type_,QUANTITY)" & _
     "select INVOICEDATE,SUBLEDGER,agentname,sum(NETAMOUNT),'Sale',sum(QUANTITY) from invoiceBQry group by INVOICEDATE,SUBLEDGER,agentname"
        
     con.Execute "insert into tmptbl(INVOICEDATE,SUBLEDGER,RepName,NETAMOUNT,Type_QUANTITY)" & _
     "select INVOICEDATE,SUBLEDGER,agentname,sum(NETAMOUNT),'Sale',sum(QUANTITY) from cashBQry group by INVOICEDATE,SUBLEDGER,agentname"
     
    Else
     
     con.Execute "insert into tmptbl(INVOICEDATE,SUBLEDGER,RepName,NETAMOUNT,Type_,QUANTITY)" & _
     "select INVOICEDATE,SUBLEDGER,agentname,sum(NETAMOUNT),'Ret',sum(QUANTITY) from CreditbQry where GROUPCODE='" & txtbkName.text & "' group by INVOICEDATE,SUBLEDGER,agentname"
        
     If debitForAgn <> "" Then
        
        con.Execute "insert into tmptbl(INVOICEDATE,SUBLEDGER,RepName,NETAMOUNT,Type_,QUANTITY)" & _
        " select INVOICEDATE,SUBLEDGER,RepName,sum(Amount),'CR',0 from PartyCreditRegisterNew where " & debitForAgn & " and saletype='" & txtbkName.text & "' group by INVOICEDATE,SUBLEDGER,RepName"

        
        con.Execute "insert into tmptbl(INVOICEDATE,SUBLEDGER,RepName,NETAMOUNT,Type_,QUANTITY)" & _
        " SELECT debitRegister.invoicedate,debitRegister.PGLD,debitRegister.Agentname,sum(debitRegister.amount),'DR',0 FROM debitRegister LEFT OUTER JOIN dbo.debitNotDet ON debitRegister.DNN = debitNotDet.DNN " & _
        " Where (debitNotDet.RepName Is Null and " & debitForAgn & " and saletype='" & txtbkName.text & "') group by INVOICEDATE,PGLD,agentname"
        
        con.Execute "insert into tmptbl(INVOICEDATE,SUBLEDGER,RepName,NETAMOUNT,Type_,QUANTITY)" & _
        " SELECT debitRegister.invoicedate,debitRegister.PGLD,debitNotDet.repname,sum(debitNotDet.Amount),'DR',0 FROM debitRegister LEFT OUTER JOIN dbo.debitNotDet ON debitRegister.DNN = debitNotDet.DNN " & _
        " Where (debitNotDet.RepName Is not Null and " & debitForAgn & " and saletype='" & txtbkName.text & "') group by INVOICEDATE,PGLD,debitNotDet.repname"

        
     End If
       
     con.Execute "insert into tmptbl(INVOICEDATE,SUBLEDGER,RepName,NETAMOUNT,Type_,QUANTITY)" & _
     "select INVOICEDATE,SUBLEDGER,agentname,sum(NETAMOUNT),'Sale',sum(QUANTITY) from invoiceBQry where GROUPCODE='" & txtbkName.text & "' group by INVOICEDATE,SUBLEDGER,agentname"
        
     con.Execute "insert into tmptbl(INVOICEDATE,SUBLEDGER,RepName,NETAMOUNT,Type_,QUANTITY)" & _
     "select INVOICEDATE,SUBLEDGER,agentname,sum(NETAMOUNT),'Sale',sum(QUANTITY) from cashBQry where GROUPCODE='" & txtbkName.text & "' group by INVOICEDATE,SUBLEDGER,agentname"

    
    End If
     
     
     saleQty = 0
     saleRetQty = 0
        
        
     If RS.State = 1 Then RS.close
     RS.Open "select SUBLEDGER,RepName from  tmptbl where " & str_date & " group by SUBLEDGER,RepName"
     
     While RS.EOF = False
        netsale = 0
        netamt_ret_ = 0
        dr_ = 0
        cr_ = 0
        
        saleQty = 0
        saleRetQty = 0
        
        
        str11 = "SELECt sum(NetAmount),sum(QUANTITY) FROM tmptbl where " & str_date & " and SUBLEDGER='" & RS!subledger & "' and RepName='" & RS!RepName & "' and type_='Ret'"
        If rs_10.State = 1 Then rs_10.close
        rs_10.Open str11, con, adOpenKeyset, adLockOptimistic
        If Not IsNull(rs_10(0)) Then
            netamt_ret_ = Round(rs_10(0), 0)
            saleRetQty = rs_10(1)
        End If
        
        v1 = 0
        
        str11 = "SELECt sum(NetAmount) FROM tmptbl where " & str_date & " and SUBLEDGER='" & RS!subledger & "' and RepName='" & RS!RepName & "' and type_='CR'"
        If rs_10.State = 1 Then rs_10.close
        rs_10.Open str11, con, adOpenKeyset, adLockOptimistic
        If Not IsNull(rs_10(0)) Then
            cr_ = Round(rs_10(0), 0)
            v1 = cr_
        End If
        
        str11 = "SELECt sum(NetAmount) FROM tmptbl where " & str_date & " and SUBLEDGER='" & RS!subledger & "' and RepName='" & RS!RepName & "' and type_='DR'"
        If rs_10.State = 1 Then rs_10.close
        rs_10.Open str11, con, adOpenKeyset, adLockOptimistic
        If Not IsNull(rs_10(0)) Then
            dr_ = Round(rs_10(0), 0)
            v1 = cr_ - dr_
        End If
        
        
        
        
        
        str11 = "SELECt sum(NetAmount),sum(QUANTITY) FROM tmptbl where " & str_date & " and SUBLEDGER='" & RS!subledger & "' and RepName='" & RS!RepName & "' and type_='Sale'"
        If rs_10.State = 1 Then rs_10.close
        rs_10.Open str11, con, adOpenKeyset, adLockOptimistic
        If Not IsNull(rs_10(0)) Then
            netsale = Round(rs_10(0), 0)
            saleQty = rs_10(1)
            
            sum2 = (netsale - (netamt_ret_ + v1))
        Else
            'sum2 = (netamt_ret + v1) * -1
            sum2 = (netsale - (netamt_ret_ + v1))
            
        End If
        
        If (netsale > 0 Or netamt_ret_ > 0 Or v1 > 0) Then
            DoEvents
            vs.rows = vs.rows + 1
            vs.TextMatrix(T1, 0) = T1
            vs.TextMatrix(T1, 1) = Mid(RS!subledger, 1, 6) & "  " & Mid(RS!subledger, 6)
            vs.TextMatrix(T1, 2) = RS!RepName
            rs_.MoveFirst
            rs_.Find "subledger='" & RS!subledger & "'"
            If rs_.EOF = False Then
               vs.TextMatrix(T1, 3) = rs_!Area
               vs.TextMatrix(T1, 4) = rs_!states
            End If
            
            vs.TextMatrix(T1, 5) = Round(netsale, 0)
            vs.TextMatrix(T1, 6) = netamt_ret_
            vs.TextMatrix(T1, 7) = Round(v1, 0)
            vs.TextMatrix(T1, 8) = sum2       '(Round(rs_10(3), 0) - (netamt_ret + Val(vs.TextMatrix(T1, 6))))
            
            vs.TextMatrix(T1, 9) = saleQty
            vs.TextMatrix(T1, 10) = saleRetQty
            vs.TextMatrix(T1, 11) = (saleQty - saleRetQty)
            
            T1 = T1 + 1
            DoEvents
            DoEvents
            
            
            
        End If
        
        
        RS.MoveNext
     Wend
     
    
'    txtTotal.Visible = False
'    txtTot1.Visible = False
'    txtTot2.Visible = False
'    txtTot3.Visible = False
'
    
    
    txtTotal = 0
    txtTot1 = 0
    txtTot2 = 0
    txtTot3 = 0
    
    Dim col5, col6, col7, col8, col9, col10
    col5 = 0
    col6 = 0
    col7 = 0
    col8 = 0
    col9 = 0
    col10 = 0
    col11 = 0
    
    For I = 1 To vs.rows - 1
    If vs.TextMatrix(I, 0) <> "" Then
       
       col5 = col5 + Val(vs.TextMatrix(I, 5))
       col6 = col6 + Val(vs.TextMatrix(I, 6))
       col7 = col7 + Val(vs.TextMatrix(I, 7))
       col8 = col8 + Val(vs.TextMatrix(I, 8))
       col9 = col9 + Val(vs.TextMatrix(I, 9))
       col10 = col10 + Val(vs.TextMatrix(I, 10))
       col11 = col11 + Val(vs.TextMatrix(I, 11))
       
       'SendKeys "{down}"
    
    End If
    Next
    
    vs.SetFocus
    For I = 1 To vs.rows - 1
      sendkeys "{down}"
    Next
    
    vs.rows = vs.rows + 1
    
    vs.TextMatrix(I, 4) = "Total"
    vs.TextMatrix(I, 5) = col5
    vs.TextMatrix(I, 6) = col6
    vs.TextMatrix(I, 7) = col7
    vs.TextMatrix(I, 8) = col8
    vs.TextMatrix(I, 9) = col9
    vs.TextMatrix(I, 10) = col10
    vs.TextMatrix(I, 11) = col11
    
    For k1 = 0 To 11
        vs.Cell(flexcpBackColor, I, k1) = vbGreen
        DoEvents
    Next
        
    rtype = "PartyWiseArea"
    cmdPrint_7.Enabled = True
    
    vs.ColWidth(0) = 600
    vs.ColWidth(1) = 2800
    vs.ColWidth(2) = 1800
    vs.ColWidth(3) = 1800
    vs.ColWidth(4) = 1200
    vs.ColWidth(5) = 900
    vs.ColWidth(6) = 900
    vs.ColWidth(7) = 900
    vs.ColWidth(8) = 900
    vs.ColWidth(9) = 900
    vs.ColWidth(10) = 900
    vs.ColWidth(11) = 900
    
    debitForAgn = ""
    
    txtAlign "partywisearea"
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

Case "Stock In Hand As On"
     cmdPrint_7.Enabled = True
     Screen.MousePointer = vbHourglass
     con.Execute "exec BookStockSummary 'ALL','M'"
     setvsStock
     fillStockGrid
     Screen.MousePointer = vbDefault
     Exit Sub
     
Case "Party Wise & Rep.wise Net Sale.."
     
     Screen.MousePointer = vbHourglass
     
 
     
     vs.Cols = 6
      
     vs.TextMatrix(0, 0) = "SNo."
     vs.TextMatrix(0, 1) = "Code"
     vs.TextMatrix(0, 2) = "Customer Name"
     vs.TextMatrix(0, 3) = "Area"
     vs.TextMatrix(0, 4) = "Rep. Name"
     vs.TextMatrix(0, 5) = "Net Amt."

     
     
     If RS.State = 1 Then RS.close
     If s = "" Then
        If txtbkName = "" Then
           str11 = "SELECT Party,District,AgentName,sum(NETAMOUNT),subledger FROM InvoiceBCashbQry where " & str_date & "  group by Party,AgentName,District,subledger"
           'str11 = "SELECT Party,District,AgentName,sum(NETAMOUNT),subledger FROM invoiceBQry where " & str_date & "  group by Party,AgentName,District,subledger"
        Else
            str11 = "SELECT Party,District,AgentName,sum(NETAMOUNT),subledger FROM InvoiceBCashbQry where groupcode='" & txtbkName.text & "'  and " & str_date & "  group by Party,AgentName,District,subledger"
           'str11 = "SELECT Party,District,AgentName,sum(NETAMOUNT),subledger FROM invoiceBQry where groupcode='" & txtbkname.Text & "'  and " & str_date & "  group by Party,AgentName,District,subledger"
        End If
     Else
         If txtbkName = "" Then
           str11 = "SELECT Party,District,AgentName,sum(NETAMOUNT),subledger FROM InvoiceBCashbQry where " & s & " and " & str_date & "  group by Party,AgentName,District,subledger"
           'str11 = "SELECT Party,District,AgentName,sum(NETAMOUNT),subledger FROM invoiceBQry where " & s & " and " & str_date & "  group by Party,AgentName,District,subledger"
        Else
           str11 = "SELECT Party,District,AgentName,sum(NETAMOUNT),subledger FROM InvoiceBCashbQry where groupcode='" & txtbkName.text & "' and " & s & " and " & str_date & "  group by Party,AgentName,District,subledger"
           'str11 = "SELECT Party,District,AgentName,sum(NETAMOUNT),subledger FROM invoiceBQry where groupcode='" & txtbkname.Text & "' and " & s & " and " & str_date & "  group by Party,AgentName,District,subledger"
        End If
     End If
     
     RS.Open str11, con, adOpenKeyset, adLockOptimistic
     For J = 1 To RS.RecordCount
        
        vs.TextMatrix(J, 0) = J
        vs.TextMatrix(J, 1) = Left(RS!subledger, 5)
        vs.TextMatrix(J, 2) = RS(0)
        vs.TextMatrix(J, 3) = RS(1)
        vs.TextMatrix(J, 4) = RS(2)
        vs.TextMatrix(J, 5) = RS(3)
        
        RS.MoveNext
     Next
     
     
     con.Execute "delete from TmpBook"
     
     For I = 1 To vs.rows - 1
     If vs.TextMatrix(I, 0) <> "" Then
      con.Execute "insert into TmpBook(BCode,BName,city,area,BalanceQty,head) values('" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "','" & vs.TextMatrix(I, 3) & "','" & vs.TextMatrix(I, 4) & "','" & Round(vs.TextMatrix(I, 5), 1) & "','" & UId & "')"
     End If
     Next
     
    txtTotal = 0
    For I = 1 To vs.rows - 1
    If vs.TextMatrix(I, 0) <> "" Then
    txtTotal = txtTotal + Val(vs.TextMatrix(I, 5))
    End If
    Next
    
    rtype = "Party Wise & Rep.wise"
    cmdPrint_7.Enabled = True
    'setvs
    
    
    vs.Cols = 6
    
    
    vs.ColWidth(0) = 850
    vs.ColWidth(1) = 1250
    vs.ColWidth(2) = 3000
    vs.ColWidth(3) = 2000
    vs.ColWidth(4) = 2600
    vs.ColWidth(5) = 1400
    
    txtAlign "partywiserep"
    
    Screen.MousePointer = vbDefault
    Exit Sub
     
End Select

setvs

Screen.MousePointer = vbDefault
cmdPrint_7.Enabled = True

End Sub
Private Sub cmdExit_12_Click()
Unload Me
End Sub
Sub setvsStock()
    vs.Cols = 8
    
    vs.FormatString = "BookCode|Book Name|>Stock in G1|>Stock in G2|>Stock@Printer|>Net For Sale|>Order in Hand|>Net Qty For Sale|>Qty Sold"
    
    
    vs.ColWidth(0) = 850
    vs.ColWidth(1) = 3750
    vs.ColWidth(2) = 1400
    vs.ColWidth(3) = 1400
    vs.ColWidth(4) = 1400
    vs.ColWidth(5) = 1400
    vs.ColWidth(6) = 1400
    vs.ColWidth(7) = 1400
    
End Sub
Sub setvs()
    vs.ColWidth(0) = 1100
    vs.ColWidth(1) = 4400
    vs.ColWidth(2) = 2100
    vs.ColWidth(2) = 2000
    
End Sub
Private Sub cmdPrint_7_Click()
    
Dim str_date As String


DSNNew
    
    Dim s As String
    s = ""
    For I = 0 To cmbAgentName.ListCount - 1
    If cmbAgentName.Selected(I) = True Then
        If s = "" Then
           s = "{ISSUEBOOK.AGENTNAME}='" & cmbAgentName.List(I) & "'"
        Else
           s = s & " Or " & "{ISSUEBOOK.AGENTNAME}='" & cmbAgentName.List(I) & "'"
        End If
    End If
    Next
 
 
 
If (cboType.text = "State Wise") Then
    rtype = "State Wise"
End If
 
 
If rtype = "billwise" Then
  
  
  If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/BillWise_RepWise.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    If txtbkName.text = "" Then
       MainMenu.cr1.ReplaceSelectionFormula "({invoiceaQry.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {invoiceaQry.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "'))"
    Else
       MainMenu.cr1.ReplaceSelectionFormula "({invoiceaQry.groupcode}='" & txtbkName.text & "' and {invoiceaQry.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {invoiceaQry.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "'))"
    End If
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.Action = 1
  End If
  
'ElseIf rtype = "Rep.Wise & Title Wise Net Qty Summary.." Then
'
'    MainMenu.cr1.Reset
'    MainMenu.cr1.ReportFileName = rptPath & "/RepWiseSaleReturn.rpt"
'    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user  & ";pwd=" & sql_pass
'    If txtbkName.Text = "" Then
'       MainMenu.cr1.ReplaceSelectionFormula "({RepWiseSaleReturnQry.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {RepWiseSaleReturnQry.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "'))"
'    Else
'       MainMenu.cr1.ReplaceSelectionFormula "({RepWiseSaleReturnQry.groupcode}='" & txtbkName.Text & "' and {RepWiseSaleReturnQry.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {RepWiseSaleReturnQry.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "'))"
'    End If
'    MainMenu.cr1.WindowShowPrintSetupBtn = True
'    MainMenu.cr1.WindowShowExportBtn = True
'    MainMenu.cr1.WindowShowRefreshBtn = True
'    MainMenu.cr1.WindowState = crptMaximized
'    MainMenu.cr1.Action = 1
ElseIf rtype = "CreditRegister" Then

    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/ReceiptRegister.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    If txtParty <> "" Then
       MainMenu.cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.states}='" & txtParty & "'"
    End If
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1

  
ElseIf rtype = "grosssale_Statewise" Then
    
' If txtbkName.Text = "" Then
    
    Dim sst As String
    sst = ""
    
    If txtParty <> "" Then
       sst = "{PartyWiseItemWiseQty.states}='" & txtParty & "'"
    End If
    
    If txtbkName.text <> "" Then
       If sst = "" Then
          sst = "{PartyWiseItemWiseQty.groupcode}='" & txtbkName.text & "'"
       Else
          sst = sst & " and {PartyWiseItemWiseQty.groupcode}='" & txtbkName.text & "'"
       End If
    End If
    
    
    
    
    
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/StateWiseItemWiseSalesDet.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    'If txtParty <> "" Then
    '   MainMenu.Cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.states}='" & txtParty & "' and {PartyWiseItemWiseQty.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
    'Else
       MainMenu.cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "') and " & sst
    'End If
    
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1
    
 
    
' End If

ElseIf cboType.text = "Party Wise & Book Wise Sale & Return" Then

   Dim bb1 As Boolean
   
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/PartyWiseItemWiseSalesReturn.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    bb1 = False
    If Check1_godwon.value = 0 Then
        If txtParty <> "" Then
        
           If txtbkName = "" Then
           MainMenu.cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.SUBLEDGER}='" & txtParty1 & "' and {PartyWiseItemWiseQty.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
           Else
           MainMenu.cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.GROUPCODE}='" & txtbkName.text & "' and {PartyWiseItemWiseQty.SUBLEDGER}='" & txtParty1 & "' and {PartyWiseItemWiseQty.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
           End If
           
           bb1 = True
        End If
    Else
           If txtbkName = "" Then
           MainMenu.cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.godown}='" & txtParty & "' and {PartyWiseItemWiseQty.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
           Else
           MainMenu.cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.GROUPCODE}='" & txtbkName.text & "' and {PartyWiseItemWiseQty.godown}='" & txtParty & "' and {PartyWiseItemWiseQty.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
           End If
           bb1 = True
    End If
    
    If bb1 = False Then
       If txtbkName = "" Then
       MainMenu.cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
       Else
       MainMenu.cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.GROUPCODE}='" & txtbkName.text & "' and {PartyWiseItemWiseQty.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
       End If
       
    End If
    
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1

    Exit Sub

  
ElseIf rtype = "grosssale" Then

   Dim bb2 As Boolean
   
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/PartyWiseItemWiseSalesDet.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    bb2 = False
    
    If Check1_godwon.value = 0 Then
        If txtParty <> "" Then
        
           If txtbkName = "" Then
           MainMenu.cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.SUBLEDGER}='" & txtParty1 & "' and {PartyWiseItemWiseQty.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
           Else
           MainMenu.cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.GROUPCODE}='" & txtbkName.text & "' and {PartyWiseItemWiseQty.SUBLEDGER}='" & txtParty1 & "' and {PartyWiseItemWiseQty.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
           End If
           
           'bb2 = True
           MainMenu.cr1.WindowShowPrintSetupBtn = True
           MainMenu.cr1.WindowShowExportBtn = True
           MainMenu.cr1.WindowState = crptMaximized
           MainMenu.cr1.WindowShowRefreshBtn = True
           MainMenu.cr1.Action = 1
           Exit Sub
           
        End If
    Else
           If txtbkName = "" Then
           MainMenu.cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.godown}='" & txtParty & "' and {PartyWiseItemWiseQty.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
           Else
           MainMenu.cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.GROUPCODE}='" & txtbkName.text & "' and {PartyWiseItemWiseQty.godown}='" & txtParty & "' and {PartyWiseItemWiseQty.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
           End If
           
           'bb2 = True
           MainMenu.cr1.WindowShowPrintSetupBtn = True
           MainMenu.cr1.WindowShowExportBtn = True
           MainMenu.cr1.WindowState = crptMaximized
           MainMenu.cr1.WindowShowRefreshBtn = True
           MainMenu.cr1.Action = 1
           Exit Sub
           
           
    End If
    
    If bb1 = False Then
       If txtbkName = "" Then
       MainMenu.cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
       Else
       MainMenu.cr1.ReplaceSelectionFormula "{PartyWiseItemWiseQty.GROUPCODE}='" & txtbkName.text & "' and {PartyWiseItemWiseQty.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {PartyWiseItemWiseQty.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
       End If
       
        MainMenu.cr1.WindowShowPrintSetupBtn = True
        MainMenu.cr1.WindowShowExportBtn = True
        MainMenu.cr1.WindowState = crptMaximized
        MainMenu.cr1.WindowShowRefreshBtn = True
        MainMenu.cr1.Action = 1
       
    End If
    

ElseIf rtype = "SchoolWiseSale" Then
    
    
    Screen.MousePointer = vbHourglass
    If txtParty <> "" Then
       'con.Execute "exec SchoolList_sp '" & txtParty & "'"
       con.Execute "exec SchoolList_sp '" & txtParty & "' , '" & txtFrom & "','" & dateAson.value & "'"
    End If
    Screen.MousePointer = vbDefault
    
    DoEvents
    DoEvents
    DoEvents
    
    
    ss = ""
    
    If txtParty <> "" Then
       ss = "{tmpSchoolWiseBkWiseNSale.states}='" & txtParty & "'"
    End If
    
    If txtbkName <> "" Then
       If ss = "" Then
       ss = "{mpSchoolWiseBkWiseNSale.agentname}='" & txtbkName & "'"
       Else
       ss = ss & "  and  " & "{tmpSchoolWiseBkWiseNSale.agentname}='" & txtbkName & "'"
       End If
    End If
    
    
    If txtgp <> "" Then
       If ss = "" Then
       ss = "{mpSchoolWiseBkWiseNSale.gpname}='" & txtgp & "'"
       Else
       ss = ss & "  and  " & "{tmpSchoolWiseBkWiseNSale.gpname}='" & txtgp & "'"
       End If
    End If
    
    
    
    
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/SchoolWiseSale.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    
    If ss <> "" Then
       MainMenu.cr1.ReplaceSelectionFormula "" & ss
       'Else
       'MainMenu.cr1.ReplaceSelectionFormula "{tmpSchoolWiseBkWiseNSale.states}='" & txtParty & "' and {tmpSchoolWiseBkWiseNSale.agentname}='" & txtbkname & "'"
    End If
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1
    
ElseIf rtype = "Rep. Wise & Book & Bill Wise Sale..." Then

   
    DoEvents
    DoEvents
    DoEvents
    stt = ""
    If (txtParty <> "") Then
    stt = "{invoicebQry.agentname}='" & txtParty & "'"
    End If
    
    If (txtbkName <> "") Then
       If stt = "" Then
          stt = "{invoicebQry.bookcode}='" & txtbkName & "'"
       Else
          stt = stt & " and {invoicebQry.bookcode}='" & txtbkName & "'"
       End If
    End If
    
    If (txtbk <> "") Then
       If stt = "" Then
          stt = "{invoicebQry.states}='" & txtbk.text & "'"
       Else
          stt = stt & " and {invoicebQry.states}='" & txtbk.text & "'"
       End If
    End If
    
    
    
    
       If stt = "" Then
          stt = "{invoicebQry.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {invoicebQry.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
       Else
          stt = stt & " and {invoicebQry.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {invoicebQry.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
       End If
    
        
      MainMenu.cr1.Reset
      MainMenu.cr1.ReportFileName = rptPath & "/BillWise_RepWisebookwise.rpt"
      MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
      MainMenu.cr1.ReplaceSelectionFormula stt
        
       MainMenu.cr1.WindowShowPrintSetupBtn = True
        MainMenu.cr1.WindowShowExportBtn = True
        MainMenu.cr1.WindowState = crptMaximized
        MainMenu.cr1.WindowShowRefreshBtn = True
        MainMenu.cr1.Action = 1
    

ElseIf rtype = "Book Wise & School Wise Net Sale.." Then
    
   
    
     
    Screen.MousePointer = vbHourglass
    'If txtParty <> "" Then
       con.Execute "exec SchoolList_sp '" & txtParty & "' , '" & txtFrom & "','" & dateAson.value & "'"
    'End If
    Screen.MousePointer = vbDefault
    
    DoEvents
    DoEvents
    DoEvents
    
   stt = ""
   If (txtParty <> "") Then
    stt = "{tmpSchoolWiseBkWiseNSale.states}='" & txtParty & "'"
   End If
    
    If (txtbkName <> "") Then
       If stt = "" Then
          stt = "{tmpSchoolWiseBkWiseNSale.sername}='" & txtParty & "'"
       Else
          stt = stt & " and {tmpSchoolWiseBkWiseNSale.sername}='" & txtbkName & "'"
       End If
       
    End If
    
    If (txtbk <> "") Then
       If stt = "" Then
          stt = "{tmpSchoolWiseBkWiseNSale.bookcode}='" & txt & "'"
       Else
          stt = stt & " and {tmpSchoolWiseBkWiseNSale.bookcode}='" & txtbk & "'"
       End If
    End If
    
       If (txtgp <> "") Then
       If stt = "" Then
          stt = "{tmpSchoolWiseBkWiseNSale.gpname}='" & txtgp & "'"
       Else
          stt = stt & " and {tmpSchoolWiseBkWiseNSale.gpname}='" & txtgp & "'"
       End If
       End If
 
        
      MainMenu.cr1.Reset
      MainMenu.cr1.ReportFileName = rptPath & "/BKWise_SchoolWiseSale.rpt"
      MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
      MainMenu.cr1.ReplaceSelectionFormula stt
        

        
        MainMenu.cr1.WindowShowPrintSetupBtn = True
        MainMenu.cr1.WindowShowExportBtn = True
        MainMenu.cr1.WindowState = crptMaximized
        MainMenu.cr1.WindowShowRefreshBtn = True
        MainMenu.cr1.Action = 1
    

ElseIf rtype = "Party Wise & Rep.wise" Then

  If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/PartyWiseRepWiseAmttobeCollect.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1
  End If


   
ElseIf (rtype = "bookwise" Or rtype = "bookwiseQty") Then

    If s = "" Then              'check_selectRep = False Then
       MsgBox "Select at least 1 Representative... ", vbCritical
       Exit Sub
    End If
    
    's = s & " and {ISSUEBOOK.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {ISSUEBOOK.invoicedate}<=datevalue('" & Format(dateAsOn.value, "MM/dd/yyyy") & "')"
    s = "{ISSUEBOOK.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {ISSUEBOOK.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
    
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/BookLadgerSale.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    If s <> "" Then
     MainMenu.cr1.ReplaceSelectionFormula s
    End If
    MainMenu.cr1.Formulas(0) = "fdate='" & Me.txtFrom.value & "'"
    MainMenu.cr1.Formulas(1) = "tdate='" & Me.dateAson.value & "'"

    MainMenu.cr1.WindowShowPrintBtn = True
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowSearchBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1
ElseIf rtype = "bookwiseret" Then
    
    If s = "" Then              'check_selectRep = False Then
       MsgBox "Select at least 1 Representative... ", vbCritical
       Exit Sub
    End If
    
    s = s & " and {ISSUEBOOK.invoicedate}>=datevalue('" & Format(txtFrom.value, "MM/dd/yyyy") & "') and {ISSUEBOOK.invoicedate}<=datevalue('" & Format(dateAson.value, "MM/dd/yyyy") & "')"
    
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/BookLadgerSaleRet.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    If s <> "" Then
     MainMenu.cr1.ReplaceSelectionFormula s
    End If
    MainMenu.cr1.Formulas(0) = "fdate='" & Me.txtFrom.value & "'"
    MainMenu.cr1.Formulas(1) = "tdate='" & Me.dateAson.value & "'"

    MainMenu.cr1.WindowShowPrintBtn = True
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowSearchBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1

ElseIf cboType.text = "Party Wise Area & Rep. Wise Net Sale..." Then

   con.Execute "delete from TmpBook where head='" & UId & "'"
   For I = 1 To vs.rows - 1
   If vs.TextMatrix(I, 0) <> "" Then
     con.Execute "insert into TmpBook(BName,repname,City,states,Qty,issueQty,orderno,BalanceQty,head,sqty,srqty) " & _
     " values('" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "','" & vs.TextMatrix(I, 3) & "','" & vs.TextMatrix(I, 4) & "','" & vs.TextMatrix(I, 5) & "','" & vs.TextMatrix(I, 6) & "','" & vs.TextMatrix(I, 7) & "','" & Round(vs.TextMatrix(I, 8), 1) & "','" & UId & "','" & vs.TextMatrix(I, 9) & "','" & vs.TextMatrix(I, 10) & "')"
   End If
   Next
   
   If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/NetSales_repwise.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    MainMenu.cr1.ReplaceSelectionFormula "{TmpBook.head}='" & UId & "'"
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1
    End If
cmdPrint_7.Enabled = False


ElseIf rtype = "PartyWiseArea" Then
   
   con.Execute "delete from TmpBook where head='" & UId & "'"
   For I = 1 To vs.rows - 1
   If vs.TextMatrix(I, 0) <> "" Then
     con.Execute "insert into TmpBook(BName,City,states,Qty,issueQty,orderno,BalanceQty,head,GrossAmt) " & _
     " values('" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "','" & vs.TextMatrix(I, 3) & "','" & vs.TextMatrix(I, 4) & "','" & vs.TextMatrix(I, 5) & "','" & vs.TextMatrix(I, 6) & "','" & Round(vs.TextMatrix(I, 7), 1) & "','" & UId & "','" & Round(vs.TextMatrix(I, 8), 1) & "')"
   End If
   Next
   
   If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/NetSales.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    MainMenu.cr1.ReplaceSelectionFormula "{TmpBook.head}='" & UId & "'"
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1
    End If
cmdPrint_7.Enabled = False

ElseIf rtype = "District Wise" Then
   
   con.Execute "delete from TmpBook where head='" & UId & "'"
   con.Execute "delete from TmpBook where head is null"
   
   For I = 1 To vs.rows - 1
   If vs.TextMatrix(I, 0) <> "" Then
     con.Execute "insert into TmpBook(BName,BalanceQty,head,issueQty,states) values('" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 3) & "','" & UId & "','" & vs.TextMatrix(I, 4) & "','" & vs.TextMatrix(I, 2) & "')"
   End If
   Next
   
   If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/NetSales_dis.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    MainMenu.cr1.ReplaceSelectionFormula "{TmpBook.head}='" & UId & "'"
    MainMenu.cr1.Formulas(0) = "fdate='" & Me.txtFrom.value & "'"
    MainMenu.cr1.Formulas(1) = "tdate='" & Me.dateAson.value & "'"

    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1
   End If
ElseIf rtype = "Party Wise Area Wise Net Quantity Sale..." Then
    
    con.Execute "exec PartyWiseItemWiseSale_Ret '" & txtFrom.value & "','" & dateAson.value & "'"
    
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/PartyWiseItemWiseQty.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    If txtbkName.text <> "" Then
       MainMenu.cr1.ReplaceSelectionFormula "{partywiseItemwiseQty.groupcode}='" & txtbkName.text & "'"
       'MainMenu.cr1.Formulas(0) = "gp_='" & txtbkName.Text & "'"
    Else
       'MainMenu.cr1.Formulas(0) = "gp_='" & "" & "'"
    End If
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1
    
ElseIf rtype = "State Wise" Then

   con.Execute "delete from TmpBook where head='" & UId & "'"
   con.Execute "delete from TmpBook where head is null"
   
   For I = 1 To vs.rows - 1
   If vs.TextMatrix(I, 0) <> "" Then
     con.Execute "insert into TmpBook(BName,BalanceQty,head,issueQty,sqty,srqty) values('" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "','" & UId & "','" & vs.TextMatrix(I, 3) & "','" & vs.TextMatrix(I, 5) & "','" & vs.TextMatrix(I, 6) & "')"
   End If
   Next
   
   If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/StateWiseNetSales.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    MainMenu.cr1.ReplaceSelectionFormula "{TmpBook.head}='" & UId & "'"
    MainMenu.cr1.Formulas(0) = "fdate='" & Me.txtFrom.value & "'"
    MainMenu.cr1.Formulas(1) = "tdate='" & Me.dateAson.value & "'"

    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1
   End If

ElseIf rtype = "Rep Wise" Then
    
   con.Execute "delete from TmpBook where head='" & UId & "'"
   con.Execute "delete from TmpBook where head is null"
   
   For I = 1 To vs.rows - 1
   If vs.TextMatrix(I, 0) <> "" Then
     con.Execute "insert into TmpBook(BName,BalanceQty,head,issueQty,area,city,OrderNo,GrossAmt) values('" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "','" & UId & "','" & vs.TextMatrix(I, 3) & "','" & vs.TextMatrix(I, 4) & "','" & vs.TextMatrix(I, 5) & "','" & vs.TextMatrix(I, 7) & "','" & vs.TextMatrix(I, 6) & "')"
   End If
   Next
   
   If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/NetSales_rep.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    MainMenu.cr1.ReplaceSelectionFormula "{TmpBook.head}='" & UId & "'"
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.Formulas(0) = "fdate='" & Me.txtFrom.value & "'"
    MainMenu.cr1.Formulas(1) = "tdate='" & Me.dateAson.value & "'"
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1
   End If
   
ElseIf rtype = "Book Wise Sales" Then
    
   con.Execute "delete from TmpBook where head='" & UId & "'"
   con.Execute "delete from TmpBook where head is null"
   
   For I = 1 To vs.rows - 1
   If vs.TextMatrix(I, 0) <> "" Then
     con.Execute "insert into TmpBook(BCode,BName,BalanceQty,head,Qty,issueQty,OrderNo) values('" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "','" & Round(vs.TextMatrix(I, 3), 3) & "','" & UId & "','" & vs.TextMatrix(I, 4) & "','" & vs.TextMatrix(I, 6) & "','" & vs.TextMatrix(I, 5) & "')"
   End If
   Next
   
   If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/NetSales_BkQty.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    MainMenu.cr1.ReplaceSelectionFormula "{TmpBook.head}='" & UId & "'"
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.Action = 1
   End If
   
  ElseIf rtype = "Book Wise Sales(Area Wise)" Then
    
   Screen.MousePointer = vbHourglass
   
   con.Execute "delete from TmpBook where head='" & UId & "'"
   con.Execute "delete from TmpBook where head is null"
   
   For I = 1 To vs.rows - 1
   If vs.TextMatrix(I, 0) <> "" Then
     con.Execute "insert into TmpBook(BCode,BName,BalanceQty,head,Qty,issueQty,OrderNo,area) values('" & vs.TextMatrix(I, 2) & "','" & vs.TextMatrix(I, 3) & "','" & Round(vs.TextMatrix(I, 4), 3) & "','" & UId & "','" & vs.TextMatrix(I, 5) & "','" & vs.TextMatrix(I, 7) & "','" & vs.TextMatrix(I, 6) & "','" & vs.TextMatrix(I, 1) & "')"
   End If
   Next
   
   If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/NetSales_BkQtyArea.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    MainMenu.cr1.ReplaceSelectionFormula "{TmpBook.head}='" & UId & "'"
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.Action = 1
   End If
 
   Screen.MousePointer = vbDefault
   
   
  ElseIf rtype = "Tital Wise & Party Wise Sale & Sale Ret. Qty" Then
   
   con.Execute "delete from TmpBook where head='" & UId & "'"
   For I = 1 To vs.rows - 1
   If vs.TextMatrix(I, 0) <> "" Then
     con.Execute "insert into TmpBook(bcode,BName,repname,City,states,Qty,issueQty,orderno,head) " & _
     " values('" & vs.TextMatrix(I, 0) & "','" & vs.TextMatrix(I, 1) & "','" & vs.TextMatrix(I, 2) & "','" & vs.TextMatrix(I, 3) & "','" & vs.TextMatrix(I, 4) & "','" & vs.TextMatrix(I, 5) & "','" & vs.TextMatrix(I, 6) & "','" & vs.TextMatrix(I, 7) & "','" & UId & "')"
   End If
   Next

   If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/TitalWisePartyWiseSale.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    MainMenu.cr1.ReplaceSelectionFormula "{TmpBook.head}='" & UId & "'"
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1
    End If
    cmdPrint_7.Enabled = False

   
   
  ElseIf rtype = "Party Payment Details" Then
  
   con.Execute "delete from tmpPayment"
   con.Execute "insert into tmpPayment(VoucherType , VoucherDate, Genledger, SubLedger, amount, DebitorCredit,DESCRIPTION,crno,PAYTYPE) select VoucherType , VoucherDate, Genledger, SubLedger, amount, DebitorCredit,DESCRIPTION,vsno,PAYTYPE from VOUCHERS where (GenLedger='SUNDRY DEBTORS' and  VoucherDate>=convert(smalldatetime, '" & txtFrom.value & "'  ,103) and VoucherDate<=convert(smalldatetime, '" & dateAson.value & "'  ,103) and SUBLEDGER not like '%IMPREST A/C%')"

   
   
   If MsgBox("Want To Print ... ?", vbInformation + vbYesNo) = vbYes Then
        MainMenu.cr1.Reset
        MainMenu.cr1.ReportFileName = rptPath & "/paymentdet.rpt"
        MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
        If txtParty.text <> "" Then
        MainMenu.cr1.ReplaceSelectionFormula "{sledger.states}='" & txtParty.text & "'"
        End If
        MainMenu.cr1.Formulas(0) = "fdate='" & Me.txtFrom.value & "'"
        MainMenu.cr1.Formulas(1) = "tdate='" & Me.dateAson.value & "'"
        MainMenu.cr1.WindowShowPrintSetupBtn = True
        MainMenu.cr1.WindowShowExportBtn = True
        MainMenu.cr1.WindowShowRefreshBtn = True
        MainMenu.cr1.WindowState = crptMaximized
        MainMenu.cr1.Action = 1
   End If
  

   
Else
   
   
   con.Execute "delete from TmpBook where head='" & UId & "'"
   con.Execute "delete from TmpBook where head is null"
   
   For I = 1 To vs.rows - 1
   If vs.TextMatrix(I, 0) <> "" Then
     con.Execute "insert into TmpBook(BName,BalanceQty,head) values('" & vs.TextMatrix(I, 1) & "','" & Round(vs.TextMatrix(I, 2), 2) & "','" & UId & "')"
   End If
   Next
   
   If MsgBox("Want To Print ?", vbInformation + vbYesNo) = vbYes Then
    MainMenu.cr1.Reset
    MainMenu.cr1.ReportFileName = rptPath & "/NetSales.rpt"
    MainMenu.cr1.Connect = "filedsn=chitradsn;uid= " & sql_user & ";pwd=" & sql_pass
    MainMenu.cr1.ReplaceSelectionFormula "{TmpBook.head}='" & UId & "'"
    MainMenu.cr1.WindowShowPrintSetupBtn = True
    MainMenu.cr1.WindowShowExportBtn = True
    MainMenu.cr1.WindowState = crptMaximized
    MainMenu.cr1.WindowShowRefreshBtn = True
    MainMenu.cr1.Action = 1
End If

cmdPrint_7.Enabled = False

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
Dim Q1, q2, J, Q1_sp As Double

Dim b1 As Boolean

b1 = False


c = 1
r = 1


'As On ==================================================================================

If rtype = "Stock In Hand As On" Then


    xl.Columns("A:H").ColumnWidth = 12
    J = 2
    xlSheet.Cells(1, 1).value = "Stock In Hand As On : " & dateAson.value
    
    For I = 1 To vs.rows - 1
        For J = 1 To vs.Cols - 1
            xlSheet.Cells(, 1).value = "dinesh"
            xlSheet.Cells(r, J).value = "saini"
        Next
    Next
    
    
    Screen.MousePointer = vbDefault
    Exit Sub



'========

ElseIf rtype = "Consolidated Sales Summary Rep. Wise ..." Then

    row_ = 1
    col_ = 1
    xl.Columns("A:H").ColumnWidth = 12
    J = 2
    xlSheet.Cells(1, 1).value = "Bilty Return Status "
    
    For I = 0 To vs.rows - 1
        For J = 0 To vs.Cols - 1
                   xlSheet.Cells(row_, col_).value = vs.TextMatrix(I, J)
               
               col_ = col_ + 1
               
        Next
        
        return_ = 0
        sale_ = 0
        
        
        If row_ > 1 Then
        
            xlSheet.Cells(1, 8).value = "Net Sale"
            xlSheet.Cells(1, 9).value = "Return (%)"
            
            sale_ = (Val(vs.TextMatrix(I, 2)) + Val(vs.TextMatrix(I, 3)))
            return_ = (Val(vs.TextMatrix(I, 4)) + Val(vs.TextMatrix(I, 5)) + Val(vs.TextMatrix(I, 6)))
            
            xlSheet.Cells(row_, 8).value = (Val(vs.TextMatrix(I, 2)) + Val(vs.TextMatrix(I, 3)) - return_)
            
            If Val(return_) > 0 Then
            If (sale_ - return_) > 0 Then
                 xlSheet.Cells(row_, 9).value = Round((return_ / sale_ * 100), 2)
             End If
                 
            End If
            
        End If
        
        row_ = row_ + 1
        
        col_ = 1
    Next
    
    Screen.MousePointer = vbDefault
    
    MsgBox ("Complete..."), vbInformation
    
    Exit Sub

ElseIf rtype = "Consolidated Sales Summary..." Then


    row_ = 1
    col_ = 1
    xl.Columns("A:H").ColumnWidth = 12
    J = 2
    xlSheet.Cells(1, 1).value = "Bilty Return Status "
    
    For I = 0 To vs.rows - 1
        For J = 0 To vs.Cols - 1
                   xlSheet.Cells(row_, col_).value = vs.TextMatrix(I, J)
               
               col_ = col_ + 1
               
        Next
        
        return_ = 0
        sale_ = 0
        
        
        If row_ > 1 Then
        
            xlSheet.Cells(1, 11).value = "Net Sale"
            xlSheet.Cells(1, 12).value = "Return (%)"
            
            sale_ = (Val(vs.TextMatrix(I, 5)) + Val(vs.TextMatrix(I, 6)))
            return_ = (Val(vs.TextMatrix(I, 7)) + Val(vs.TextMatrix(I, 8)) + Val(vs.TextMatrix(I, 9)))
            
            xlSheet.Cells(row_, 11).value = (Val(vs.TextMatrix(I, 5)) + Val(vs.TextMatrix(I, 6)) - return_)
            
            If Val(return_) > 0 Then
            If (sale_ - return_) > 0 Then
                 xlSheet.Cells(row_, 12).value = Round((return_ / sale_ * 100), 2)
             End If
                 
            End If
            
        End If
        
        row_ = row_ + 1
        
        col_ = 1
    Next
    
    Screen.MousePointer = vbDefault
    MsgBox ("Complete..."), vbInformation
    Exit Sub


ElseIf rtype = "Bilty Return Status..." Then
     
    row_ = 1
    col_ = 1
   
    xl.Columns("A:H").ColumnWidth = 12
    J = 2
    xlSheet.Cells(1, 1).value = "Bilty Return Status "
    
    For I = 0 To vs.rows - 1
        For J = 0 To vs.Cols - 1
            If (col_ = 2 Or col_3) Then
               xlSheet.Cells(row_, col_).value = Format(vs.TextMatrix(I, J), "dd/MM/yyyy")
            Else
               xlSheet.Cells(row_, col_).value = vs.TextMatrix(I, J)
            End If
            col_ = col_ + 1
        Next
        row_ = row_ + 1
        col_ = 1
    Next
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
ElseIf (cboType.text = "Book Wise (Sale & Sale Return Bill List)..." Or cboType.text = "Tital Wise & Party Wise Sale & Sale Ret. Qty" Or cboType.text = "Party Payment Details") Then

    row_ = 1
    col_ = 1
   
    xl.Columns("A:H").ColumnWidth = 12
    J = 2
    xlSheet.Cells(1, 1).value = "Bilty Return Status "
    
    For I = 0 To vs.rows - 1
        For J = 0 To vs.Cols - 1
               xlSheet.Cells(row_, col_).value = vs.TextMatrix(I, J)
              col_ = col_ + 1
        Next
        row_ = row_ + 1
        col_ = 1
    Next
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    

End If
'========================================================================================
   
    If rtype = "bookwiseQty" Then
    
        Dim dt1, dt2, dt1_r, dt2_r, dt1_sp, dt2_sp As String
        
        dt1 = Format(txtFrom.value, "MM/dd/yyyy")
        dt2 = Format(dateAson.value, "MM/dd/yyyy")
        
        dt1_r = Format(txtFrom.value, "MM/dd/yyyy")
        dt2_r = Format(dateAson.value, "MM/dd/yyyy")
        
        dt1_sp = "01/08/" & Mid(last_dbase, 12, 2)
    
        dt2_sp = "31/07/" + Mid(last_dbase, 14, 2)
        
        dt1_sp = Format(dt1_sp, "MM/dd/yyyy")
        dt2_sp = Format(dt2_sp, "MM/dd/yyyy")

    
        con.Execute ("exec Sp_tmpSaleSpRegister '','" & txtbkName.text & "','" & last_dbase & "','" & dt1 & "','" & dt2 & "','" & dt1_r & "','" & dt2_r & "','" & dt1_sp & "','" & dt2_sp & "'")
 
    End If
    

    xl.Columns("A:H").ColumnWidth = 12
    J = 3
    xlSheet.Cells(1, 1).value = "Representative Wise Sales As On : " & dateAson.value
    
    For I = 0 To cmbAgentName.ListCount - 1
      
        If cmbAgentName.Selected(I) = True Then
           r = 1
           xlSheet.Cells(r, J).value = cmbAgentName.List(I)
        
        ''Raws fill==========================================================
        Q1 = 0
        q2 = 0
        
        If RS.State = 1 Then RS.close
        ''RS.Open "SELECT BOOKCODE,BOOKNAME FROM BOOKS where bookcode='ACM2' group by BOOKCODE,BOOKNAME", con
        If txtbkName = "" Then
        RS.Open "SELECT BOOKCODE,BOOKNAME FROM BookQry_ group by BOOKCODE,BOOKNAME", con
        Else
        RS.Open "SELECT BOOKCODE,BOOKNAME FROM BookQry_ where groupcode='" & txtbkName.text & "' group by BOOKCODE,BOOKNAME", con
        End If
        
        While RS.EOF = False
           
           Q1 = 0
           q2 = 0
           Q1_sp = 0
           
           
           If rtype = "bookwiseQty" Then
            
            If rs_1.State = 1 Then rs_1.close
            rs_1.Open "select BOOKCODE,agentname,BOOKNAME,(sum(quantity_sale)-sum(quantity_return)) as netsale FROM tmpSaleAndSp " & _
            " where GROUPCODE='" & txtbkName.text & "' and agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "'  group by BOOKCODE,agentname,BOOKNAME", con


            If (rs_1.RecordCount > 0) Then
                If Not IsNull(rs_1(3)) Then
                   Q1 = Round(rs_1(3), 0)
                End If
            End If
            
              If rs_1.State = 1 Then rs_1.close
            rs_1.Open "select BOOKCODE,agentname,BOOKNAME,(sum(quantity_sp)-sum(quantity_spreturn)) as netsale FROM tmpSaleAndSp " & _
            " where GROUPCODE='" & txtbkName.text & "' and agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "'  group by BOOKCODE,agentname,BOOKNAME", con


            If (rs_1.RecordCount > 0) Then
                If Not IsNull(rs_1(3)) Then
                   Q1_sp = Round(rs_1(3), 0)
                End If
            End If
         
            

            
''            If rs_1.State = 1 Then rs_1.close
''            rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([QUANTITY]) as qty FROM Billwise_CashBQryinvoiceBQry " & _
''            " where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "' and (invoicedate>=convert(smalldatetime,'" & txtFrom.value & "',103) and invoicedate<=convert(smalldatetime,'" & dateAsOn.value & "',103))  group by BOOKCODE,agentname,BOOKNAME", con
''            If (rs_1.RecordCount > 0) Then
''            If Not IsNull(rs_1(3)) Then
''               Q1 = rs_1(3)
''            End If
''            End If
''
''            If rs_1.State = 1 Then rs_1.close
''            If Check1_saleret.value = 0 Then
''                rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([QUANTITY]) as qty FROM creditBQry " & _
''                " where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "' and (invoicedate>=convert(smalldatetime,'" & txtFrom.value & "',103) and invoicedate<=convert(smalldatetime,'" & dateAsOn.value & "',103))  group by BOOKCODE,agentname,BOOKNAME", con
''            Else
''                rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([QUANTITY]) as qty FROM creditBQry " & _
''                " where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "' and (invoicedate>=convert(smalldatetime,'" & txtFrom.value & "',103) and invoicedate<=convert(smalldatetime,'" & txtSaleRetDate.value & "',103))  group by BOOKCODE,agentname,BOOKNAME", con
''
''            End If
''
''
''            If (rs_1.RecordCount > 0) Then
''            If Not IsNull(rs_1(3)) Then
''               Q1 = Q1 - rs_1(3)
''            End If
''            End If
            
            
           ElseIf rtype = "bookwise" Then
            
            If rs_1.State = 1 Then rs_1.close
   
            
             
                        rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([NETAMOUNT]) as qty FROM Billwise_CashBQryinvoiceBQry " & _
                        " where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "' and (invoicedate>=convert(smalldatetime,'" & txtFrom.value & "',103) and invoicedate<=convert(smalldatetime,'" & dateAson.value & "',103))  group by BOOKCODE,agentname,BOOKNAME", con
            
            
                        If (rs_1.RecordCount > 0) Then
                            If Not IsNull(rs_1(3)) Then
                               Q1 = Round(rs_1(3), 0)
                            End If
                        End If
            
                        If rs_1.State = 1 Then rs_1.close
                        If Check1_saleret.value = 0 Then
                            rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([NETAMOUNT]) as qty FROM creditBQry " & _
                            " where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "' and (invoicedate>=convert(smalldatetime,'" & txtFrom.value & "',103) and invoicedate<=convert(smalldatetime,'" & dateAson.value & "',103))  group by BOOKCODE,agentname,BOOKNAME", con
                        Else
                            rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([NETAMOUNT]) as qty FROM creditBQry " & _
                            " where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "' and (invoicedate>=convert(smalldatetime,'" & txtFrom.value & "',103) and invoicedate<=convert(smalldatetime,'" & txtSaleRetDate.value & "',103))  group by BOOKCODE,agentname,BOOKNAME", con
            
                        End If
            
                        If (rs_1.RecordCount > 0) Then
                            If Not IsNull(rs_1(3)) Then
                               Q1 = Q1 - Round(rs_1(3), 2)
                            End If
                        End If
            
            
           ElseIf rtype = "bookwiseret" Then
            
            If rs_1.State = 1 Then rs_1.close
            rs_1.Open "SELECT BOOKCODE,agentname,BOOKNAME,sum([NETAMOUNT]) as qty FROM creditBQry " & _
            " where agentname='" & cmbAgentName.List(I) & "' and BOOKCODE='" & RS!Bookcode & "' and (invoicedate>=convert(smalldatetime,'" & txtFrom.value & "',103) and invoicedate<=convert(smalldatetime,'" & dateAson.value & "',103))  group by BOOKCODE,agentname,BOOKNAME", con
            If (rs_1.RecordCount > 0) Then
                If Not IsNull(rs_1(3)) Then
                   Q1 = Q1 - Round(rs_1(3), 2)
                End If
            End If
            
            
           
           End If
           
           
           
           r = r + 1
           
           If b1 = False Then
              xlSheet.Cells(r, 1).value = RS!Bookcode
              xlSheet.Cells(r, 2).value = RS!Bookname
           End If
           
           xlSheet.Cells(r, J).value = Q1
           
           If rtype = "bookwiseQty" Then
               xlSheet.Cells(r, J + 1).value = Q1_sp
           End If
           
           
          RS.MoveNext
        
        Wend
        
        '====================================================================
        b1 = True
        J = J + 2
        End If
        
       
    Next





Screen.MousePointer = vbDefault


Exit Sub
Screen.MousePointer = vbDefault
err:
MsgBox err.DESCRIPTION




End Sub

Private Sub Form_Load()
   
   
   Me.Width = 9420
   Me.Height = 9300
   cboType.ListIndex = 0
   
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
    RS.close
   
   
   
   
If RS.State = 1 Then RS.close
RS.Open "select yarfrom,yarto from setup1 where " & stringyear & "", con
If RS.EOF = False Then
 
 txtFrom.value = Format(RS!yarfrom, "dd/MM/yyyy")
 dateAson.value = Format(RS!yarto, "dd/MM/yyyy")
 txtSaleRetDate.value = Format(RS!yarto, "dd/MM/yyyy")
End If
   
End Sub

Private Sub txtBilty_Dist_GotFocus()
 If PopUpValue1 <> "" Then
    txtBilty_Dist.text = PopUpValue1
    PopUpValue1 = ""
    
 End If

End Sub

Private Sub txtBilty_Dist_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    
    searchType = "party"
    value = "select distinct(DISTCODE) as District from SLEDGER where gledger='SUNDRY DEBTORS' and " & stringyear & " "
    popuplist_client value, con
    
 End If

End Sub

Private Sub txtBilty_Party_GotFocus()
 
 If PopUpValue1 <> "" Then
    txtBilty_Party.text = PopUpValue3
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    
 End If
 
End Sub

Private Sub txtBilty_Party_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    
    searchType = "party"
    value = "select distinct(Party),Code,subledger from SLEDGER where gledger='SUNDRY DEBTORS' and " & stringyear & "  order by party"
    popuplist_client value, con
    
 End If

End Sub

Private Sub txtBilty_State_GotFocus()
 If PopUpValue1 <> "" Then
    txtBilty_State.text = PopUpValue1
    PopUpValue1 = ""
    
 End If

End Sub

Private Sub txtBilty_State_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
    
    searchType = "party"
    value = "select distinct(states) as State from SLEDGER where gledger='SUNDRY DEBTORS' and " & stringyear & " "
    popuplist_client value, con
    
 End If

End Sub


Private Sub txtbk_GotFocus()
If PopUpValue1 <> "" Then

  If (cboType = "Book Wise & School Wise Net Sale.." Or cboType = "Rep. Wise & Book & Bill Wise Sale...") Then
     txtbk = PopUpValue1
     'txtbkName = ""
 End If
 PopUpValue1 = ""
 PopUpValue2 = ""
 
End If

End Sub

Private Sub txtbk_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then

   
If (cboType = "Book Wise & School Wise Net Sale.." Or cboType = "") Then

     
   value = "select BOOKCODE,BOOKNAME,SerName from BOOKS order by BOOKCODE"
   popuplist_client value, con
   set_focus = True

ElseIf cboType.text = "Rep. Wise & Book & Bill Wise Sale..." Then
   
      searchType = "party"
    value = "select distinct(states) as State from SLEDGER where gledger='SUNDRY DEBTORS' and " & stringyear & " "
    popuplist_client value, con

   
End If

End If
End Sub

Private Sub txtbkName_GotFocus()
If PopUpValue1 <> "" Then

  If cboType = "Book Wise & School Wise Net Sale.." Then
     txtbkName = PopUpValue1
  ElseIf (cboType = "Rep. Wise & Book & Bill Wise Sale...") Then
    txtbkName = PopUpValue1
  Else
     txtbkName = PopUpValue1
  End If
   
   PopUpValue1 = ""
   PopUpValue3 = ""
   cmdPrint_7.Enabled = True
   
End If


End Sub

Private Sub txtbkName_KeyDown(KeyCode As Integer, Shift As Integer)
   
If KeyCode = 113 Then

   
If cboType = "School Wise & Book Wise Net Sale.." Then
    
    searchType = "party"
    value = "select Rep as Representative from SalesRepQry order by Rep"
    popuplist_client value, CON_blue
    set_focus = True

ElseIf (cboType = "Rep. Wise & Book & Bill Wise Sale..." Or cboType = "Tital Wise & Party Wise Sale & Sale Ret. Qty") Then

   searchType = "party"
   value = "select BOOKCODE,BOOKNAME,SerName from BOOKS order by BOOKCODE"
   popuplist_client value, con
   set_focus = True
   
ElseIf (rtype = "Book Wise Sales" Or cboType.text = "State Wise & Book Wise Gross Sales.." Or rtype = "Book Wise Sales(Area Wise)" Or cboType.text = "District Wise" Or cboType.text = "Party Wise & Book Wise Gross Sales" Or rtype = "Consolidated Sales Summary..." Or rtype = "Consolidated Sales Summary Rep. Wise ..." Or cboType.text = "Party Wise & Rep.wise Net Sale.." Or cboType = "Rep Wise" Or rtype = "Party Wise Area Wise Net Quantity Sale..." Or cboType.text = "State Wise" Or cboType = "Representative & Book Wise Net Sales(Amt.)" Or cboType = "Representative & Book Wise Net Sales(Qty.)" Or rtype = "Party Wise Area & Rep. Wise Net Sale..." Or rtype = "PartyWiseArea" Or rtype = "billwise" Or cboType = "Representative & Book Wise Sales Return" Or rtype = "Rep.Wise & Title Wise Net Qty Summary..") Then

   searchType = "party"
   value = "select distinct iif(GROUPCODE_sub is null,GROUPCODE,GROUPCODE_sub) as GROUPCODE from BOOKS where (freeItem='n') order by GROUPCODE"
   
   popuplist_client value, con
   set_focus = True

   
Else

    searchType = "party"
    value = "select distinct SerName from BOOKS where len(SerName)>0"
    popuplist_client value, con
   set_focus = True

   
End If

End If
 

End Sub

Private Sub txtGP_GotFocus()
If PopUpValue1 <> "" Then

     txtgp.text = PopUpValue1
     PopUpValue1 = ""
End If
End Sub

Private Sub txtGP_KeyDown(KeyCode As Integer, Shift As Integer)
   
If KeyCode = 113 Then
   searchType = "party"
   value = "select distinct iif(GROUPCODE_sub is null,GROUPCODE,GROUPCODE_sub) as GROUPCODE from BOOKS where freeitem='n' order by GROUPCODE"
   popuplist_client value, con
   set_focus = True
End If

End Sub

Private Sub txtParty_GotFocus()

If cboType.text = "Rep. Wise & Book & Bill Wise Sale..." Then
   txtParty.text = PopUpValue1
   PopUpValue1 = ""
   Exit Sub
End If


If PopUpValue1 <> "" Then
   txtParty = PopUpValue1
   txtParty1 = PopUpValue3
   
   PopUpValue1 = ""
   PopUpValue3 = ""
   cmdPrint_7.Enabled = True
   
End If

End Sub
Private Sub txtParty_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then

If (cboType.text = "Party Wise & Book Wise Gross Sales" Or cboType.text = "Party Wise & Book Wise Sale & Return") Then
    
    
If Check1_godwon.value = 0 Then

    searchType = "party"
    lblCr = "dr"
    value = "select distinct(Party),Code,subledger from SLEDGER where gledger='SUNDRY DEBTORS' and " & stringyear & "  order by party"
    popuplist_client value, con
    set_focus = True
    
Else
    
    'If rs_godwn.State = 1 Then rs_godwn.close
    'rs_godwn.Open "select * from GodownMaster where len(Godwn)<=3 and " & stringyear & " order by id", CON, adOpenForwardOnly, adLockReadOnly

    value = "select Godwn from GodownMaster where len(Godwn)<=3 order by id"
    popuplist_client value, con
    set_focus = True


    
End If
    
ElseIf cboType.text = "Rep. Wise & Book & Bill Wise Sale..." Then
       searchType = "party"
        value = "select Rep as Representative from SalesRepQry order by Rep"
        popuplist_client value, CON_blue
        set_focus = True
Else

    searchType = "party"
    value = "select States from SLEDGER where gledger='SUNDRY DEBTORS' and " & stringyear & " group by States"
    popuplist_client value, con
    set_focus = True
    
End If



End If

End Sub
Private Sub txtRep_GotFocus()

If PopUpValue1 <> "" Then
    txtRep.text = PopUpValue1
    PopUpValue1 = ""
 End If
 
End Sub

Private Sub txtRep_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 113 Then
    
    searchType = "party"
    value = "select Rep as Representative from SalesRepQry where (email is not null and len(email)>1) order by Representative"
    popuplist_client value, CON_blue
    
 End If
 
 
 
End Sub

Private Sub vs_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    
    If cboType = "Party Payment Details" Then
    
    If vs.Col = 8 Then
       con.Execute "update VOUCHERS set PayType='" & LCase(vs.TextMatrix(vs.RowSel, 8)) & "' where vsno=" & vs.TextMatrix(vs.RowSel, 0) & ""
       sendkeys "{down}"
    End If
    
    End If
    
 End If
 


End Sub
