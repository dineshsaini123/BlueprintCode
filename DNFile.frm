VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Debitnotefile 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9132
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   14472
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9523.138
   ScaleMode       =   0  'User
   ScaleWidth      =   14475.76
   Begin Crystal.CrystalReport CR 
      Left            =   10755
      Top             =   6975
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame panel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Debit Note"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   8844
      Left            =   36
      TabIndex        =   13
      Top             =   90
      Width           =   14388
      Begin VB.TextBox txtchecked 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   2460
         MaxLength       =   100
         TabIndex        =   44
         Top             =   7776
         Width           =   540
      End
      Begin VB.CommandButton cmdListBlankOrd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Empty No"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   432
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   7740
         Width           =   1056
      End
      Begin VB.ListBox List_emptyList 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5064
         Left            =   432
         TabIndex        =   42
         Top             =   2664
         Visible         =   0   'False
         Width           =   1488
      End
      Begin VB.CheckBox Check_header 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print With Header"
         Height          =   195
         Left            =   7050
         TabIndex        =   41
         Top             =   7872
         Width           =   1770
      End
      Begin VB.ComboBox cmbAgentName 
         Height          =   315
         Left            =   8190
         TabIndex        =   8
         Top             =   1170
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtAmtwords 
         Height          =   285
         Left            =   1452
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   8280
         Visible         =   0   'False
         Width           =   7800
      End
      Begin VB.ComboBox cmbgroup 
         Height          =   1488
         Left            =   7245
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   35
         Top             =   4620
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.PictureBox pic1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   840
         Left            =   405
         ScaleHeight     =   840
         ScaleWidth      =   8832
         TabIndex        =   27
         Top             =   6804
         Width           =   8835
         Begin VB.CommandButton cmdNHPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "N&HPrint"
            Height          =   555
            Left            =   6960
            Picture         =   "DNFile.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   840
            Width           =   165
         End
         Begin VB.CommandButton Commandadd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   705
            Left            =   45
            Picture         =   "DNFile.frx":0BE4
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   30
            Width           =   1035
         End
         Begin VB.CommandButton CommandReturn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Return"
            Height          =   705
            Left            =   7695
            Picture         =   "DNFile.frx":17C8
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   45
            Width           =   1035
         End
         Begin VB.CommandButton CommandPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   705
            Left            =   6597
            Picture         =   "DNFile.frx":23AC
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   60
            Width           =   1035
         End
         Begin VB.CommandButton Commandsearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Search"
            Height          =   705
            Left            =   5505
            Picture         =   "DNFile.frx":2F90
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   30
            Width           =   1035
         End
         Begin VB.CommandButton Commanddelete 
            BackColor       =   &H00FFFFFF&
            Caption         =   "De&lete"
            Height          =   705
            Left            =   4413
            Picture         =   "DNFile.frx":3B74
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   30
            Width           =   1035
         End
         Begin VB.CommandButton Commandedit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            Height          =   705
            Left            =   1137
            Picture         =   "DNFile.frx":4758
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   30
            Width           =   1035
         End
         Begin VB.CommandButton Commandabandon 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aba&ndon"
            Height          =   705
            Left            =   3321
            Picture         =   "DNFile.frx":4B9A
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   30
            Width           =   1035
         End
         Begin VB.CommandButton Commandsave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sa&ve"
            Enabled         =   0   'False
            Height          =   705
            Left            =   2229
            Picture         =   "DNFile.frx":5124
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   30
            Width           =   1035
         End
      End
      Begin VB.TextBox TCNN 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1650
         TabIndex        =   1
         Top             =   465
         Width           =   930
      End
      Begin VB.TextBox GText 
         Height          =   285
         Left            =   6585
         TabIndex        =   20
         Top             =   795
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.TextBox SText 
         Height          =   285
         Left            =   6615
         TabIndex        =   19
         Top             =   1170
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7248
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   6264
         Width           =   1425
      End
      Begin VB.ComboBox Subcombo1 
         Height          =   1488
         Left            =   2835
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   17
         Top             =   4572
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.ComboBox gencombo1 
         Height          =   1488
         Left            =   390
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   9
         Top             =   4572
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox GCombo 
         Height          =   315
         Left            =   1650
         TabIndex        =   3
         Top             =   795
         Width           =   4965
      End
      Begin VB.ComboBox Scombo 
         Height          =   315
         Left            =   1650
         TabIndex        =   4
         Top             =   1140
         Width           =   4965
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1620
         MaxLength       =   150
         TabIndex        =   6
         Top             =   4080
         Width           =   10836
      End
      Begin VB.TextBox CdateT 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   8208
         TabIndex        =   16
         Top             =   792
         Visible         =   0   'False
         Width           =   1308
      End
      Begin VB.TextBox DrCrText 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   270
         Left            =   6300
         MaxLength       =   1
         TabIndex        =   15
         Text            =   "C"
         Top             =   4776
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox AText 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   270
         Left            =   5355
         TabIndex        =   14
         Top             =   4776
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSMask.MaskEdBox Cbdate 
         Height          =   315
         Left            =   3285
         TabIndex        =   2
         Top             =   450
         Width           =   1095
         _ExtentX        =   1926
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid DGrid 
         Height          =   1668
         Left            =   132
         TabIndex        =   7
         Top             =   4500
         Width           =   11112
         _ExtentX        =   19600
         _ExtentY        =   2942
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   10813439
         AllowUserResizing=   2
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VSFlex7Ctl.VSFlexGrid VS1 
         Height          =   2136
         Left            =   144
         TabIndex        =   5
         Top             =   1800
         Width           =   13980
         _cx             =   24659
         _cy             =   3768
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483645
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
         Rows            =   6
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   330
         RowHeightMax    =   0
         ColWidthMin     =   0
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
            Left            =   14640
            MultiLine       =   -1  'True
            TabIndex        =   46
            Top             =   6360
            Width           =   195
         End
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Checked :"
         Height          =   252
         Left            =   1692
         TabIndex        =   45
         Top             =   7812
         Width           =   684
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks for Letter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   384
         Index           =   3
         Left            =   108
         TabIndex        =   40
         Top             =   1536
         Width           =   1776
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rep. Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   9585
         TabIndex        =   39
         Top             =   1260
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Amt in words :"
         Height          =   252
         Left            =   432
         TabIndex        =   38
         Top             =   8316
         Visible         =   0   'False
         Width           =   1032
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   948
         Left            =   360
         Top             =   6756
         Width           =   8916
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Index           =   1
         Left            =   6228
         TabIndex        =   26
         Top             =   6324
         Width           =   1008
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Genral Ledger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   25
         Top             =   795
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Ledger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   24
         Top             =   1185
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Narration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Index           =   0
         Left            =   96
         TabIndex        =   23
         Top             =   4044
         Width           =   792
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2790
         TabIndex        =   22
         Top             =   495
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit Note No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   21
         Top             =   495
         Width           =   1425
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   360
      TabIndex        =   10
      Top             =   8040
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   60
         Width           =   765
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Width           =   855
      End
   End
End
Attribute VB_Name = "Debitnotefile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RS As ADODB.Recordset
Dim mvBookMark As Variant
Dim cmdAdd As Boolean
Dim cmdedit As Boolean
Dim LRC As Integer
Dim LCC As Integer
Dim Glastrow As Integer
Dim Datachange As Boolean
Dim emptyInv_bool As Boolean
Dim Li As Integer
Sub CNbandon()

Text3.text = ""
Me.Commandadd.Enabled = True
Me.Commandedit.Enabled = True
Me.Commandsearch.Enabled = True
Me.Commandsave.Enabled = False
Me.Commanddelete.Enabled = True
Me.Commandabandon.Enabled = True
Me.CommandPrint.Enabled = True


End Sub
Sub Setgrid()

  DGrid.Cols = 5
  DGrid.Row = 0
  DGrid.Col = 0
  DGrid.text = "Gen Ledger"
  DGrid.Col = 1
  DGrid.text = "Sub Ledger"
  DGrid.Col = 2
  DGrid.text = Format(DGrid.text, "0.00")
  DGrid.text = "Amount"
  DGrid.Col = 3
  DGrid.text = "Debit/Credit"
  DGrid.Col = 4
  DGrid.text = "Group"
  
  DGrid.ColWidth(0) = 4000
  DGrid.ColWidth(1) = 3000
  DGrid.ColWidth(2) = 1400
  DGrid.ColWidth(3) = 1400
  DGrid.ColWidth(4) = 1200
  
 
 vs1.Cols = 6
   
 vs1.Clear
 vs1.FormatString = "SN.|Description||Amount|Rep.Name"
 
 vs1.TextMatrix(1, 0) = "1"
 vs1.TextMatrix(2, 0) = "2"
 vs1.TextMatrix(3, 0) = "3"
 vs1.TextMatrix(4, 0) = "4"
 vs1.TextMatrix(5, 0) = "5"
 
 vs1.ColHidden(5) = True
 
 vs1.ColWidth(0) = 500
 vs1.ColWidth(1) = 9600
 vs1.ColWidth(2) = 0
 vs1.ColWidth(3) = 1200
 vs1.ColWidth(4) = 1400


End Sub
Sub Controlclear()
     CdateT.text = ""
     GCombo.text = ""
     Scombo.text = ""
     Text2.text = ""
     Text3.text = ""
End Sub
Sub Gridrefresh()

gencombo1.Visible = False
Subcombo1.Visible = False
AText.Visible = False
DrCrText.Visible = False
cmbgroup.Visible = False
DoEvents


      Dim grs1 As New ADODB.Recordset
      If grs1.State = 1 Then grs1.close
         If TCNN.text = "" Then
            
            grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit,groupcode FROM  DNFB WHERE   " & stringyear & " and DNN = 0", con, adOpenStatic
            Set DGrid.DataSource = grs1
            DGrid.Refresh
            Setgrid
   
         ElseIf cmdedit = True Then
                grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit,groupcode FROM  DNFB where  " & stringyear & "", con, adOpenStatic
                DGrid.rows = 99
                DGrid.TopRow = 1
                Setgrid
                
                If TCNN.text <> "" Then
                For I = 1 To 99 - 1 'grs1.RecordCount - 1
                DGrid.Row = I
                DGrid.Col = 2
                DGrid.text = Format(DGrid.text, "0.00")
                DGrid.Refresh
                Next I
                End If
                
         ElseIf cmdAdd = True Then
                grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit,groupcode,groupcode FROM  DNFB where " & stringyear & "", con, adOpenStatic
                DGrid.Refresh
                DGrid.rows = 99
                DGrid.TopRow = 1
                Setgrid
                
                If TCNN.text <> "" Then
                For I = 1 To 99 - 1   'grs1.RecordCount - 1
                DGrid.Row = I
                DGrid.Col = 2
                DGrid.text = Format(DGrid.text, "0.00")
                DGrid.Refresh
                Next I
                End If
                
         Else
                grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit,groupcode FROM  DNFB WHERE   " & stringyear & " and DNN = " + Trim(TCNN.text) + "", con, adOpenStatic
                Set DGrid.DataSource = grs1
                DGrid.Refresh
                Setgrid
                
                
                If TCNN.text <> "" Then
                For I = 1 To grs1.RecordCount - 1
                    DGrid.Row = I
                    DGrid.Col = 2
                    DGrid.text = Format(DGrid.text, "0.00")
                    DGrid.Refresh
                Next I
                End If
                            
         End If
                '======================
                    If TCNN.text <> "" Then
                    If rs1.State = 1 Then rs1.close
                    rs1.Open "select * from DebitNotDet where Dnn=" & TCNN.text & " order by SN", con, adOpenDynamic, adLockOptimistic
                    For I = 1 To rs1.RecordCount
                        If RS.EOF = False Then
                           vs1.TextMatrix(I, 0) = rs1!sn
                           vs1.TextMatrix(I, 1) = rs1!NARR
                        If Not IsNull(rs1!paymentamt) Then
                            If rs1!paymentamt > 0 Then
                               vs1.TextMatrix(I, 2) = rs1!paymentamt & ""
                            End If
                        End If

                        vs1.TextMatrix(I, 3) = rs1!amount
                        vs1.TextMatrix(I, 4) = rs1!RepName & ""
                        vs1.TextMatrix(I, 5) = rs1!recno & ""
                        rs1.MoveNext
                    End If
                  Next I
                End If
                '======================
         
    
 
       
       
       
     DoEvents
     AText.Visible = False
       DGrid.Col = 2
End Sub
Sub GridEdit()
    Dim grs1 As New ADODB.Recordset
    If grs1.State = 1 Then grs1.close
    If TCNN.text = "" Then
       grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit FROM  TempCNF1B where  " & stringyear & "", con, adOpenStatic
    End If
End Sub
Private Sub AText_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If DGrid.Row >= 1 Then
           DGrid.RemoveItem DGrid.Row
           gencombo1.text = ""
           gencombo1.Visible = False
           AText.text = ""
           AText.Visible = False
           DGrid.SetFocus
       End If
   End If
End If

End Sub
Private Sub AText_KeyPress(KeyAscii As Integer)


If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Or KeyAscii = 8 Or KeyAscii = 13 Then
Else
   KeyAscii = 0
End If


If KeyAscii = 13 Then

    If Val(AText.text) <= 0 Then
          MsgBox "please Enter  Amount.."
          AText.SetFocus
          Exit Sub
     End If
     If AText.text = "" Then
          MsgBox "please Enter Amount.."
          AText.SetFocus
          Exit Sub
     End If

     DGrid.text = Format(AText.text, "0.00")
     If AText = "" Then DGrid.text = 0
         
      DGrid.Col = DGrid.Col + 1
      DGrid_Click
      Glastrow = DGrid.Row
       
End If


End Sub
Private Sub Cbdate_Change()
CdateT.text = Cbdate.text
End Sub
Private Sub Cbdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  sendkeys "{tab}"
End If
End Sub
Private Sub Cbdate_LostFocus()

If Trim(Cbdate.text) = "__/__/____" Then
    MsgBox "Please Enter Date..."
    Cbdate.SetFocus
End If
If Trim(Cbdate.text) <> "__/__/____" Then
    If Not checkdate(Trim(Cbdate.text), Cbdate) Then
        Cbdate.SetFocus
    End If
End If

End Sub
Private Sub CdateT_Change()
 On Error GoTo er1
 Cbdate.text = CdateT.text
er1:  If err.Number = 380 Then
       Exit Sub
      End If
End Sub



Private Sub cmbAgentName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub cmbgroup_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
      Dim rs2 As New ADODB.Recordset
      Dim rs3 As New ADODB.Recordset
      If cmbgroup.text <> "" Then
          Dim rs4 As New ADODB.Recordset
          rs4.Open "Select* from groups where groupcode='" + Trim(cmbgroup.text) + "' and " & stringyear, con, adOpenForwardOnly, adLockReadOnly, adCmdText
          If rs4.RecordCount <= 0 Then
             MsgBox "No valid group"
             cmbgroup.Visible = True
             cmbgroup.SetFocus
             Exit Sub
          End If
     End If
        DGrid.text = cmbgroup.text
        cmbgroup.Visible = False
        Glastrow = DGrid.Row
        DGrid.Row = DGrid.Row + 1
        LRC = LRC + 1
        DGrid.Col = 0
       DGrid_Click
       
        
   End If


End Sub

Private Sub cmdListBlankOrd_Click()

List_emptyList.Enabled = True
List_emptyList.Visible = True
Dim k1 As Integer
Dim b1 As Boolean
b1 = False


If emptyInv_bool = False Then
   List_emptyList.Visible = True
   emptyInv_bool = True
Else
   List_emptyList.Visible = False
   emptyInv_bool = False
End If


List_emptyList.Clear


Set RS = New ADODB.Recordset
Set RS = con.Execute("exec searchList 'DNFA'")

While RS.EOF = False

    If b1 = False Then
       k1 = RS(0)
       b1 = True
    End If
    
    If k1 <> RS(0) Then
       List_emptyList.AddItem (RS(0) - 1)
       b1 = False
    End If

k1 = k1 + 1
RS.MoveNext

Wend


End Sub

Private Sub cmdNHPrint_Click()
  
printch = "Dnfa"
ino = TCNN
printch1 = "DNN"

Printheader = False
GenrateReport

End Sub
Sub GenrateReport()
   Dim rs7 As ADODB.Recordset
   Dim rs1 As ADODB.Recordset
   Dim kk As ADODB.Recordset
   Dim trs As ADODB.Recordset
   Dim paperWidth As Integer
   Dim kkk As ADODB.Recordset
   Dim Tot As Double
   Dim MaxLine, Pno, Line As Integer
   Dim called1 As Boolean
   Dim Glist1 As String
   Dim ID1 As String
   Dim Gc As String
   Dim Gc1 As String
   Dim FooterYes As Boolean
   Dim NetTotal As Double
   Dim GTotal As Double
   Dim J As Integer
   NetTotal = 0
   I = 0
   
   GTotal = 0
   FooterYes = False
   Set kkk = New ADODB.Recordset
   Set rs1 = New ADODB.Recordset
   Set rs7 = New ADODB.Recordset
   Set kk = New ADODB.Recordset
   Set trs = New ADODB.Recordset
   Tot = 0
   Line = 0
   Pno = 1
   MaxLine = 72
   called1 = False
   called2 = False
   main.reportname = "Dis. Sales"
   main.reportdata
   main.repors.Find "reportname='" + Trim(main.reportname) + "'"
   MaxLine = main.repors!totalline
   If main.repors!comp = True Then
      paperWidth = Int(main.repors!totalcolumn * 1.75)
   Else
      paperWidth = main.repors!totalcolumn
   End If
   Open "" + VB.App.Path + "\vipin.txt" For Output As #1
   MaxLine = 72
   called1 = False
   Pno = 1
   paperWidth = 96
header:
   For I = 1 To main.repors!TopMargin
       Print #1, ""
       Line = Line + 1
   Next
   If FooterYes = True Then
       While Line < 72
           Print #1, ""
           Line = Line + 1
       Wend
       Line = 0
       FooterYes = False
   End If
   If kkk.State = 1 Then kkk.close
   CNSetup
   kkk.Open "select * from setup1 where " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
   If Printheader = True Then
     
   If Not kkk.BOF Then
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, Chr(27) + Chr(71); Chr(27) + Chr(77) + Chr(14)
     Print #1, Tab(((paperWidth - (Len(Trim(kkk!cname)) * 2)) / 2)); Chr(27) + Chr(77) + Chr(14); Trim(kkk!cname)
     Print #1, Tab((paperWidth - (Len(Trim(kkk!add1)) * 2)) / 2); Chr(27) + Chr(77); dspace(Trim(kkk!add1))
     Print #1, Tab((paperWidth - (Len(Trim(kkk!phone1)) * 2)) / 2); Trim(kkk!phone1) & "," & Trim(kkk!phone2)
     Line = Line + 8
   End If
Else
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, ""
     Print #1, Chr(27) + Chr(77)
     Line = Line + 8
End If
   Print #1, Chr(27) + Chr(71); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("DEBIT NOTE")))) / 2 - 3); Chr(14); "DEBIT NOTE"; Chr(20)
   Line = Line + 1
   If Printheader = True Then
      Print #1, Tab(63); kkk!uptt
      Print #1, Tab(63); kkk!cst
      Line = Line + 2
   End If
   If Printheader = False Then
      Print #1, ""
      Print #1, ""
      Line = Line + 2
   End If
   Print #1, repli("-", paperWidth)
   Print #1, ""
   Line = Line + 2
   If called1 = True Then
        called1 = False
        GoTo printagain1
    End If
'convert(smalldatetime,'" + Trim(Cbdate.Text) + "',103)", CON
If rs7.State = 1 Then rs7.close
rs7.Open "Select * from DNFA where Dnn =" & TCNN.text & "  and Dnd = convert(smalldatetime,'" + Trim(Cbdate.text) + "',103) and " & stringyear, con
If rs7.RecordCount > 0 Then
   Print #1, Chr(27) + Chr(71); "To,   S.L. Code : "; Tab(20); Mid$(rs7!psld, 1, 5); Tab(50); "Debit Note No. : "; Chr(27) + Chr(72); Trim(rs7!dnn); Tab(83); Chr(27) + Chr(71); "Date : "; Chr(27) + Chr(72); rs7!dnd
   Line = Line + 1
   If kkk.State = 1 Then kkk.close
   kkk.Open "select * from sledger where subledger='" + Trim(rs7!psld) + "' and " & stringyear, con, adOpenStatic, adLockReadOnly, adCmdText
   If Not kkk.EOF Then
      Print #1, Tab(5); "M/s " & kkk!DESCFORINVOICE
      Print #1, Tab(5); IIf(IsNull(kkk!address1), " ", kkk!address1)
      Print #1, Tab(5); IIf(IsNull(kkk!address2), " ", kkk!address2)
      Print #1, Tab(5); IIf(IsNull(kkk!address3), " ", kkk!address3)
      Print #1, ""
      kkk.close
   End If
   Print #1, ""
   Print #1, Tab(5); "Narration         : "; Tab(30); rs7!n
   Print #1, ""
   Print #1, Tab(0); repli("-", paperWidth)
   Print #1, Tab(5); "GenLedger"; Tab(30); ""; Tab(85); "Amount"
   Print #1, repli("-", paperWidth)
   Line = Line + 11
   If trs.State = 1 Then trs.close
   trs.Open "Select * from DNFB where Dnn =" & TCNN.text & "  and Dnd = convert(smalldatetime,'" + Trim(Cbdate.text) + "',103) and " & stringyear, con
   If trs.RecordCount > 0 Then
      While Not trs.EOF
            Print #1, Tab(5); trs!gld; Tab(35); IIf(IsNull(trs!sld), "", trs!sld); Tab(80); rsets(Trim(Format(Str(trs!a), "0.00")), 12)
            Line = Line + 1
            If Line > MaxLine - 5 Then
                FooterYes = True
                Pno = Pno + 1
                called1 = True
                GoTo header
printagain1:
                called1 = False
            End If
            trs.MoveNext
        Wend
    End If
    While Line <= 58
         Print #1, ""
         Line = Line + 1
    Wend
    
    Print #1, ""
    Print #1, Tab(1); "Net Amount Dr. In Your A/C : "; Tab(80); rsets(Trim(Format(Str(rs7!na), "0.00")), 12)
    Print #1, ""
    Print #1, Tab(1); toword(rs7!na)
    Print #1, repli("-", paperWidth)
    Dim tempdata As ADODB.Recordset
    Set tempdata = New ADODB.Recordset
    CNSetup
    tempdata.Open "select * from setup1 where " & stringyear, con, adOpenStatic, adLockReadOnly
    Print #1, Tab(1); "E.& O.E"
    Print #1, Tab(1); tempdata!COURT; Tab(65); "FOR " + Trim(tempdata!cname)
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Close #1
    
    PrintOption.Show
    
End If



End Sub

Private Sub Command1_Click()
On Error GoTo er

If Not RS.BOF Or Not RS.EOF Then RS.MoveNext
If RS.EOF And RS.RecordCount > 0 Then
    Beep
     
    RS.MoveLast
  End If
  
er:   If err.Number = 3021 Then
         
         Exit Sub
       
    End If
  End Sub

Private Sub Command2_Click()
If Not RS.BOF Then RS.MovePrevious
  If RS.BOF And RS.RecordCount > 0 Then
    Beep
    RS.MoveFirst
  End If
End Sub

Private Sub Command3_Click()
   RS.CancelBatch
  If mvBookMark > 0 Then
    RS.Bookmark = mvBookMark
  Else
    RS.MoveFirst
  End If
  Gridrefresh
  pic1.Visible = True
  Pic2.Visible = False
End Sub

Public Sub Commandabandon_Click()

 On Error Resume Next

  
  
  'If RS.RecordCount > 0 Then
  '   RS.CancelUpdate
  '   RS.MoveFirst
  'End If
  maxdrno
  
  cmdAdd = False
  cmdedit = False
 
  'Gridrefresh
  pic1.Visible = True
  Frame2.Enabled = True
  gencombo1.Visible = False
  Subcombo1.Visible = False
  AText.Visible = False
  DrCrText.Visible = False
  Frame1.Enabled = False
  
  CNbandon
  
  mnuMenu_ = "menudebitnote"
  SetButton Commandadd, Commandedit, Commandsave, Commanddelete
  
End Sub
Sub maxdrno()
     Dim trs As New ADODB.Recordset
     trs.Open "Select max(DNN)as mcnn from DNFA where  " & stringyear & "", con, adOpenStatic, adCmdText
     If trs.RecordCount <= 1 And IsNull(trs!Mcnn) Then
       TCNN.text = 1
     Else
      TCNN.text = trs!Mcnn + 1
    End If
    
    DGrid.rows = 100
    DGrid.Cols = 0
    DGrid.Cols = 5
    Dim I
    For I = 0 To 99
        DGrid.RowHeight(I) = 270
    Next


     Setgrid

End Sub
Private Sub Commandadd_Click()
 On Error Resume Next

 Me.Commandadd.Enabled = False
 Me.Commandedit.Enabled = False
 Me.Commandsearch.Enabled = False
 Me.Commandsave.Enabled = True
 Me.Commanddelete.Enabled = False
 Me.Commandabandon.Enabled = True
 Me.CommandPrint.Enabled = False
 Subcombo1.Visible = False
 Dim rs1  As New ADODB.Recordset
cmdAdd = True
cmdedit = False
LRC = 1
LCC = 0
Frame1.Enabled = True
TCNN.Enabled = True
TCNN.SetFocus

Text3.text = ""
txtchecked.text = ""

'CdateT.Text = ""
'' With RS
'' AddNew
'' End With


If cmdedit = False Then
   maxdrno
End If

''DGrid.Rows = 100
''DGrid.Cols = 0
''DGrid.Cols = 5
''Dim I
''For I = 0 To 99
''        DGrid.RowHeight(I) = 270
''Next
''
''
''Setgrid


DoEvents
GCombo.text = "SUNDRY DEBTORS"
Dim rs2 As New ADODB.Recordset
            rs2.Open "Select * from sledger where  " & stringyear & " and GLEDGER='" + Trim(GCombo.text) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
            If Not rs2.EOF Then
                    Scombo.Clear
                    Do While Not rs2.EOF
                        Scombo.AddItem rs2(1)
                        If Not rs2.EOF Then
                            rs2.MoveNext
                        End If
                    Loop
      
            Else
                    Scombo.Clear
                    Scombo.text = ""
                    Text2.SetFocus
                    Exit Sub
            End If

DGrid.Row = 1
DGrid.Col = 0
DGrid.TopRow = 1
Frame2.Enabled = False
Text2.text = ""

End Sub

Private Sub Commanddelete_Click()


''   If checkAuthentication("DNFA", "dnn", Val(TCNN.Text)) = True Then
''      MsgBox "You are Not Authorised ...", vbCritical
''      Exit Sub
''   End If

Dim rs1 As New ADODB.Recordset
Dim rs_h As New ADODB.Recordset

 If rs1.State = 1 Then rs1.close
    rs1.Open "select top 100 * from DNFA where DNN=" & Trim(TCNN.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
    If rs1.EOF = False Then
       If rs_h.State = 1 Then rs_h.close
       rs_h.Open "select top 100 * from DNFA where Dnn=" & Trim(TCNN.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
           If rs1!bAuthorized = True Then
              MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
              Exit Sub
          End If
    End If


createLog UserName, TCNN, "debit note", " Delete : " & Text3.text, Date


If RS.RecordCount > 0 Then
  If MsgBox("Are you sure.......", vbYesNo) = vbYes Then
  If TCNN.text <> "" Then
  
  
    If (AuditTrail = "y") Then
    
    If (txtchecked.text = "y") Then
    
        actionType_ = "Delete"
        vtype1_ = "D"
        vtypeNew = "D"
        vdate_ = Trim(Cbdate.text)
        vno_ = Trim(TCNN.text)
        
        frmAuditTrailLog_Rem.Show 1
        
     End If
    
    End If

  
  
  
       con.Execute "DELETE FROM  DNFA WHERE   " & stringyear & " and DNN=" + TCNN.text + ""
       con.Execute "DELETE FROM  DNFB WHERE   " & stringyear & " and DNN=" + TCNN.text + ""
       con.Execute "DELETE FROM  DebitNotDet where DNN=" + Trim(TCNN.text) + ""
       Call Commandadd_Click
  End If
  With RS
    '.delete
    On Error Resume Next
     If Not RS.BOF And Not RS.EOF Then RS.MoveFirst
    
     

     
     
  End With
    
  Gridrefresh
  Exit Sub


End If

End If

End Sub
Sub Commandedit_Click()
 
Dim rs1 As New ADODB.Recordset
Dim rs_h As New ADODB.Recordset

 If rs1.State = 1 Then rs1.close
    
  
  rs1.Open "select top 100 * from DNFA where DNN=" & Trim(TCNN.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
  If rs1.EOF = False Then
    If rs_h.State = 1 Then rs_h.close
    rs_h.Open "select top 100 * from DNFA where Dnn=" & Trim(TCNN.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
        If rs1!bAuthorized = True Then
           MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
        Exit Sub
     End If
  End If
  
  
 mnuMenu_ = "menudebitnote"
  
 
 
 If RS.RecordCount > 0 Then
        If RS.RecordCount <= 0 Then
            cmdAdd = False
            cmadd = False
            cmdedit = False
   
            Me.Commandadd.Enabled = True
            Me.Commandedit.Enabled = True
            Me.Commandsearch.Enabled = True
            Me.Commandsave.Enabled = False
            Me.Commanddelete.Enabled = True
            Me.Commandabandon.Enabled = True
            Me.CommandPrint.Enabled = True
            Gridrefresh
            gencombo1.Visible = False
            Subcombo1.Visible = False
            AText.Visible = False
            DrCrText.Visible = False
            Frame2.Enabled = True

          
          Exit Sub
        End If
 
        Me.Commandadd.Enabled = False
        Me.Commandedit.Enabled = False
        Me.Commandsearch.Enabled = False
        Me.Commandsave.Enabled = True
        Me.Commanddelete.Enabled = False
        Me.Commandabandon.Enabled = True
        Me.CommandPrint.Enabled = False
        cmdAdd = False
        cmdedit = True
        TCNN.Enabled = False
        Cbdate.SetFocus

        Gridrefresh
        
        If cmdedit = True Then
          
          
''          Dim rs3 As New ADODB.Recordset
''          rs3.Open "SELECT * FROM  DNFB WHERE   " & stringyear & " and DNN=" + Trim(TCNN.Text) + "", CON, adOpenStatic, adLockOptimistic, adCmdText
''          If rs3.RecordCount >= 0 Then
''             DGrid.Rows = 99
''             DGrid.TopRow = 1
''             DGrid.Row = rs3.RecordCount + 1
''          End If
          
          
    End If
    
    


    
    Frame2.Enabled = False
    VB.Screen.ActiveForm.ActiveControl.SelLength = 2
 End If
 
 
 SetButton Commandadd, Commandedit, Commandsave, Commanddelete
 
End Sub

Private Sub CommandPrint_Click()
GenrateReport
s1 = 4
PrintOption.Show

End Sub

Private Sub CommandReturn_Click()
Unload Me
''''MainMenu.Toolbar1.Visible = True
End Sub
Private Sub Commandsave_Click()
   
   
Dim Checked_YesNo  As Integer

 If (txtchecked.text = "y") Then
      Checked_YesNo = 1
 Else
      Checked_YesNo = 0
 End If
   
Dim rs1 As New ADODB.Recordset
Dim rs_h As New ADODB.Recordset

 If rs1.State = 1 Then rs1.close
    rs1.Open "select top 100 * from DNFA where DNN=" & Trim(TCNN.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
    If rs1.EOF = False Then
       If rs_h.State = 1 Then rs_h.close
       rs_h.Open "select top 100 * from DNFA where Dnn=" & Trim(TCNN.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
           If rs1!bAuthorized = True Then
              MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
              Exit Sub
          End If
    End If

   
   
   createLog UserName, TCNN, "debit note", " Save : " & Text3.text, Date
   
   
   If checkData = True Then
      Exit Sub
   End If
   
   Dim Grs As New ADODB.Recordset
   If TCNN.text = "" Then
       Commandabandon_Click
       Exit Sub
   End If
   
  
   
   con.Execute "DELETE FROM  DNFB WHERE   " & stringyear & " and DNN=" + Trim(TCNN.text) + ""
   TCNN.Enabled = True
   Grs.Open "select * from DNFB where  " & stringyear & "", con, adOpenDynamic, adLockOptimistic, adCmdText
   Dim sum As Double
   sum = 0
   Dim I, J As Integer
   I = 1
   DGrid.Row = 1
   DGrid.Col = 0
   If DGrid.text = "" Then
         RS.CancelUpdate
         Exit Sub
   End If
   While DGrid.text <> ""
        DGrid.Row = I
        Grs.AddNew
        Grs!dnn = Val(TCNN.text)
        Grs!dnd = Cbdate.text
        Grs!groupcode = DGrid.TextMatrix(I, 4)
        
        For J = 0 To 3
          DGrid.Col = J
                If J = 0 Then
                   If DGrid.text = "" Then
                      MsgBox "Please fill Gen Ledger"
                      DGrid_Click
                      Exit Sub
                       
                   Else
                     Grs!gld = DGrid.text
                   End If
               End If
                If J = 1 Then
                    If DGrid.text = "" Then
                       Grs!sld = Null
                    Else
                       Grs!sld = DGrid.text
                    End If
                End If
                If J = 2 Then
                   If DGrid.text = "" Then
                      Grs!a = 0
                   Else
                        Grs!a = DGrid.text
                   End If
               End If
               If J = 3 Then
                   If DGrid.text = "" Then
                         MsgBox "Please fill Correct Entry  "
                         DGrid_Click
                         Exit Sub
                    Else
                         Grs!dc = DGrid.text
                    End If
                    If DGrid.text = "D" Then
                       DGrid.Col = 2
                       sum = sum + Format(Val(DGrid.text), "0.00")
                    Else
                       DGrid.Col = 2
                       sum = sum - Format(Val(DGrid.text), "0.00")
                   End If
             End If
          Next J
          
          Grs!fyear = main.session
          Grs!setupid = main.setupid

          Grs.update
          I = I + 1
          DGrid.Row = I
          DGrid.Col = 0
   Wend
   
   
  '----------------------
   Set rs1 = New ADODB.Recordset
   rs1.Open "select * FROM  DNFA WHERE   " & stringyear & " and DNN=" + Trim(TCNN.text) + "", con, adOpenDynamic, adLockOptimistic
   If rs1.EOF = True Then
      rs1.AddNew
   End If
   
   If (AuditTrail = "y") Then
      rs1!Checked_YesNo = Checked_YesNo
    End If
   
   rs1!dnn = TCNN.text
   rs1!dnd = Cbdate.text
   rs1!Pgld = GText.text
   If SText.text = "" Then
      rs1!psld = Null 'SText.Text
   Else
      rs1!psld = SText.text
   End If
   
   rs1!n = Text2.text
   rs1!na = Abs(Format(sum, "0.00"))
   If sum >= 0 Then
      rs1!dc = "C"
   Else
      rs1!dc = "D"
   End If
   
   rs1!fyear = main.session
   rs1!setupid = main.setupid
   rs1!Amtwords = Trim(txtAmtwords.text)
   
   If cmbAgentName.text <> vs1.TextMatrix(1, 4) Then
      cmbAgentName.text = vs1.TextMatrix(1, 4)
   End If
   rs1!agentname = Trim(cmbAgentName.text)
   
   If InStr(Scombo.text, "(EM)") > 0 Then
      rs1!saletype = "EM"
   Else
      rs1!saletype = "BP"
   End If
   rs1.update
 
''   '----------------------
''   If InStr(Scombo.Text, "(EM)") > 0 Then
''     con.Execute "update DNFA SET saletype='EM' where DNN=" & TCNN.Text & ""
''   End If

   '------------------------------------------------
   
   Set RS = New ADODB.Recordset
   RS.Open "select * from debitNotDet where dnn=" & TCNN.text & "", con, adOpenDynamic, adLockOptimistic
   If RS.EOF = False Then
      con.Execute "delete from debitNotDet where Dnn=" & TCNN.text & ""
   End If
   
   s10 = ""
   For I = 1 To vs1.rows - 1
      If vs1.TextMatrix(I, 1) <> "" Then
       RS.AddNew
       RS!sn = vs1.TextMatrix(I, 0)
       RS!dnn = TCNN.text
       If vs1.TextMatrix(I, 1) = "" Then
          RS!NARR = "-"
       Else
          RS!NARR = vs1.TextMatrix(I, 1)
       End If
       
       If s10 = "" Then
           s10 = vs1.TextMatrix(I, 1)
       Else
           s10 = s10 & ", " & vs1.TextMatrix(I, 1)
       End If
       RS!paymentamt = Val(vs1.TextMatrix(I, 2))
       RS!amount = Val(vs1.TextMatrix(I, 3))
       RS!RepName = vs1.TextMatrix(I, 4)
       RS!recno = vs1.TextMatrix(I, 5)
       RS.update
      End If
   Next
   '------------------------------------------------
   
   If Len(s10) > 0 Then
      con.Execute "update DNFA SET desc_='" & s10 & "' where DNN=" & TCNN.text & ""
   End If

   
   
   'Frame1.Enabled = False
   cmdAdd = False
   cmadd = False
   cmdedit = False
   Me.Commandadd.Enabled = True
   Me.Commandedit.Enabled = True
   Me.Commandsearch.Enabled = True
   Me.Commandsave.Enabled = False
   Me.Commanddelete.Enabled = True
   Me.Commandabandon.Enabled = True
   Me.CommandPrint.Enabled = True
   
   Gridrefresh
   gencombo1.Visible = False
   Subcombo1.Visible = False
   AText.Visible = False
   DrCrText.Visible = False
   Frame2.Enabled = True
   
   
'  mnuMenu_ = "menudebitnote"
'  ButtonPermission Commandsave, Commanddelete, Commandedit
  
  
   mnuMenu_ = "menudebitnote"
 'SetButton Commandadd, Commandedit, Commandsave, Commanddelete
   
   Me.Commandsave.Enabled = False
   Me.Commanddelete.Enabled = False
   Me.Commandadd.SetFocus
   
   
      If (AuditTrail = "y") Then
    
    If (txtchecked.text = "y") Then
    
        actionType_ = "Edit"
        vtype1_ = "D"
        vtypeNew = "D"
        vdate_ = Trim(Cbdate.text)
        vno_ = Trim(TCNN.text)
        
        frmAuditTrailLog_Rem.Show 1
        
     End If
    
    End If
 
   
End Sub

Private Sub Commandsave_GotFocus()
If Val(Text3) > 0 Then
txtAmtwords = toword(Text3)
End If
End Sub

Private Sub Commandsearch_Click()


searchType = "inv"
popuplist10 "select DNN,DND,PSLD,NA from DNFA where " & stringyear & "  order by DNN", con


End Sub
Private Sub DGrid_AfterUpdate()
Dim rs2 As New ADODB.Recordset
rs2.Open "Select sum(a) as tot from  tempcnf1b where  " & stringyear & "", con, adOpenStatic, adCmdText
If rs2.RecordCount > 0 Then
If IsNull(rs2!Tot) = True Then
    Text3.text = 0
 Else
   Text3.text = rs2!Tot
End If
End If
End Sub
Private Sub Commandsearch_GotFocus()

If PopUpValue1 <> "" Then
     TCNN.text = PopUpValue1
     TCNN_LostFocus
     'totalAmt
     PopUpValue1 = ""
End If

End Sub
Private Sub DGrid_Click()
If DGrid.Row > 0 Then
       Select Case DGrid.Col
    
           Case 0
                gencombo1.text = DGrid.text
                gencombo1.Visible = True
                Subcombo1.Visible = False
                AText.Visible = False
                DrCrText.Visible = False
                cmbgroup.Visible = False
                
                gencombo1.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.top - 15, DGrid.CellWidth
                gencombo1.SetFocus

            Case 1
                Subcombo1.text = DGrid.text
                Subcombo1.Visible = True
                gencombo1.Visible = False
                AText.Visible = False
                DrCrText.Visible = False
                cmbgroup.Visible = False
                
                If Subcombo1.ListCount > 0 Then
                   Subcombo1.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.top - 15, DGrid.CellWidth
                   Subcombo1.SetFocus
                End If
    
            Case 2
                AText = DGrid.text
                AText.Visible = True
                gencombo1.Visible = False
                Subcombo1.Visible = False
                DrCrText.Visible = False
                cmbgroup.Visible = False
                
                AText.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.top - 15, DGrid.CellWidth
                AText.SetFocus
 
            Case 3
                DrCrText.text = "C"
                DrCrText.text = DGrid.text
                If DGrid.text = "" Then DrCrText.text = "C"
                DrCrText.Visible = True
                gencombo1.Visible = False
                Subcombo1.Visible = False
                AText.Visible = False
                cmbgroup.Visible = False
                
                DrCrText.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.top - 15, DGrid.CellWidth
                DrCrText.SetFocus
             Case 4
                cmbgroup.text = DGrid.text
                Subcombo1.Visible = False
                gencombo1.Visible = False
                cmbgroup.Visible = True
              
                AText.Visible = False
                DrCrText.Visible = False
                If cmbgroup.ListCount > 0 Then
                   cmbgroup.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.top - 15, DGrid.CellWidth
                   cmbgroup.SetFocus
                End If

            End Select
   
  End If
    
    
    
    
End Sub

Private Sub DrCrText_Change()
DrCrText.text = UCase(DrCrText.text)
End Sub

Private Sub DrCrText_KeyPress(KeyAscii As Integer)



If KeyAscii = 13 Then
     
  If DrCrText.text = "" Then
        MsgBox "please Enter  D or C."
        DrCrText.SetFocus
        Exit Sub
   End If
     
     
     Glastrow = DGrid.Row
     DGrid.text = DrCrText.text
     DGrid.Row = DGrid.Row + 1
     
     LRC = LRC + 1
     DGrid.Col = 0
     DGrid_Click
     
    Text3.text = 0
    For J = 1 To DGrid.rows - 1
       Text3.text = (Val(Text3.text) + Val(DGrid.TextMatrix(J, 2)))
    Next
     
     
     
End If
End Sub

Private Sub Form_Activate()

Commandadd.Enabled = True
Commandsave.Enabled = False

If Commandadd.Visible = True Then
' Commandadd.SetFocus
End If


  

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
Private Sub Form_Load()
  
   Me.top = 100
   Me.Left = 100
    
   Me.Width = 14568
   Me.Height = 9576
  
   Me.Caption = "Debit Note"
  
  
  
   Dim rs2 As New ADODB.Recordset
   Set RS = New ADODB.Recordset
   cmdAdd = False
   cmdedit = False
   Datachange = False
   Me.Left = 0
   Me.top = 0
   GCombo.Clear
   gencombo1.Clear
   
'   If RS.State = 1 Then RS.close
'   RS.Open " Select *  from DNFA where  " & stringyear & " order by dnn", con, adOpenDynamic, adLockOptimistic, adCmdText
'   If RS.RecordCount > 0 Then
'        RS.MoveLast
'        'Set TCNN.DataSource = RS
'        'Set CdateT.DataSource = RS
'        'Set Text2.DataSource = RS
'        'Set Text3.DataSource = RS
'        'Set GText.DataSource = RS
'        'Set SText.DataSource = RS
'        'TCNN_LostFocus
'   End If
   
   If rs2.State = 1 Then rs2.close
   rs2.Open "Select * from gledger where   " & stringyear & " and  slf = 1 order by gledger", con, adOpenStatic, adLockReadOnly
   If Not rs2.EOF Then
        Do While Not rs2.EOF
            GCombo.AddItem rs2(1)
            If Not rs2.EOF Then
               rs2.MoveNext
            End If
        Loop
   End If
   
   If rs2.State = 1 Then rs2.close
   rs2.Open "Select * from gledger where  " & stringyear & " order by gledger", con, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs2.EOF Then
      Do While Not rs2.EOF
         gencombo1.AddItem rs2(1)
         If Not rs2.EOF Then
            rs2.MoveNext
         End If
      Loop
   End If
   
 
   If rs2.State = 1 Then rs2.close
   rs2.Open "Select * from groups where " & stringyear & " order by groupcode", con, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs2.EOF Then
        Do While Not rs2.EOF
            cmbgroup.AddItem rs2!groupcode
            If Not rs2.EOF Then
               rs2.MoveNext
            End If
        Loop
   End If
 
 
     If rs1.State = 1 Then rs1.close
    rs1.Open "select Rep as Representative from SalesRepQry where (email is not null and len(email)>1) order by Rep", CON_blue
    cmbAgentName.Clear
    If Not rs1.EOF Then
       Do While Not rs1.EOF
          If IsNull(rs1(0)) = False Then
             Me.cmbAgentName.AddItem rs1(0)
             
             If s5 = "" Then
                s5 = rs1(0)
             Else
                s5 = s5 & "|" & rs1(0)
             End If
             
          End If
          If Not rs1.EOF Then rs1.MoveNext
        Loop
    End If
 
    vs1.ColComboList(4) = s5
  '=============================================================
 
   If RS.State = 1 Then RS.close
   RS.Open " Select *  from DNFA where  " & stringyear & " order by dnn", con, adOpenDynamic, adLockOptimistic, adCmdText
   If RS.RecordCount <= 0 Then
        Commandedit.Enabled = False
   End If

   If RS.RecordCount > 0 Then
        
        If inviceNo <> "" Then
          RS.MoveFirst
          RS.Find "dnn=" & inviceNo & ""
          TCNN.text = RS!dnn
          inviceNo = ""
        Else
            RS.MoveLast
            TCNN.text = RS!dnn
        End If
        
        
        TCNN_LostFocus
   End If
 
 
  '=============================================================
 
  BackColorFrom Me
  
  'SetButton Commandadd, Commandedit, Commandsave, Commanddelete
  
  Me.Enabled = True
  
  
 'ButtonPermission Commandsave, Commanddelete, Commandedit
  
 mnuMenu_ = "menudebitnote"
 
 SetButton Commandadd, Commandedit, Commandsave, Commanddelete
 
 Commanddelete.Enabled = False
 Commandsave.Enabled = False
   
  
  
  
 
End Sub

Private Sub Form_Resize()
panel.Left = (Me.ScaleWidth - panel.Width) / 2
panel.top = (Me.ScaleHeight - panel.Height) / 2

End Sub

Private Sub GCombo_Change()
GText.text = GCombo.text

End Sub

Private Sub GCombo_Click()
GText.text = GCombo.text
End Sub

Private Sub GCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Dim GEN As String
     Dim SC1 As String
     
     GEN = GCombo.text
     SC1 = Scombo.text
     
     Dim rs2 As New ADODB.Recordset
            
           rs2.Open "Select * from sledger where  " & stringyear & " and GLEDGER='" + Trim(GCombo.text) + "'", CCON, adOpenStatic, adLockReadOnly, adCmdText
           'Set rs2 = CON.Execute("exec fatch_ledger '" & Trim(GCombo.Text) & "'")
           'ssss = rs2.RecordCount
            
            Scombo.Clear
            If Not rs2.EOF Then
                 Do While Not rs2.EOF
                        ''Scombo.AddItem rs2(1)
                        Scombo.AddItem rs2("subledger")
                        If Not rs2.EOF Then
                            rs2.MoveNext
                        End If
                    Loop
      
            Else
                    Scombo.Clear
                    Scombo.text = ""
                    Text2.SetFocus
                    Exit Sub
            End If
     
        If KeyAscii = 13 Then
                    sendkeys "{tab}"
                    Datachange = False
         End If
  Else
  
     Datachange = True
  
  
  End If
 
  If Datachange = False Then
       GCombo.text = GEN
       Scombo.text = SC1
       Datachange = False
  End If
 
 

End Sub
Private Sub GCombo_LostFocus()
''  If GCombo.Text = "" Then
''            GCombo.SetFocus
''  Else
''       If GCombo.Text <> "" Then
''            Dim rs4 As New ADODB.Recordset
''             rs4.Open "Select* from gledger where   " & stringyear & " and slf = 1 and  GLEDGER='" + Trim(GCombo.Text) + "'", CON, adOpenStatic
''            If rs4.RecordCount <= 0 Then
''                 MsgBox "No valid Gen.Ledger"
''                 'GCombo.SetFocus
''            End If
''
''        End If
''   End If
End Sub

Private Sub gencombo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   If MsgBox("Are You Sure, Delete it ?", vbYesNo) = vbYes Then
      If DGrid.Row >= 1 Then
           DGrid.RemoveItem DGrid.Row
           gencombo1.text = ""
           gencombo1.Visible = False
           DGrid.SetFocus
       End If
   End If
End If

End Sub

Private Sub gencombo1_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
      Dim rs2 As New ADODB.Recordset
      Dim rs3 As New ADODB.Recordset
    
      If gencombo1.text = "" Then
         gencombo1.Visible = False
         Commandsave.SetFocus
         Exit Sub
      End If
      If gencombo1.text <> "" Then
          Dim rs4 As New ADODB.Recordset
          rs4.Open "Select* from gledger where  " & stringyear & " and GLEDGER='" + Trim(gencombo1.text) + "'", con, adOpenForwardOnly, adLockReadOnly, adCmdText
          If rs4.RecordCount <= 0 Then
             MsgBox "No valid Gen.Ledger"
             gencombo1.Visible = True
             gencombo1.SetFocus
             Exit Sub
          End If
     
         DGrid.text = gencombo1.text
         'If rs4!slf = True Then
            Subcombo1.Clear
            rs2.Open "Select * from sledger where  " & stringyear & " and GLEDGER='" + Trim(gencombo1.text) + "'", con, adOpenForwardOnly, adLockReadOnly, adCmdText
            DGrid.Col = 0
            If rs2.RecordCount > 0 Then
                rs2.MoveFirst
                Do While Not rs2.EOF
                    Subcombo1.AddItem rs2(1)
                    If Not rs2.EOF Then
                        rs2.MoveNext
                    End If
                Loop
                DGrid.Col = DGrid.Col + 1
                DGrid_Click
            Else
                DGrid.Col = DGrid.Col + 1
                DGrid.text = ""
                DGrid.Col = DGrid.Col + 1
                Subcombo1.Visible = False
                DGrid_Click
            End If
         'End If
   End If
End If



End Sub

Private Sub GText_Change()
GCombo.text = GText.text

End Sub

Private Sub Scombo_Change()
   SText.text = Scombo.text
End Sub

Private Sub Scombo_Click()
    SText.text = Scombo.text
End Sub

Private Sub Scombo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
     If cmdAdd = True Then
        'SendKeys "{Down}"
     End If
   sendkeys "{tab}"
End If
End Sub
Private Sub Scombo_LostFocus()

''If Scombo.Text <> "" And GCombo.Text <> "" Then
''  Dim rs4 As New ADODB.Recordset
''  rs4.Open "Select* from sledger where  " & stringyear & " and GLEDGER='" + Trim(GCombo.Text) + "' and SubLedger='" + Trim(Scombo.Text) + "'", CON, adOpenStatic
''  If rs4.RecordCount <= 0 Then
''     MsgBox "No valid Sub Ledger"
''     Scombo.SetFocus
''  End If
''
''
''End If
''If Scombo.ListCount > 0 And Scombo.Text = "" Then
''     Scombo.SetFocus
''End If

End Sub
Function checkData() As Boolean
  
checkData = False
  
If Scombo.text <> "" And GCombo.text <> "" Then
  Dim rs4 As New ADODB.Recordset
  rs4.Open "Select* from sledger where  " & stringyear & " and GLEDGER='" + Trim(GCombo.text) + "' and SubLedger='" + Trim(Scombo.text) + "'", con, adOpenStatic
  If rs4.RecordCount <= 0 Then
     MsgBox "No valid Sub Ledger"
     checkData = True
     Scombo.SetFocus
  End If


End If
If Scombo.ListCount > 0 And Scombo.text = "" Then
     Scombo.SetFocus
End If
  
  
End Function
Private Sub SText_Change()
 Scombo.text = SText.text
End Sub

Private Sub Subcombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Subcombo1.text <> "" And gencombo1.text <> "" Then

                Dim rs4 As New ADODB.Recordset
                rs4.Open "Select* from sledger where  " & stringyear & " and GLEDGER='" + Trim(gencombo1.text) + "' and SubLedger='" + Trim(Subcombo1.text) + "'", con, adOpenStatic
                  If rs4.RecordCount <= 0 Then
                            MsgBox "No valid Sub Ledger"
                            Subcombo1.SetFocus
                            Exit Sub
                  End If

       End If
       If Subcombo1.ListCount > 0 And Subcombo1.text = "" Then
               Subcombo1.SetFocus
               Exit Sub
       End If
       DGrid.Col = 0
       If DGrid.text = "" Then
                 DGrid.Col = 1
                 If DGrid.text = "" Then
                      Commandsave.SetFocus
                      Exit Sub
                 End If
       End If
       DGrid.Col = 1
       DGrid.text = Subcombo1.text
       DGrid.Col = DGrid.Col + 1
       DGrid_Click
 End If
 
 
 
End Sub

Private Sub TCNN_Change()
 Gridrefresh
 
    
End Sub

Private Sub TCNN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  sendkeys "{tab}"
End If
End Sub
Private Sub TCNN_LostFocus()

Dim rs1 As New ADODB.Recordset
     If TCNN.text = "" Then
        Exit Sub
     End If
     
     
     '================================================================
     
     If rs1.State = 1 Then rs1.close
     rs1.Open "select * FROM  DNFA WHERE   " & stringyear & " and DNN=" + Trim(TCNN.text) + "", con, adOpenDynamic, adLockOptimistic
     If rs1.EOF = False Then
     
     
     If (AuditTrail = "y") Then
        If (rs1!Checked_YesNo = True) Then
            txtchecked.text = "y"
        Else
            txtchecked.text = "n"
        End If
      End If
        
     
        TCNN.text = rs1!dnn
        Cbdate.text = rs1!dnd
        GText.text = rs1!Pgld
        SText.text = rs1!psld & ""
        Text2.text = rs1!n
        Text3.text = rs1!na & ""
        cmbAgentName.text = rs1!agentname & ""
        
        Gridrefresh
     End If
     
     
     
     
     
 '=================================================================
     
   If rs1.State = 1 Then rs1.close
   rs1.Open "select * from DebitNotDet where dnn=" & TCNN.text & " order by SN", con, adOpenDynamic, adLockOptimistic
   For I = 1 To rs1.RecordCount
      If RS.EOF = False Then
       vs1.TextMatrix(I, 0) = rs1!sn
       vs1.TextMatrix(I, 1) = rs1!NARR
       vs1.TextMatrix(I, 2) = rs1!paymentamt & ""
       vs1.TextMatrix(I, 3) = rs1!amount
       vs1.TextMatrix(I, 4) = rs1!RepName & ""
       vs1.TextMatrix(I, 5) = rs1!recno & ""
       rs1.MoveNext
      End If
   Next
  

 '=================================================================
     
     
     
     '=================================================================
     
     If rs1.State = 1 Then rs1.close
     rs1.Open "Select top 10 * from  DNFA  where  " & stringyear & " and DNN = " + TCNN.text + "", con, adOpenStatic, adLockOptimistic, adCmdText
     If Not rs1.EOF Then
     If cmdAdd Then
            MsgBox "Invoice already exist..."
            TCNN.SetFocus
            Exit Sub
     End If
     
 
     
     
     

End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  DGrid.Col = 0
  DGrid.SetFocus
  DGrid.Row = 1
  DGrid_Click
  
End If

KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text3_Change()
Text3 = Format(Val(Text3.text), "0.00")
End Sub
Private Sub vs1_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      If vs1.Col = 1 Then
         
         If vs1.TextMatrix(vs1.RowSel, 1) <> "" Then
           vs1.TextMatrix(vs1.RowSel, 1) = UCase(vs1.TextMatrix(vs1.RowSel, 1))
           
            'If cboCat.Text = "CD" Then
               sendkeys "{right}"
            'Else
            '   SendKeys "{right}"
            '   SendKeys "{right}"
            'End If
         
           
         End If
      
      ElseIf vs1.Col = 2 Then
         
             sendkeys "{right}"
          
      ElseIf vs1.Col = 3 Then
         
          
           vs1.TextMatrix(vs1.RowSel, 2) = Val(vs1.TextMatrix(vs1.RowSel, 2))
           sendkeys "{right}"
      
      ElseIf vs1.Col = 4 Then
         
         If vs1.Row = 1 Then
           If vs1.TextMatrix(1, 4) <> "" Then
              cmbAgentName.text = vs1.TextMatrix(1, 4)
           End If
         End If
      
         sendkeys "{home}"
         sendkeys "{down}"
         
      If vs1.Row = 5 Then
         sendkeys "{tab}"
      End If
         
      End If
   End If
End Sub

Private Sub vs1_SelChange()
If vs1.Row = 1 Then
  If vs1.TextMatrix(1, 4) <> "" Then
     cmbAgentName.text = vs1.TextMatrix(1, 4)
  End If
End If
End Sub
