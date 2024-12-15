VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Creditnotefile 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   10056
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   16188
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10479.41
   ScaleMode       =   0  'User
   ScaleWidth      =   22288.21
   Begin Crystal.CrystalReport cr 
      Left            =   13140
      Top             =   6480
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame panel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Credit Note"
      Enabled         =   0   'False
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
      Height          =   9912
      Left            =   60
      TabIndex        =   14
      Top             =   60
      Width           =   16104
      Begin VB.TextBox txtchecked 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   3936
         MaxLength       =   100
         TabIndex        =   45
         Top             =   8244
         Width           =   540
      End
      Begin VB.CheckBox Check1_withheader 
         BackColor       =   &H00E0E0E0&
         Caption         =   "With Header"
         Height          =   195
         Left            =   7125
         TabIndex        =   44
         Top             =   8328
         Width           =   1350
      End
      Begin VB.ComboBox cbocat 
         Height          =   315
         ItemData        =   "CNFile.frx":0000
         Left            =   8460
         List            =   "CNFile.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   780
         Width           =   2100
      End
      Begin VB.TextBox txtTODNO 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   960
         MaxLength       =   100
         TabIndex        =   40
         Top             =   8232
         Width           =   1035
      End
      Begin VB.ComboBox cmbgroup 
         Height          =   1488
         Left            =   7140
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   38
         Top             =   5748
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   504
         Left            =   180
         TabIndex        =   35
         Top             =   7452
         Visible         =   0   'False
         Width           =   1830
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
            Height          =   390
            Left            =   900
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   60
            Width           =   855
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
            Height          =   390
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   60
            Width           =   855
         End
      End
      Begin VB.PictureBox pic1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   810
         Left            =   180
         ScaleHeight     =   816
         ScaleWidth      =   9192
         TabIndex        =   26
         Top             =   8808
         Width           =   9195
         Begin VB.CommandButton Commandsave 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Sa&ve"
            Enabled         =   0   'False
            Height          =   675
            Left            =   2205
            Picture         =   "CNFile.frx":002E
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   60
            Width           =   1110
         End
         Begin VB.CommandButton Commandabandon 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Aba&ndon"
            Height          =   675
            Left            =   3390
            Picture         =   "CNFile.frx":0C12
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   60
            Width           =   1065
         End
         Begin VB.CommandButton Commandedit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Edit"
            Height          =   675
            Left            =   1140
            Picture         =   "CNFile.frx":119C
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   60
            Width           =   1065
         End
         Begin VB.CommandButton Commanddelete 
            BackColor       =   &H00FFFFFF&
            Caption         =   "De&lete"
            Height          =   675
            Left            =   4485
            Picture         =   "CNFile.frx":15DE
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   60
            Width           =   1065
         End
         Begin VB.CommandButton Commandsearch 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Search"
            Height          =   675
            Left            =   5580
            Picture         =   "CNFile.frx":21C2
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   60
            Width           =   1185
         End
         Begin VB.CommandButton CommandPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Print"
            Height          =   675
            Left            =   6840
            Picture         =   "CNFile.frx":2DA6
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   60
            Width           =   1185
         End
         Begin VB.CommandButton CommandReturn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Return"
            Height          =   675
            Left            =   8040
            Picture         =   "CNFile.frx":398A
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   60
            Width           =   1065
         End
         Begin VB.CommandButton Commandadd 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Add"
            Height          =   675
            Left            =   45
            Picture         =   "CNFile.frx":456E
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   60
            Width           =   1065
         End
         Begin VB.CommandButton cmdNHPrint 
            BackColor       =   &H00FFFFFF&
            Caption         =   "N&HPrint"
            Height          =   675
            Left            =   6675
            Picture         =   "CNFile.frx":5152
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   840
            Visible         =   0   'False
            Width           =   75
         End
      End
      Begin VB.ComboBox cmbAgentName 
         Height          =   288
         Left            =   3180
         TabIndex        =   24
         Top             =   7584
         Visible         =   0   'False
         Width           =   2820
      End
      Begin VB.TextBox TCNN 
         Appearance      =   0  'Flat
         DataField       =   "Cnn"
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   390
         Width           =   1065
      End
      Begin VB.TextBox GText 
         DataField       =   "pgld"
         Height          =   285
         Left            =   8340
         TabIndex        =   17
         Top             =   1128
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.TextBox SText 
         DataField       =   "psld"
         Height          =   285
         Left            =   6810
         TabIndex        =   16
         Top             =   1140
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "na"
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
         Left            =   7305
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   7560
         Width           =   1305
      End
      Begin VB.ComboBox Subcombo1 
         Height          =   1872
         Left            =   3420
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   11
         Top             =   5556
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.ComboBox gencombo1 
         Height          =   1872
         Left            =   360
         Sorted          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   9
         Top             =   5556
         Visible         =   0   'False
         Width           =   4275
      End
      Begin VB.ComboBox GCombo 
         Height          =   315
         Left            =   1680
         TabIndex        =   4
         Top             =   780
         Width           =   4830
      End
      Begin VB.ComboBox Scombo 
         Height          =   288
         Left            =   1680
         TabIndex        =   6
         Top             =   1116
         Width           =   4815
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         DataField       =   "n"
         Height          =   324
         Left            =   1500
         MaxLength       =   250
         TabIndex        =   8
         Top             =   5148
         Width           =   10608
      End
      Begin VB.TextBox CdateT 
         Appearance      =   0  'Flat
         DataField       =   "Cnd"
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
         Left            =   4515
         TabIndex        =   3
         Top             =   390
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.TextBox AText 
         Appearance      =   0  'Flat
         DataField       =   "cnd"
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
         Left            =   7290
         TabIndex        =   12
         Top             =   5664
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox DrCrText 
         Appearance      =   0  'Flat
         DataField       =   "cnd"
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
         Left            =   8520
         MaxLength       =   1
         TabIndex        =   13
         Text            =   "D"
         Top             =   5664
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSMask.MaskEdBox Cbdate 
         Height          =   315
         Left            =   3330
         TabIndex        =   2
         Top             =   390
         Width           =   1125
         _ExtentX        =   1969
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid DGrid 
         Height          =   1908
         Left            =   168
         TabIndex        =   10
         Top             =   5532
         Width           =   11952
         _ExtentX        =   21082
         _ExtentY        =   3366
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   7917545
         BackColorBkg    =   14737632
         AllowUserResizing=   2
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin MSMask.MaskEdBox txtTODDate 
         Height          =   312
         Left            =   2040
         TabIndex        =   41
         Top             =   8232
         Width           =   1032
         _ExtentX        =   1820
         _ExtentY        =   550
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VSFlex7Ctl.VSFlexGrid VS1 
         Height          =   3540
         Left            =   108
         TabIndex        =   7
         Top             =   1620
         Width           =   15816
         _cx             =   27898
         _cy             =   6244
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
         Rows            =   11
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   310
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
            TabIndex        =   47
            Top             =   6360
            Width           =   195
         End
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Checked :"
         Height          =   252
         Left            =   3168
         TabIndex        =   46
         Top             =   8280
         Width           =   684
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Note Category"
         Height          =   255
         Left            =   6900
         TabIndex        =   43
         Top             =   780
         Width           =   1635
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "V. No :"
         Height          =   252
         Left            =   240
         TabIndex        =   42
         Top             =   8232
         Width           =   792
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
         Height          =   372
         Index           =   3
         Left            =   96
         TabIndex        =   39
         Top             =   1392
         Width           =   1620
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0078CFE9&
         BorderWidth     =   4
         Height          =   960
         Left            =   48
         Top             =   8712
         Width           =   9396
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
         Height          =   192
         Index           =   2
         Left            =   2112
         TabIndex        =   25
         Top             =   7644
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Note No."
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
         Left            =   96
         TabIndex        =   23
         Top             =   456
         Width           =   1332
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
         Left            =   168
         TabIndex        =   22
         Top             =   5172
         Width           =   792
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
         Height          =   192
         Left            =   96
         TabIndex        =   21
         Top             =   1080
         Width           =   996
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
         Height          =   348
         Left            =   96
         TabIndex        =   20
         Top             =   780
         Width           =   1368
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
         Left            =   6240
         TabIndex        =   19
         Top             =   7620
         Width           =   1008
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
         Left            =   2820
         TabIndex        =   18
         Top             =   465
         Width           =   420
      End
   End
End
Attribute VB_Name = "Creditnotefile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RS As ADODB.Recordset
'Dim con As New ADODB.Connection
Dim mvBookMark As Variant
Dim cmdAdd As Boolean
Dim cmdedit As Boolean
Dim LRC As Integer
Dim LCC As Integer
Dim Glastrow As Integer
Dim Datachange As Boolean
Sub CNbandon()

Me.Commandadd.Enabled = True
Me.Commandedit.Enabled = True
Me.Commandsearch.Enabled = True
Me.Commandsave.Enabled = False
Me.Commanddelete.Enabled = True
Me.Commandabandon.Enabled = True
Me.CommandPrint.Enabled = True
Me.cmdNHPrint.Enabled = True
        
End Sub
Sub Setgrid()
  
  DGrid.Cols = 5
  
  DGrid.Row = 0
  DGrid.Col = 0
  DGrid.text = "Gen Ledger"
  DGrid.Col = 1
  DGrid.text = "Sub Ledger"
  DGrid.Col = 2
  'DGrid.Text = Format(DGrid.Text, "0.00")
  DGrid.text = "Amount"
  DGrid.Col = 3
  DGrid.text = "Debit/Credit"
  
  DGrid.Col = 4
  DGrid.text = "Group"
  DGrid.ColWidth(0) = 3800
  DGrid.ColWidth(1) = 3200
  DGrid.ColWidth(2) = 1500
  DGrid.ColWidth(3) = 1500
  DGrid.ColWidth(4) = 1500
  DGrid.Col = 0


   
 vs1.Cols = 7
   
 vs1.Clear
 vs1.FormatString = "SN.|Description|TRec.Amt|Amount|Rep.Name"
 
 vs1.TextMatrix(1, 0) = "1"
 vs1.TextMatrix(2, 0) = "2"
 vs1.TextMatrix(3, 0) = "3"
 vs1.TextMatrix(4, 0) = "4"
 vs1.TextMatrix(5, 0) = "5"
 vs1.TextMatrix(6, 0) = "6"
 vs1.TextMatrix(7, 0) = "7"
 vs1.TextMatrix(8, 0) = "8"
 vs1.TextMatrix(9, 0) = "9"
 vs1.TextMatrix(10, 0) = "10"
 
 vs1.ColHidden(5) = True
 vs1.ColHidden(6) = True
 
 vs1.ColWidth(0) = 500
 vs1.ColWidth(1) = 10400
 vs1.ColWidth(2) = 950
 vs1.ColWidth(3) = 1150
 vs1.ColWidth(4) = 1250
 ''vs1.ColWidth(6) = 1000


 

End Sub
Sub Gridrefresh()
cmbgroup.Visible = False

DoEvents
      Dim grs1 As New ADODB.Recordset
      If grs1.State = 1 Then grs1.close
         If TCNN.text = "" Then
            grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit,Groupcode FROM  CNF1B WHERE  " & stringyear & " and CNN = 0", con, adOpenStatic
            Set DGrid.DataSource = grs1
            DGrid.Refresh
            Setgrid

         ElseIf cmdedit = True Then
             DoEvents
                grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit,Groupcode FROM  CNF1B where  " & stringyear & " and CNN = " + Trim(TCNN.text) + "", con, adOpenStatic
                DGrid.rows = 99
                DGrid.TopRow = 1
                Setgrid
                
         ElseIf cmdAdd = True Then
         DoEvents
                grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit,Groupcode FROM  CNF1B where  " & stringyear & " and CNN = " + Trim(TCNN.text) + "", con, adOpenStatic
                DGrid.rows = 99
                DGrid.TopRow = 1
                Setgrid
         Else
         DoEvents
                grs1.Open "SELECT gld as GenLedger,Sld as SubLedger, a as Amount ,Dc as DebitCredit,Groupcode FROM  CNF1B WHERE   " & stringyear & " and CNN = " + Trim(TCNN.text) + "", con, adOpenStatic
                Set DGrid.DataSource = grs1
                DGrid.Refresh
                Setgrid
         End If
         
     If TCNN.text <> "" Then
     DoEvents
        For I = 1 To grs1.RecordCount - 1
          DGrid.Row = I
          DGrid.Col = 2
          DGrid.text = Format(DGrid.text, "0.00")
          DGrid.Refresh
        Next I
    End If
    DoEvents
    DGrid.Refresh
    DGrid.Col = 0
    
    
   If TCNN.text <> "" Then
   
    If rs1.State = 1 Then rs1.close
    rs1.Open "select * from CreditNotDet where cnn=" & TCNN.text & " order by SN", con, adOpenDynamic, adLockOptimistic
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
        vs1.TextMatrix(I, 6) = rs1!fyear & ""
        rs1.MoveNext
       End If
    Next
   
   End If
    
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

Private Sub AText_LostFocus()
If Val(AText.text) <= 0 Then
        AText.Visible = True
        AText.SetFocus
        
        Exit Sub
 End If
'totalAmt
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
   Cbdate.SetFocus
   Exit Sub
End If
If Trim(Cbdate.text) <> "__/__/____" Then
    If Not checkdate(Trim(Cbdate.text), Cbdate) Then
        Cbdate.SetFocus
    End If
End If
End Sub

Private Sub cbocat_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    sendkeys "{tab}"
 End If
End Sub

Private Sub CdateT_Change()
 'On Error GoTo er1
 
 If CdateT = "" Then
  
    'Cbdate.Text = "__/__/____"
 Else
   Cbdate.text = CdateT.text
 End If
er1:  If err.Number = 380 Then
       Exit Sub
      End If
End Sub

Private Sub cmbAgentName_LostFocus()
If cmbAgentName.text = "" Then
   MsgBox "Enter a Agent Name.. "
   'cmbAgentName.SetFocus
   Exit Sub
Else
  Dim rs1 As New ADODB.Recordset
  rs1.Open "select rep  from SalesRepQry where rep='" & cmbAgentName.text & "'", CON_blue
  If rs1.EOF = True Then
     MsgBox "Enter valid Agent Name.. "
     cmbAgentName.SetFocus
  End If
End If
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

Private Sub cmdNHPrint_Click()
''     printch = "cnf1a"
''     ino = TCNN
''     printch1 = "CNN"
''
''
''Printheader = False
''GenrateReport

s1 = 5
PrintOption.Show
 
End Sub

Private Sub Command1_Click()
On Error GoTo er
If Not RS.BOF Or Not RS.EOF Then
    RS.MoveNext
End If
If RS.EOF And RS.RecordCount > 0 Then
    Beep
    RS.MoveLast
  End If
  
er:   If err.Number = 3021 Then
         Exit Sub
      End If
  End Sub

Private Sub Command2_Click()
  Dim gs As ADODB.Recordset
  Set gs = New ADODB.Recordset
  If Not RS.BOF Then
     RS.MovePrevious
    
 
End If
 DoEvents
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
  
  If RS.RecordCount > 0 Then
      RS.CancelUpdate
      RS.MoveFirst
  End If
  
  cmbAgentName.text = ""
  
  cmdAdd = False
  cmdedit = False
  Gridrefresh
  pic1.Visible = True
  Frame2.Enabled = True
  gencombo1.Visible = False
  Subcombo1.Visible = False
  AText.Visible = False
  DrCrText.Visible = False
  Frame1.Enabled = False
  CNbandon
  
  
  mnuMenu_ = "menucreditnote"
  SetButton Commandadd, Commandedit, Commandsave, Commanddelete
  
End Sub

Private Sub Commandadd_Click()
 
On Error Resume Next
 
 Dim rs6 As New ADODB.Recordset
 Me.Commandadd.Enabled = False
 Me.Commandedit.Enabled = False
 Me.Commandsearch.Enabled = False
 Me.Commandsave.Enabled = True
 Me.Commanddelete.Enabled = False
 Me.Commandabandon.Enabled = True
 Me.CommandPrint.Enabled = False
 
 txtchecked.text = ""
 
 Dim rs1  As New ADODB.Recordset
cmdAdd = True
cmdedit = False
LRC = 1
LCC = 0
'frame1.Enabled = True
TCNN.Enabled = True
TCNN.SetFocus
cmbAgentName.text = ""

'With RS
'    .AddNew
'End With

CdateT.text = ""
Text2.text = ""

txtTODNO.text = ""
txtTODDate.text = "__/__/____"


If cmdedit = False Then
     Dim trs As New ADODB.Recordset
     trs.Open "Select max(cnn)as mcnn from cnf1a where  " & stringyear & "", con, adOpenStatic, adCmdText
     If trs.RecordCount <= 1 And IsNull(trs!Mcnn) Then
       TCNN.text = 1
     Else
       TCNN.text = trs!Mcnn + 1
    End If
End If

DGrid.rows = 100
DGrid.Cols = 0
DGrid.Cols = 5
Dim I
For I = 0 To 99
   DGrid.RowHeight(I) = 270
Next
Setgrid




DoEvents
GCombo.text = "SUNDRY DEBTORS"
Dim rs2 As New ADODB.Recordset
rs2.Open "Select * from sledger where  " & stringyear & " and GLEDGER='" + Trim(GCombo.text) + "'", con, adOpenForwardOnly, adLockReadOnly, adCmdText
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
Text3 = ""
'cboCat.ListIndex = 3


mnuMenu_ = "menucreditnote"
'SetButton Commandadd, Commandedit, Commandsave, Commanddelete
  

End Sub

Private Sub Commanddelete_Click()
On Error Resume Next


Dim rs1 As New ADODB.Recordset
Dim rs_h As New ADODB.Recordset

 If rs1.State = 1 Then rs1.close
    rs1.Open "select top 100 * from CNF1A where CNN=" & Trim(TCNN.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
    If rs1.EOF = False Then
       If rs_h.State = 1 Then rs_h.close
       rs_h.Open "select top 100 * from CNF1A where cnn=" & Trim(TCNN.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
           If rs1!bAuthorized = True Then
              MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
              Exit Sub
          End If
    End If


createLog UserName, TCNN, "credit note", " Delete : " & Text3.text, Date


If RS.RecordCount > 0 Then
  If MsgBox("Are you sure.......", vbYesNo) = vbYes Then
  If TCNN.text <> "" Then
  
  
      If (AuditTrail = "y") Then
      
      If (txtchecked.text = "y") Then
      
          actionType_ = "Delete"
          vtype1_ = "C"
          vtypeNew = "C"
          vdate_ = Trim(Cbdate.text)
          vno_ = Trim(TCNN.text)
          
          frmAuditTrailLog_Rem.Show 1
          
       End If
      
      End If
    
  
       con.Execute "DELETE FROM  CNF1A WHERE   " & stringyear & " and CNN=" + Trim(TCNN.text) + ""
       con.Execute "DELETE FROM  CNF1B WHERE   " & stringyear & " and CNN=" + Trim(TCNN.text) + ""
       con.Execute "DELETE FROM  CreditNotDet where CNN=" + Trim(TCNN.text) + ""
  End If
  
  With RS
    .delete
    
     If Not RS.BOF And Not RS.EOF Then RS.MoveFirst
    
     End With
    
   Gridrefresh
    
   GCombo.text = ""
   Scombo.text = ""
   TCNN.text = ""
   Cbdate.text = "__/__/____"
   GText.text = ""
   cmbAgentName.text = ""
   SText.text = ""
   Text2.text = ""
   Commandadd_Click
   
  
  
  Exit Sub


End If

End If


End Sub

 Sub Commandedit_Click()
 
 
Dim rs1 As New ADODB.Recordset
Dim rs_h As New ADODB.Recordset

 If rs1.State = 1 Then rs1.close
    rs1.Open "select top 100 * from CNF1A where CNN=" & Trim(TCNN.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
    If rs1.EOF = False Then
       If rs_h.State = 1 Then rs_h.close
       rs_h.Open "select top 100 * from CNF1A where cnn=" & Trim(TCNN.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
           If rs1!bAuthorized = True Then
              MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
              Exit Sub
          End If
    End If
 
 
 
 
 
 If RS.RecordCount > 0 Then
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
        'Frame1.Enabled = True
        Cbdate.SetFocus

        Gridrefresh
        
        If cmdedit = True Then
          Dim rs3 As New ADODB.Recordset
        rs3.Open "SELECT * FROM  CNF1B WHERE   " & stringyear & " and CNN=" + Trim(TCNN.text) + "", con, adOpenStatic, adLockOptimistic, adCmdText
          If rs3.RecordCount >= 0 Then
             DGrid.rows = 99
             DGrid.TopRow = 1
             DGrid.Row = rs3.RecordCount + 1
          End If
    End If
    
    
    Frame2.Enabled = False
    VB.Screen.ActiveForm.ActiveControl.SelLength = 2
 End If
 
 
 mnuMenu_ = "menucreditnote"
  SetButton Commandadd, Commandedit, Commandsave, Commanddelete
  
 
End Sub
Private Sub CommandPrint_Click()
    
'CR.Reset
'CR.Connect = constr
'CR.ReportFileName = strrptpath & "\REPORTS\CREDITNote.rpt"
'CR.ReplaceSelectionFormula "{CNF1A.CNN} = " & TCNN & ""
'CR.WindowState = crptMaximized
'CR.Action = 1

''printch = "cnf1a"
''ino = TCNN
''printch1 = "CNN"
''
''
''Printheader = True
'GenrateReport
     

s1 = "5"
PrintOption.Show

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
   Print #1, Chr(27) + Chr(71); Chr(27) + Chr(72); Tab((paperWidth - (Len(Trim("CREDIT NOTE")))) / 2 - 3); Chr(14); "CREDIT NOTE"; Chr(20)
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
If rs7.State = 1 Then rs7.close
'rs7.Open "Select * from cnf1a where cnn =" & TCNN.Text & "  and cnd = cdate('" + Trim(Cbdate.Text) + "')", CON
rs7.Open "Select * from cnf1a where " & stringyear & " and cnn =" & TCNN.text & "  and cnd = convert(smalldatetime,'" + Trim(Cbdate.text) + "',103)", con

If rs7.RecordCount > 0 Then
   Print #1, Chr(27) + Chr(71); "To,   S.L. Code : "; Tab(20); Mid$(rs7!psld, 1, 5); Tab(50); "Credit Note No. : "; Chr(27) + Chr(72); Trim(rs7!cnn); Tab(83); Chr(27) + Chr(71); "Date : "; Chr(27) + Chr(72); rs7!Cnd
   Line = Line + 1
   If kkk.State = 1 Then kkk.close
   kkk.Open "select * from sledger where " & stringyear & " and subledger='" + Trim(rs7!psld) + "'", con, adOpenStatic, adLockReadOnly, adCmdText
   If Not kkk.EOF Then
      Print #1, Tab(5); "M/s " & kkk!DESCFORINVOICE
      Print #1, Tab(5); IIf(IsNull(kkk!address1), " ", kkk!address1); Tab(50); " Agent Name      : " + Creditnotefile.cmbAgentName
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
   trs.Open "Select * from cnf1b where " & stringyear & " and cnn =" & TCNN.text & "  and Cnd = convert(smalldatetime,'" + Trim(Cbdate.text) + "',103)", con
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
    Print #1, Tab(1); "Net Amount Cr. In Your A/C : "; Tab(80); rsets(Trim(Format(Str(rs7!na), "0.00")), 12)
    Print #1, ""
    Print #1, Tab(1); toword(rs7!na)
    Print #1, repli("-", paperWidth)
    Dim tempdata As ADODB.Recordset
    Set tempdata = New ADODB.Recordset
    CNSetup
    tempdata.Open "setup1", con, adOpenStatic, adLockReadOnly, adCmdTable
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
Private Sub CommandReturn_Click()
'rs.Close
Unload Me
'''''MainMenu.Toolbar1.Visible = True
End Sub
Private Sub Commandsave_Click()

On Error GoTo err:


Dim Checked_YesNo  As Integer

 If (txtchecked.text = "y") Then
      Checked_YesNo = 1
 Else
      Checked_YesNo = 0
 End If
 
 
   
Dim rs1 As New ADODB.Recordset
Dim rs_h As New ADODB.Recordset

 If rs1.State = 1 Then rs1.close
    rs1.Open "select top 100 * from CNF1A where CNN=" & Trim(TCNN.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
    If rs1.EOF = False Then
       If rs_h.State = 1 Then rs_h.close
       rs_h.Open "select top 100 * from CNF1A where cnn=" & Trim(TCNN.text) & " and " & stringyear, con, adOpenKeyset, adLockReadOnly
           If rs1!bAuthorized = True Then
              MsgBox "You can'nt change, Bill Already Locked !!", vbExclamation, "Alert"
              Exit Sub
          End If
    End If
   
   createLog UserName, TCNN, "credit note", " Save/Edit : " & Text3.text, Date
   
   sum1 = 0
   If cboCat = "CD" Then
        For k1 = 1 To vs1.rows - 1
               If vs1.TextMatrix(k1, 2) <> "" Then
                  sum1 = sum1 + Val(vs1.TextMatrix(k1, 2))
                  GoTo aaa10
               End If
        Next
        
aaa10:
        If sum1 = 0 Then
        MsgBox "Enter TRec.Amt  ....", vbCritical
        Exit Sub
        End If
   
   End If
   
   
   
   
   
   If checkData = True Then
      Exit Sub
   End If
   
   
   Dim rs_A As New ADODB.Recordset
   Dim Grs As New ADODB.Recordset
   
   If TCNN.text = "" Then
       Commandabandon_Click
       Exit Sub
   End If
    
   con.Execute "DELETE FROM  CNF1B WHERE  " & stringyear & " and CNN=" + Trim(TCNN.text) + ""
   TCNN.Enabled = True
   
   
   
   Grs.Open "select * from cnf1b where  " & stringyear & "", con, adOpenDynamic, adLockPessimistic
   
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
        Grs!cnn = Val(TCNN.text)
        Grs!Cnd = Cbdate.text
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
                        Grs!a = Format(Val(DGrid.text), "0.00")
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
   
   If rs_A.State = 1 Then rs_A.close
   rs_A.Open "select * from cnf1a where cnn=" & TCNN.text & "", con, adOpenDynamic, adLockPessimistic
   If rs_A.EOF = True Then
   rs_A.AddNew
   End If
      
      
    If txtTODNO.text <> "" Then
       rs_A!todid = txtTODNO.text
       If IsDate(txtTODDate.text) Then
          rs_A!toddate = txtTODDate.text
       End If
    Else
       rs_A!todid = Null
       rs_A!toddate = Null
    End If
      
   If (AuditTrail = "y") Then
      rs_A!Checked_YesNo = Checked_YesNo
   End If
   
   rs_A!cnn = TCNN.text
   rs_A!Cnd = Cbdate.text
   rs_A!Pgld = GText.text
   
   If cmbAgentName.text <> vs1.TextMatrix(1, 4) Then
      cmbAgentName.text = vs1.TextMatrix(1, 4)
   End If
   rs_A!agentname = cmbAgentName.text
   
   If SText.text = "" Then
      rs_A!psld = Null
   Else
      rs_A!psld = SText.text
   End If

   If Text2.text <> "" Then
   rs_A!n = Text2.text
   End If
   If sum >= 0 Then
       rs_A!dc = "C"
   Else
       rs_A!dc = "D"
   End If
   
   rs_A!na = Abs(Format(sum, "0.00"))
   rs_A!fyear = main.session
   rs_A!setupid = main.setupid
   rs_A!CNCategory = cboCat.text
  
   If sum > 0 Then
      rs_A.update
   End If
   
   '------------------------------------------------
   
   Set RS = New ADODB.Recordset
   RS.Open "select * from CreditNotDet where cnn=" & TCNN.text & "", con, adOpenDynamic, adLockOptimistic
   If RS.EOF = False Then
      con.Execute "delete from CreditNotDet where cnn=" & TCNN.text & ""
   End If
   s10 = ""
   For I = 1 To vs1.rows - 1
      If vs1.TextMatrix(I, 1) <> "" Then
       RS.AddNew
       RS!fyear = vs1.TextMatrix(I, 6)
       RS!sn = vs1.TextMatrix(I, 0)
       RS!cnn = TCNN.text
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
       
       RS!paymentamt = (vs1.TextMatrix(I, 2))
       RS!amount = Val(vs1.TextMatrix(I, 3))
       RS!RepName = vs1.TextMatrix(I, 4)
       RS!recno = vs1.TextMatrix(I, 5)
       
       RS.update
       
      End If
   Next
   '------------------------------------------------
   
   
   If s10 <> "" Then
     con.Execute "update CNF1A set desc_='" & s10 & "' where cnn=" & TCNN.text & ""
   End If
   
   If InStr(Scombo.text, "(EM)") > 0 Then
      con.Execute "update cnf1a SET saletype='EM' where cnn=" & TCNN.text & ""
   Else
      con.Execute "update cnf1a SET saletype='BP' where cnn=" & TCNN.text & ""
   End If
   
  
       
       cmdAdd = False
       cmadd = False
       cmdedit = False
        Me.Commandadd.Enabled = True
        Me.Commandedit.Enabled = True
        Me.Commandsearch.Enabled = True
        Me.Commandsave.Enabled = False
        Me.Commanddelete.Enabled = True
        Me.Commandabandon.Enabled = True
        Me.cmdNHPrint.Enabled = True
        Gridrefresh
        gencombo1.Visible = False
        Subcombo1.Visible = False
        AText.Visible = False
        DrCrText.Visible = False
        
        Frame2.Enabled = True
        CommandPrint.Enabled = True
        Commandsave.Enabled = False
        Commanddelete.Enabled = False
        Commandadd.SetFocus
        
        mnuMenu_ = "menucreditnote"
       'SetButton Commandadd, Commandedit, Commandsave, Commanddelete
       
       
      If (AuditTrail = "y") Then
      
      If (txtchecked.text = "y") Then
      
          actionType_ = "Edit"
          vtype1_ = "C"
          vtypeNew = "C"
          vdate_ = Trim(Cbdate.text)
          vno_ = Trim(TCNN.text)
          
          frmAuditTrailLog_Rem.Show 1
          
       End If
      
      End If
       
  
  
Exit Sub

err:

MsgBox "" & err.DESCRIPTION
        
End Sub

Private Sub Commandsearch_Click()

'sqlqry = "select distinct CNN,CND,PGLD,NA from CNF1A where CNN"
'orderby = "order by CNN"
    
searchType = "inv"
popuplist10 "select distinct CNN,CND,PSLD AS PartyName,NA from CNF1A where " & stringyear & "  order by CNN", con
   
    
End Sub
Sub totalAmt()
    
    Text3.text = 0
    For I = 1 To DGrid.rows - 1
       If DGrid.TextMatrix(I, 2) <> "" Then
         Text3.text = Val(Text3.text) + Val(DGrid.TextMatrix(I, 2))
       End If
    Next
    
    
    Text3.text = Format(Text3.text, "0.00")
    
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
     totalAmt
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
                cmbgroup.Visible = False
                DrCrText.Visible = False
                gencombo1.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.top - 15, DGrid.CellWidth
                gencombo1.SetFocus

            Case 1
                Subcombo1.text = DGrid.text
                Subcombo1.Visible = True
                cmbgroup.Visible = False
                gencombo1.Visible = False
                AText.Visible = False
                DrCrText.Visible = False
                If Subcombo1.ListCount > 0 Then
                        'Subcombo1.Text = DGrid.Text
                        'Subcombo1.Visible = True
                    Subcombo1.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.top - 15, DGrid.CellWidth
                    Subcombo1.SetFocus
                End If
    
            Case 2
                AText = DGrid.text
                cmbgroup.Visible = False
                AText.Visible = True
                gencombo1.Visible = False
                Subcombo1.Visible = False
                DrCrText.Visible = False
                AText.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.top - 15, DGrid.CellWidth
                AText.SetFocus
 
            Case 3
                DrCrText.text = "D"
                DrCrText.text = DGrid.text
                If DGrid.text = "" Then DrCrText.text = "D"
                DrCrText.Visible = True
                gencombo1.Visible = False
                cmbgroup.Visible = False
                Subcombo1.Visible = False
                AText.Visible = False
                DrCrText.Move DGrid.CellLeft + DGrid.Left, DGrid.CellTop + DGrid.top - 15, DGrid.CellWidth
                DrCrText.SetFocus
                totalAmt
                
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
End If
End Sub

Private Sub Form_Activate()
Commandadd.Enabled = True
If Commandadd.Visible = True Then
 Commandadd.SetFocus
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
Private Sub Form_Load()

On Error Resume Next

    Me.top = 100
    Me.Left = 100
    
    Me.Width = 16284
    Me.Height = 10500
    

    

   Me.Caption = "Credit Note"
   
   
   Dim rs2 As New ADODB.Recordset
   
   
   
   Me.top = 0
   Me.Left = 0
   Set RS = New ADODB.Recordset
   cmdAdd = False
  
   cmdedit = False
  'ue
   Datachange = False
   
   s5 = ""
   
    If RS.State = 1 Then RS.close
    RS.Open "select Rep as Representative from SalesRepQry where (email is not null and len(email)>1) order by Rep", CON_blue
    cmbAgentName.Clear
    If Not RS.EOF Then
       Do While Not RS.EOF
          If IsNull(RS(0)) = False Then
             Me.cmbAgentName.AddItem RS(0)
             
             If s5 = "" Then
                s5 = RS(0)
             Else
                s5 = s5 & "|" & RS(0)
             End If
             
          End If
          If Not RS.EOF Then RS.MoveNext
        Loop
    End If
   
   vs1.ColComboList(4) = s5
   s5 = ""
   'cboCat.ListIndex = 3
   
   GCombo.Clear
   gencombo1.Clear
   '''''MainMenu.Toolbar1.Visible = True
   If RS.State = 1 Then RS.close
   RS.Open "Select *  from CNF1A where  " & stringyear & " order by cnn  ", con, adOpenDynamic, adLockOptimistic, adCmdText
   If RS.RecordCount > 0 Then
     RS.MoveLast
     
     If inviceNo <> "" Then
     RS.MoveFirst
     RS.Find "cnn=" & inviceNo & ""
     End If
     
     
    If (RS!Checked_YesNo = True) Then
       txtchecked.text = "y"
    Else
        txtchecked.text = "n"
    End If
     
     TCNN.text = RS!cnn
     CdateT.text = RS!Cnd
     Text2.text = RS!n
     GCombo.text = RS!Pgld
     GText.text = RS!Pgld
     addLedger
     Scombo.text = RS!psld
     SText.text = RS!psld
     Text3.text = RS!na
     cmbAgentName.text = RS!agentname & ""
     cboCat = RS!CNCategory
     
     If Not IsNull(RS!todid) Or RS!todid = "" Then
       txtTODNO.text = RS!todid
       txtTODDate.text = RS!toddate
       
     End If
     
     'txtTODNO.Enabled = False
     'txtTODDate.Enabled = False
     
     
     'Set TCNN.DataSource = RS
     'Set CdateT.DataSource = RS
     'Set Text2.DataSource = RS
     'Set Text3.DataSource = RS
     'Set GText.DataSource = RS
     'Set SText.DataSource = RS
     inviceNo = ""
  End If
   
  '--------------------
  
  If TCNN.text <> "" Then
    If rs1.State = 1 Then rs1.close
   rs1.Open "select * from CreditNotDet where cnn=" & TCNN.text & " order by SN", con, adOpenDynamic, adLockOptimistic
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
   Next
   End If
   
   '--------------------
   
    If rs2.State = 1 Then rs2.close
   rs2.Open "Select * from groups where " & stringyear & "order by groupcode", con, adOpenStatic, adLockReadOnly, adCmdText
   If Not rs2.EOF Then
        Do While Not rs2.EOF
            cmbgroup.AddItem rs2!groupcode
            If Not rs2.EOF Then
               rs2.MoveNext
            End If
        Loop
   End If

   
   
   '----------------------
    If rs2.State = 1 Then rs2.close
   rs2.Open "Select * from gledger where   " & stringyear & " and slf = 1  order by gledger", con, adOpenStatic, adLockOptimistic
   
   If Not rs2.EOF Then
        Do While Not rs2.EOF
           GCombo.AddItem rs2(1)
            If Not rs2.EOF Then
                rs2.MoveNext
            End If
        Loop
    End If
   If rs2.State = 1 Then rs2.close
   rs2.Open "Select * from gledger where  " & stringyear & " order by gledger", con, adOpenStatic, adLockOptimistic
   
   If Not rs2.EOF Then
        Do While Not rs2.EOF
            gencombo1.AddItem rs2(1)
            
            If Not rs2.EOF Then
                rs2.MoveNext
            End If
        Loop
    End If
    
    


    
    
    cmdedit = False
    If RS.RecordCount <= 0 Then
        Commandedit.Enabled = False
    End If
    
    
   mnuMenu_ = "menucreditnote"
   SetButton Commandadd, Commandedit, Commandsave, Commanddelete
   Commanddelete.Enabled = False
   Commandsave.Enabled = False
   
   BackColorFrom Me
   
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
'Call GCombo_KeyPress(0)
End Sub

Private Sub GCombo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Dim GEN As String
     Dim SC1 As String
     
     GEN = GCombo.text
     SC1 = Scombo.text
     
     addLedger
     
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
Sub addLedger()
     Dim rs2 As New ADODB.Recordset
     rs2.Open "Select * from sledger where  " & stringyear & " and GLEDGER='" + Trim(GCombo.text) + "'", CCON, adOpenStatic, adLockReadOnly
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

End Sub
Private Sub GCombo_LostFocus()
''  If GCombo.Text = "" Then
''            GCombo.SetFocus
''  Else
''       If GCombo.Text <> "" Then
''            Dim rs4 As New ADODB.Recordset
''            rs4.Open "Select* from gledger  where   " & stringyear & " and slf = 1 and GLEDGER = '" + Trim(GCombo.Text) + "'", CON, adOpenStatic
''            If rs4.RecordCount <= 0 Then
''                 MsgBox "No valid Gen.Ledger"
''                 'GCombo.SetFocus
''            End If
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
       End If
    
       DGrid.text = gencombo1.text
       Subcombo1.Clear
       rs2.Open "Select * from sledger where  " & stringyear & " and GLEDGER='" + Trim(gencombo1.text) + "'", CCON, adOpenForwardOnly, adLockReadOnly, adCmdText
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
  sendkeys "{tab}"
  vs1.Col = 1
  vs1.Row = 1
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
   
'If cmdEdit = True Then
'    cmdEdit = False
   Gridrefresh
'End If


    
End Sub

Private Sub TCNN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   checkCreditnot
End If
End Sub
Sub checkCreditnot()
    Dim rs1 As New ADODB.Recordset
     Dim rs6 As New ADODB.Recordset
     
     If TCNN.text = "" Then Exit Sub
     
     rs6.Open "Select invoiceno from CREDITA  where " & stringyear & " and invoiceno = " + TCNN.text + "", con, adOpenStatic, adLockOptimistic, adCmdText
     If rs6.RecordCount > 0 Then
            MsgBox "Credit Note(Item) already exist..."
            TCNN.SetFocus
            Exit Sub
     End If

     
     rs1.Open "Select top 10 * from  CNF1A  where  " & stringyear & " and cnn = " + TCNN.text + "", con, adOpenStatic, adLockOptimistic, adCmdText
     If Not rs1.EOF Then
     If cmdAdd Then
            MsgBox "Credit Note already exist..."
            TCNN.SetFocus
            Exit Sub
     End If
  
  End If
  
  sendkeys "{tab}"
End Sub

Private Sub TCNN_LostFocus()
 
 checkCreditnot
 
 
 Dim rs6 As New ADODB.Recordset
     If TCNN.text = "" Then
       
        TCNN.SetFocus
       
       Exit Sub
     Else
 '======================
   'If RS.State = 1 Then RS.close
   
   
   
   
   Set RS = New ADODB.Recordset
   RS.Open "Select *  from CNF1A where CNN=" & TCNN.text & " order by cnn  ", con, adOpenDynamic, adLockOptimistic, adCmdText
   If RS.RecordCount > 0 Then
   
   
   If (AuditTrail = "y") Then
   
    If (RS!Checked_YesNo = True) Then
       txtchecked.text = "y"
    Else
        txtchecked.text = "n"
    End If
    
   End If
    
    
    
   
     TCNN.text = RS!cnn
     CdateT.text = RS!Cnd
     Text2.text = RS!n & ""
     GCombo.text = RS!Pgld
     GText.text = RS!Pgld
     addLedger
     Scombo.text = RS!psld
     SText.text = RS!psld
     cmbAgentName.text = RS!agentname & ""
     
     cboCat.text = RS!CNCategory & ""
     
     
     If Not IsNull(RS!todid) Or RS!todid = "" Then
        txtTODNO.text = RS!todid
        txtTODDate.text = RS!toddate
     End If

     
     
   End If
   
  
   If rs1.State = 1 Then rs1.close
   rs1.Open "select * from CreditNotDet where cnn=" & TCNN.text & " order by SN", con, adOpenDynamic, adLockOptimistic
   For I = 1 To rs1.RecordCount
      If RS.EOF = False Then
       vs1.TextMatrix(I, 0) = rs1!sn
       vs1.TextMatrix(I, 1) = rs1!NARR
       vs1.TextMatrix(I, 2) = rs1!paymentamt & ""
       vs1.TextMatrix(I, 3) = rs1!amount
       vs1.TextMatrix(I, 4) = rs1!RepName & ""
       vs1.TextMatrix(I, 5) = rs1!recno & ""
       vs1.TextMatrix(I, 6) = rs1!fyear & ""
       rs1.MoveNext
      End If
   Next
  
 '========================
   End If
     
  
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If DGrid.Row = 0 Then
   DGrid.rows = DGrid.rows + 1
End If

   DGrid.Row = 1
   DGrid.Col = 0
   DGrid.SetFocus
   DGrid.Row = 1
   DGrid.Col = 0
   DGrid_Click
  
End If


KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub
Private Sub Text3_Change()
Text3 = Format(Val(Text3.text), "0.00")
End Sub

Private Sub vs1_GotFocus()
 If PopUpValue1 <> "" Then
    
    vs1.TextMatrix(vs1.RowSel, 2) = PopUpValue3
    vs1.TextMatrix(vs1.RowSel, 5) = PopUpValue1
    vs1.TextMatrix(vs1.RowSel, 6) = Right(popupvalue4, 7)
    
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    popupvalue4 = ""
    
 End If
End Sub
Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)



If KeyCode = 113 Then
   
   searchType = "inv"
   
If cboCat.text = "CD" Then
   
   popuplist10 "select RecNo,Dates,Amount,convert( varchar(30), recno)+''+ fyear as id  " & _
   " from partywiseIssueReceveQry where PartyName='" & Scombo.text & "' and convert( varchar(30), recno)+''+ fyear not in(select convert( varchar(30), recno)+''+ fyear from CreditNotDetQry where (RecNo is not null or RecNo<>'')) order by Dates", con

End If
   
   
End If


End Sub
Private Sub vs1_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = 13 Then
      If vs1.Col = 1 Then
         
         If vs1.TextMatrix(vs1.RowSel, 1) <> "" Then
           vs1.TextMatrix(vs1.RowSel, 1) = UCase(vs1.TextMatrix(vs1.RowSel, 1))
           
            If cboCat.text = "CD" Then
               sendkeys "{right}"
            Else
               sendkeys "{right}"
               sendkeys "{right}"
            End If
         
           
         End If
      
      ElseIf vs1.Col = 2 Then
             If (vs1.TextMatrix(vs1.RowSel, 2) <> "") Then
             sendkeys "{right}"
             End If
          
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
         
      If vs1.Row = 10 Then
         sendkeys "{tab}"
      End If
         
      End If
   End If
End Sub
Private Sub vs1_SelChange()
   
   If vs1.Col = 2 Then
      vs1.Editable = flexEDNone
      
   Else
      vs1.Editable = flexEDKbdMouse
   End If
   
  If vs1.Row = 1 Then
  If vs1.TextMatrix(1, 4) <> "" Then
     cmbAgentName.text = vs1.TextMatrix(1, 4)
  End If
  End If
   
End Sub
