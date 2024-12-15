VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmGodown 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Godown Master"
   ClientHeight    =   8628
   ClientLeft      =   2376
   ClientTop       =   2400
   ClientWidth     =   16212
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8628
   ScaleWidth      =   16212
   Begin VB.ComboBox cboBinder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Left            =   3720
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   2052
      Width           =   3885
   End
   Begin VB.OptionButton Option1_both 
      BackColor       =   &H0078CFE9&
      Caption         =   "Both"
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
      Left            =   4440
      TabIndex        =   2
      Top             =   288
      Width           =   1695
   End
   Begin VB.OptionButton Option2_Printer 
      BackColor       =   &H0078CFE9&
      Caption         =   "Printer"
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
      Left            =   2580
      TabIndex        =   1
      Top             =   288
      Width           =   1695
   End
   Begin VB.OptionButton Option1_Binder 
      BackColor       =   &H0078CFE9&
      Caption         =   "Binder"
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
      Left            =   960
      TabIndex        =   0
      Top             =   288
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   960
      MaxLength       =   200
      TabIndex        =   4
      Top             =   1548
      Width           =   6645
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFD7AE&
      Caption         =   "Contact Person"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2805
      TabIndex        =   47
      Top             =   9612
      Visible         =   0   'False
      Width           =   6135
      Begin VB.TextBox txtContact1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   225
         MaxLength       =   50
         TabIndex        =   16
         Top             =   645
         Width           =   3180
      End
      Begin VB.TextBox txtMobile1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3465
         MaxLength       =   50
         TabIndex        =   17
         Top             =   630
         Width           =   2415
      End
      Begin VB.TextBox txtContact5 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   225
         MaxLength       =   50
         TabIndex        =   24
         Top             =   1950
         Width           =   3180
      End
      Begin VB.TextBox txtContact4 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   225
         MaxLength       =   50
         TabIndex        =   22
         Top             =   1605
         Width           =   3180
      End
      Begin VB.TextBox txtContact3 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   225
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1290
         Width           =   3180
      End
      Begin VB.TextBox txtMobile2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3465
         MaxLength       =   50
         TabIndex        =   19
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtContact2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   225
         MaxLength       =   50
         TabIndex        =   18
         Top             =   960
         Width           =   3180
      End
      Begin VB.TextBox txtMobile3 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3465
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1290
         Width           =   2415
      End
      Begin VB.TextBox txtMobile4 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3465
         MaxLength       =   50
         TabIndex        =   23
         Top             =   1620
         Width           =   2415
      End
      Begin VB.TextBox txtMobile5 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3465
         MaxLength       =   50
         TabIndex        =   25
         Top             =   1950
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   2
         Left            =   225
         TabIndex        =   49
         Top             =   375
         Width           =   660
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   3465
         TabIndex        =   48
         Top             =   390
         Width           =   735
      End
   End
   Begin VB.CheckBox Godown 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Godown"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10728
      TabIndex        =   15
      Top             =   3384
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox desc 
      ForeColor       =   &H00C00000&
      Height          =   288
      Left            =   12612
      TabIndex        =   28
      Top             =   8856
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox ob 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   12696
      TabIndex        =   29
      Top             =   8856
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox cid 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   7635
      MaxLength       =   10
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox na 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   960
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1068
      Width           =   6645
   End
   Begin VB.TextBox mobile 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2790
      MaxLength       =   100
      TabIndex        =   13
      Top             =   8916
      Visible         =   0   'False
      Width           =   6165
   End
   Begin VB.TextBox city 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2970
      MaxLength       =   100
      TabIndex        =   5
      Top             =   8796
      Visible         =   0   'False
      Width           =   6165
   End
   Begin VB.TextBox txtwestage 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   960
      MaxLength       =   4
      TabIndex        =   6
      Top             =   2088
      Width           =   645
   End
   Begin VB.TextBox pno2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2970
      MaxLength       =   100
      TabIndex        =   11
      Top             =   8688
      Visible         =   0   'False
      Width           =   6165
   End
   Begin VB.TextBox Faxno 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   2970
      MaxLength       =   100
      TabIndex        =   12
      Top             =   9012
      Visible         =   0   'False
      Width           =   6165
   End
   Begin VB.TextBox emailid 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   1020
      MaxLength       =   100
      TabIndex        =   14
      Top             =   10080
      Visible         =   0   'False
      Width           =   4830
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   945
      ScaleHeight     =   876
      ScaleWidth      =   5088
      TabIndex        =   31
      Top             =   2772
      Width           =   5085
      Begin VB.CommandButton save 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   720
         Left            =   1260
         Picture         =   "frmGodown.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton Del 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   720
         Left            =   2535
         Picture         =   "frmGodown.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton Abandon 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   720
         Left            =   45
         Picture         =   "frmGodown.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   75
         Width           =   1185
      End
      Begin VB.CommandButton REPORTCD 
         Caption         =   "&Print"
         Height          =   585
         Left            =   4860
         TabIndex        =   33
         Top             =   1020
         Width           =   1245
      End
      Begin VB.CommandButton close 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         Height          =   720
         Left            =   3780
         Picture         =   "frmGodown.frx":23AC
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   75
         Width           =   1245
      End
      Begin VB.CommandButton Help 
         Caption         =   "&Help"
         Height          =   450
         Left            =   240
         TabIndex        =   34
         Top             =   150
         Visible         =   0   'False
         Width           =   15
      End
   End
   Begin VB.ListBox List1 
      Height          =   240
      Left            =   12684
      Sorted          =   -1  'True
      TabIndex        =   30
      Top             =   8868
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox add2 
      Appearance      =   0  'Flat
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   12984
      MaxLength       =   100
      TabIndex        =   27
      Top             =   8868
      Visible         =   0   'False
      Width           =   300
   End
   Begin Crystal.CrystalReport Cr1 
      Left            =   300
      Top             =   3288
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VSFlex7Ctl.VSFlexGrid vs1 
      Height          =   4740
      Left            =   0
      TabIndex        =   53
      Top             =   3780
      Width           =   16092
      _cx             =   28384
      _cy             =   8361
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   11162880
      BackColorFixed  =   7917545
      ForeColorFixed  =   8388608
      BackColorSel    =   16777153
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   330
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmGodown.frx":2F90
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
   Begin VB.Label lblLink 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Link To"
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
      Left            =   2700
      TabIndex        =   52
      Top             =   2124
      Width           =   624
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Westage :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   180
      TabIndex        =   51
      Top             =   2148
      Width           =   840
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   180
      TabIndex        =   50
      Top             =   1596
      Width           =   1740
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0078CFE9&
      BorderWidth     =   4
      Height          =   960
      Left            =   900
      Top             =   2724
      Width           =   5148
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Balance"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   12444
      TabIndex        =   46
      Top             =   8832
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ledger Description"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   12456
      TabIndex        =   45
      Top             =   8832
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press F2 For Search "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   192
      Left            =   960
      TabIndex        =   44
      Top             =   828
      Width           =   1800
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1188
      TabIndex        =   43
      Top             =   8952
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   " Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   228
      Left            =   132
      TabIndex        =   42
      Top             =   1128
      Width           =   1740
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "City   "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1368
      TabIndex        =   41
      Top             =   8796
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Resi Phone Nos. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   228
      Left            =   1368
      TabIndex        =   40
      Top             =   8748
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax Nos. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   228
      Left            =   1368
      TabIndex        =   39
      Top             =   9072
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Office Phone Nos. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1188
      TabIndex        =   38
      Top             =   8736
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Email ID "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   1188
      TabIndex        =   37
      Top             =   9252
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Address2 :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   228
      Index           =   0
      Left            =   12912
      TabIndex        =   36
      Top             =   8652
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Label LABELUSERNAME 
      BackStyle       =   0  'Transparent
      Height          =   336
      Left            =   2436
      TabIndex        =   35
      Top             =   6864
      Width           =   6888
   End
End
Attribute VB_Name = "frmGodown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ref As ADODB.Recordset
Dim flag As Boolean
Dim value As String
Dim RS As New ADODB.Recordset

Sub COMPINI()

na.text = ""
add1.text = ""

MaxNo
End Sub
Sub fillGrig()

Dim vs_rs As New ADODB.Recordset

Set vs_rs = New ADODB.Recordset

If vs_rs.State = 1 Then vs_rs.close

If module_ = "Invoicing" Then

vs_rs.Open "select Godwn as [Binder And Printer],Address,Id from Godownmaster where len(Godwn)<5 order by Godwn", con
Set vs1.DataSource = vs_rs

Else

vs_rs.Open "select Godwn as [Binder And Printer],LinkTo,Address,GSTIN,ContactName,ContactNo,Id from Godownmaster where len(Godwn)>10 order by Godwn", con
Set vs1.DataSource = vs_rs


End If


If module_ = "Invoicing" Then

Me.Width = 11000

vs1.Cols = 3

vs1.FormatString = "Binder And Printer|Address|Id"

vs1.ColWidth(0) = 3200
vs1.ColWidth(1) = 3200
vs1.ColWidth(2) = 2400
'vs1.ColWidth(3) = 500


vs1.Width = 10000

Else

vs1.Cols = 7
vs1.FormatString = "Binder And Printer|LinkTo|Address|GSTIN|Contact Person|Mobile|ID"

vs1.ColWidth(0) = 3000
vs1.ColWidth(1) = 3000
vs1.ColWidth(2) = 4700

vs1.ColWidth(3) = 1500
vs1.ColWidth(4) = 1500
vs1.ColWidth(5) = 1500
vs1.ColWidth(6) = 400
End If


End Sub
Private Sub ABANDON_Click()

na.text = ""
txtAddress.text = ""
txtwestage = ""
cboBinder.text = ""
MaxNo
fillGrig
na.SetFocus
End Sub

Private Sub close_Click()
Unload Me
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   sendkeys "{tab}"
'End If
End Sub



Private Sub Del_Click()

X = MsgBox("Are you sure you wish to delete the selected item ", 4, "Confirmation")
If X = 6 Then
   
   con.Execute "Delete  from Godownmaster where id= " & cid.text & " and " & stringyear & ""
   na.text = ""
   txtAddress = ""
   txtwestage = ""
   na.SetFocus
End If

End Sub
Sub MaxNo()
   
   If RS.State = 1 Then RS.close
   RS.Open "select max(id) from Godownmaster", con
   If IsNull(RS.Fields(0).value) Then
      cid.text = 1
   Else
      cid.text = RS.Fields(0).value + 1
   End If
   
End Sub
Private Sub Form_Load()

Me.top = 500
Me.Left = 500

MaxNo


If module_ = "Paper" Then
   Me.Caption = "Printer/Binder..."
   txtwestage.Visible = True
   Label11.Visible = True
   Me.Option1_Binder.Visible = True
   Me.Option2_Printer.Visible = True
   Option1_both.Visible = True
   cboBinder.Visible = True
   lblLink.Visible = True
Else
   Me.Option1_Binder.value = False
   Me.Option2_Printer.value = False
   
   Me.Option1_Binder.Visible = False
   Me.Option2_Printer.Visible = False
   Option1_both.Visible = False
   
   cboBinder.Visible = False
   lblLink.Visible = False
   
   txtwestage.Visible = False
   Label11.Visible = False
   

End If

If module_ = "Invoicing" Then

    If RS.State = 1 Then RS.close
    RS.Open "select Godwn from Godownmaster where len(Godwn)<10", con
    While RS.EOF = False
     cboBinder.AddItem (RS(0))
     RS.MoveNext
    Wend

Else

    If RS.State = 1 Then RS.close
    RS.Open "select Godwn from Godownmaster where len(Godwn)>=10", con
    While RS.EOF = False
     cboBinder.AddItem (RS(0))
     RS.MoveNext
    Wend


End If



fillGrig

BackColorFrom Me

End Sub

Private Sub Na_GotFocus()



If PopUpValue1 <> "" Then

    na.text = PopUpValue1
    cid = PopUpValue3
    txtAddress = PopUpValue2
     
    If popupvalue4 = "b" Then
       Option1_Binder.value = True
    ElseIf popupvalue4 = "p" Then
       Option2_Printer.value = True
    ElseIf popupvalue4 = "pb" Then
       Option1_both.value = True
       
    End If
    
    If Not IsNull(popupvalue5) Then
       txtwestage.text = popupvalue5
    End If
    
    If rs1.State = 1 Then RS.close
    rs1.Open "select linkto from Godownmaster where id='" & cid & "'", con
    If rs1.EOF Then
       cboBinder.text = RS(0) & ""
    End If
    
    
    
    PopUpValue1 = ""
    PopUpValue2 = ""
    PopUpValue3 = ""
    popupvalue4 = ""
    popupvalue5 = ""

End If


End Sub

Private Sub na_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
  txtAddress.SetFocus
End If



If KeyCode = 113 Then
   
If module_ = "Paper" Then
   'If Option1_Binder.value = True Then
      value = "Select Godwn as [Printer/Binder],Address,Id,binder_printer,westage from Godownmaster where " & stringyear & " and (binder_printer='b' or binder_printer='p' or binder_printer='pb')  order by id"
      'value = "Select Godwn as [Printer/Binder],Address,Id,binder_printer,westage from Godownmaster where " & stringyear & " and binder_printer='b' and westage>0 order by id"
   'Else
   '   value = "Select Godwn as [Printer/Binder],Address,Id,binder_printer,westage from Godownmaster where " & stringyear & " and binder_printer='p' order by id"
   'End If
   popuplistModel10 value, con
ElseIf module_ = "Stock System" Then
   '   If Option1_Binder.value = True Then
   '   value = "Select Godwn as [Printer/Binder],Address,Id,binder_printer,westage from Godownmaster where " & stringyear & " and binder_printer='b' order by id"
   'Else
      value = "Select Godwn as [Printer/Binder],Address,Id,binder_printer,westage from Godownmaster where " & stringyear & " and (binder_printer='b' or binder_printer='pb') order by id"
   'End If
   popuplistModel10 value, con

Else
   value = "Select Godwn as [Godown Name],Address,Id,binder_printer from Godownmaster where " & stringyear & " and binder_printer='g' order by id"
   popuplistModel10 value, con

End If


End If

End Sub

Private Sub na_LostFocus()
na.text = UCase(na)
End Sub

Private Sub save_Click()

Dim voucher As Boolean

Set RS = New ADODB.Recordset

If na = "" Then
   MsgBox "Enter Godown Master ...", vbCritical
   Exit Sub
End If

If RS.State = 1 Then RS.close
RS.Open "select * from Godownmaster where id='" & Trim(cid.text) & "' and " & stringyear & "", con, adOpenDynamic, adLockOptimistic

If RS.EOF = True Then
        
   RS.AddNew
   RS!id = cid
   RS!Godwn = Trim(na.text)
   RS!fyear = session
   RS!setupid = setupid
   RS!Address = Trim(txtAddress.text)
   RS!linkto = Trim(cboBinder.text)
   
   If Option1_Binder.value = True Then
      RS!binder_printer = "b"
   ElseIf Option2_Printer.value = True Then
      RS!binder_printer = "b"
   ElseIf Option1_both.value = True Then
      RS!binder_printer = "pb"
   Else
      RS!binder_printer = "g"
   End If
   RS!westage = Val(txtwestage.text)
   RS.update
    
   MsgBox " Record Saved "
   na.SetFocus
Else
   RS!linkto = Trim(cboBinder.text)
   RS!Godwn = (na.text)
   RS!Address = Trim(txtAddress.text)
   If Option1_Binder.value = True Then
      RS!binder_printer = "b"
   ElseIf Option2_Printer.value = True Then
      RS!binder_printer = "p"
   ElseIf Option1_both.value = True Then
      RS!binder_printer = "pb"
   Else
      RS!binder_printer = "g"
   End If
   RS!westage = Val(txtwestage.text)
   RS.update
End If


End Sub


Private Sub txtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  If module_ = "Paper" Then
     txtwestage.SetFocus
  End If
End If
End Sub

Private Sub txtAddress_LostFocus()
txtAddress = UCase(txtAddress)
End Sub

Private Sub txtwestage_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  cboBinder.SetFocus
End If
End Sub

Private Sub vs1_DblClick()

If rs1.State = 1 Then rs1.close

If module_ = "Invoicing" Then
rs1.Open "Select Godwn as [Printer/Binder],Address,Id,binder_printer from Godownmaster where id='" & vs1.TextMatrix(vs1.RowSel, 2) & "'", con
Else
rs1.Open "Select Godwn as [Printer/Binder],Address,Id,binder_printer,westage,linkto from Godownmaster where id='" & vs1.TextMatrix(vs1.RowSel, 6) & "'", con

End If

If rs1.EOF = False Then

    na.text = rs1(0)
    txtAddress = rs1(1)
    cid = rs1(2)
         
    If rs1!binder_printer = "b" Then
       Option1_Binder.value = True
    ElseIf rs1!binder_printer = "p" Then
       Option2_Printer.value = True
    ElseIf rs1!binder_printer = "pb" Then
       Option1_both.value = True
       
    End If
        
        
   If module_ <> "Invoicing" Then
   
    If Not IsNull(rs1!westage) Then
       txtwestage.text = rs1!westage
    End If

    cboBinder.text = ""
    If Not IsNull(rs1!linkto) Then
       cboBinder.text = rs1!linkto
    End If
    
    End If

End If



End Sub

Private Sub vs1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then
'
'   If vs1.Col = 3 Then
'      con.Execute "update Godownmaster set GSTIN='" & vs1.TextMatrix(vs1.RowSel, 3) & "' where id='" & vs1.TextMatrix(vs1.RowSel, 6) & "'"
'   End If
'
'    If vs1.Col = 4 Then
'      con.Execute "update Godownmaster set ContactName='" & vs1.TextMatrix(vs1.RowSel, 4) & "' where id='" & vs1.TextMatrix(vs1.RowSel, 6) & "'"
'   End If
'
'   If vs1.Col = 5 Then
'      con.Execute "update Godownmaster set ContactNo='" & vs1.TextMatrix(vs1.RowSel, 5) & "' where id='" & vs1.TextMatrix(vs1.RowSel, 6) & "'"
'   End If
'
'   sendkeys "{right}"
'
'End If
End Sub

Private Sub vs1_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then

   If vs1.Col = 3 Then
      con.Execute "update Godownmaster set GSTIN='" & vs1.TextMatrix(vs1.RowSel, 3) & "' where id='" & vs1.TextMatrix(vs1.RowSel, 6) & "'"
      sendkeys "{right}"
   End If

    If vs1.Col = 4 Then
      con.Execute "update Godownmaster set ContactName='" & vs1.TextMatrix(vs1.RowSel, 4) & "' where id='" & vs1.TextMatrix(vs1.RowSel, 6) & "'"
      sendkeys "{right}"
   End If

   If vs1.Col = 5 Then
      con.Execute "update Godownmaster set ContactNo='" & vs1.TextMatrix(vs1.RowSel, 5) & "' where id='" & vs1.TextMatrix(vs1.RowSel, 6) & "'"
      sendkeys "{home}"
      sendkeys "{down}"
      sendkeys "{right}"
      sendkeys "{right}"
      sendkeys "{right}"
   End If

   

End If

End Sub
Private Sub vs1_SelChange()
  
  If vs1.Col <= 2 Then
      vs1.Editable = flexEDNone
   Else
      vs1.Editable = flexEDKbdMouse
   End If
End Sub
